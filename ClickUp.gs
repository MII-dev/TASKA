// ==========================================
// ====== ClickUp API Integration ===========
// ==========================================

const CLICKUP_BASE_URL = 'https://api.clickup.com/api/v2';
const CLICKUP_DESC_PREVIEW_LENGTH = 200;
const CLICKUP_MAX_PAGES = 20;
const CLICKUP_TOKEN_PROPERTY = 'CLICKUP_TOKEN';

/**
 * Reads the stored ClickUp token.
 * Kept server-side so AI functions can reach ClickUp without the client
 * having to hand them a token on every call.
 */
function getClickUpToken() {
  const token = PropertiesService.getUserProperties().getProperty(CLICKUP_TOKEN_PROPERTY);
  return (token && token.trim()) ? token.trim() : null;
}

/**
 * Stores the ClickUp token after verifying it works.
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} {hasToken, user}
 */
function setClickUpToken(token) {
  if (!token || token.trim() === '') {
    throw new Error('ClickUp API токен не надано.');
  }

  // Validate before storing — /user is the cheapest authenticated endpoint
  const userData = clickUpFetch('/user', token.trim());

  PropertiesService.getUserProperties().setProperty(CLICKUP_TOKEN_PROPERTY, token.trim());
  invalidateClickUpCache();
  invalidateAiContextCache();

  return { hasToken: true, user: mapClickUpUser(userData.user) };
}

/**
 * Forgets the stored ClickUp token.
 */
function clearClickUpTokenServer() {
  PropertiesService.getUserProperties().deleteProperty(CLICKUP_TOKEN_PROPERTY);
  invalidateClickUpCache();
  invalidateAiContextCache();
  return { hasToken: false };
}

/**
 * Whether ClickUp is connected — lets the client render the right view
 * without ever holding the token itself.
 */
function getClickUpStatus() {
  return { hasToken: !!getClickUpToken() };
}

/**
 * Resolves the token to use: an explicitly passed one wins, otherwise the stored one.
 */
function resolveClickUpToken(token) {
  const resolved = (token && token.trim()) ? token.trim() : getClickUpToken();
  if (!resolved) {
    throw new Error('ClickUp не підключено. Введіть Personal API Token у вкладці ClickUp.');
  }
  return resolved;
}

function invalidateClickUpCache() {
  try {
    CacheService.getUserCache().removeAll(['cu_tasks_open', 'cu_tasks_all', 'cu_awaiting_reply']);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути ClickUp: ' + e.message);
  }
}

/**
 * Low-level request to the ClickUp API v2.
 * @param {string} endpoint - API endpoint path (e.g. "/user", "/task/abc123")
 * @param {string} token - ClickUp Personal API Token
 * @param {string} method - HTTP method ('get' by default)
 * @param {Object} payload - Optional JSON body for post/put requests
 * @returns {Object} Parsed JSON response
 */
function clickUpFetch(endpoint, token, method, payload) {
  if (!token || token.trim() === '') {
    throw new Error('ClickUp API токен не надано.');
  }

  const options = {
    method: method || 'get',
    headers: {
      'Authorization': token.trim(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(CLICKUP_BASE_URL + endpoint, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code === 401) {
    throw new Error('Невалідний ClickUp API токен. Перевірте токен і спробуйте ще раз.');
  }

  if (code < 200 || code >= 300) {
    throw new Error('ClickUp API помилка (HTTP ' + code + '): ' + text);
  }

  return text ? JSON.parse(text) : {};
}

/**
 * Makes an HTTP GET request to the ClickUp API v2.
 */
function clickUpApiRequest(endpoint, token) {
  return clickUpFetch(endpoint, token, 'get');
}

/**
 * Builds the "Space → Folder → List" breadcrumb for a task.
 */
function buildClickUpPath(task) {
  const listName = (task.list || {}).name || '';
  const folderName = (task.folder || {}).name || '';
  const spaceName = (task.space || {}).name || '';
  // If folder is hidden (same as list), don't duplicate
  const folderHidden = (task.folder || {}).hidden || false;

  let path = spaceName;
  if (folderName && !folderHidden) {
    path += ' → ' + folderName;
  }
  if (listName) {
    path += ' → ' + listName;
  }
  return path;
}

/**
 * Normalizes a ClickUp user object into the shape the UI expects.
 */
function mapClickUpUser(user) {
  if (!user) return null;
  const name = user.username || user.email || '';
  const initials = user.initials || name.split(' ').map(function (p) {
    return p.charAt(0).toUpperCase();
  }).join('').substring(0, 2);

  return {
    id: user.id,
    name: name,
    initials: initials,
    color: user.color || '#7c4dff'
  };
}

/**
 * Normalizes a raw ClickUp task into the shape the UI expects.
 * The description is truncated here — the full text is served by getClickUpTaskDetails().
 */
function mapClickUpTask(task, teamName) {
  const statusObj = task.status || {};
  const fullDesc = task.text_content || task.description || '';

  return {
    id: task.id,
    name: task.name,
    description: fullDesc.substring(0, CLICKUP_DESC_PREVIEW_LENGTH),
    hasMoreDescription: fullDesc.length > CLICKUP_DESC_PREVIEW_LENGTH,
    status: {
      name: statusObj.status || 'unknown',
      color: statusObj.color || '#808080',
      type: statusObj.type || ''
    },
    priority: (task.priority && task.priority.id) ? parseInt(task.priority.id, 10) : null,
    dueDate: task.due_date ? parseInt(task.due_date, 10) : null,
    startDate: task.start_date ? parseInt(task.start_date, 10) : null,
    timeEstimate: task.time_estimate ? parseInt(task.time_estimate, 10) : null,
    timeSpent: task.time_spent ? parseInt(task.time_spent, 10) : null,
    parent: task.parent || null,
    listId: (task.list || {}).id || null,
    listName: (task.list || {}).name || '',
    url: task.url || ('https://app.clickup.com/t/' + task.id),
    path: buildClickUpPath(task),
    workspace: teamName || '',
    tags: (task.tags || []).map(function (tag) { return tag.name; }),
    creator: mapClickUpUser(task.creator),
    assignees: (task.assignees || []).map(mapClickUpUser).filter(Boolean)
  };
}

const CLICKUP_STATUS_CACHE_SECONDS = 21600; // 6 h — list statuses change rarely
const CLICKUP_FETCHALL_CHUNK = 40;

/**
 * Normalizes the statuses array of a ClickUp list resource.
 */
function mapClickUpStatuses(listResource) {
  return ((listResource || {}).statuses || []).map(function (s) {
    return {
      name: s.status,
      color: s.color || '#808080',
      type: s.type || '',
      orderindex: s.orderindex
    };
  });
}

/**
 * Fetches the status sets of several lists at once, so the UI can offer a status
 * dropdown per task. Cached results are reused; the rest are fetched in parallel.
 * @param {Array} listIds - Unique ClickUp list IDs
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} Map of listId -> array of {name, color, type, orderindex}
 */
function fetchClickUpStatusesForLists(listIds, token) {
  const result = {};
  if (!listIds || listIds.length === 0) return result;

  const cache = CacheService.getUserCache();
  const cacheKeys = listIds.map(function (id) { return 'cu_statuses_' + id; });
  const cached = cache.getAll(cacheKeys) || {};

  const missing = [];
  listIds.forEach(function (id) {
    const raw = cached['cu_statuses_' + id];
    if (raw) {
      try {
        result[id] = JSON.parse(raw);
        return;
      } catch (e) {
        // Corrupted cache entry — just refetch it
      }
    }
    missing.push(id);
  });

  if (missing.length === 0) return result;

  const toCache = {};

  // Fetch the cache misses in parallel batches instead of one request at a time
  for (var offset = 0; offset < missing.length; offset += CLICKUP_FETCHALL_CHUNK) {
    var chunk = missing.slice(offset, offset + CLICKUP_FETCHALL_CHUNK);

    var requests = chunk.map(function (listId) {
      return {
        url: CLICKUP_BASE_URL + '/list/' + listId,
        method: 'get',
        headers: {
          'Authorization': token.trim(),
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
    });

    var responses;
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      Logger.log('Не вдалося завантажити статуси списків: ' + e.message);
      continue;
    }

    responses.forEach(function (response, i) {
      var listId = chunk[i];
      if (response.getResponseCode() !== 200) {
        Logger.log('Статуси списку ' + listId + ' недоступні (HTTP ' + response.getResponseCode() + ')');
        return;
      }
      try {
        var statuses = mapClickUpStatuses(JSON.parse(response.getContentText()));
        result[listId] = statuses;
        toCache['cu_statuses_' + listId] = JSON.stringify(statuses);
      } catch (e) {
        Logger.log('Не вдалося розібрати статуси списку ' + listId + ': ' + e.message);
      }
    });
  }

  if (Object.keys(toCache).length > 0) {
    cache.putAll(toCache, CLICKUP_STATUS_CACHE_SECONDS);
  }

  return result;
}

/**
 * Fetches the set of statuses available for a single list.
 * @param {string} listId - The ClickUp list ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Array} Array of {name, color, type, orderindex}
 */
function getClickUpListStatuses(listId, token) {
  const map = fetchClickUpStatusesForLists([listId], token);
  return map[listId] || [];
}

/**
 * Fetches tasks assigned to the current user from all ClickUp workspaces,
 * without the per-list status sets. Walks every page — ClickUp returns 100 per page.
 *
 * Split out from getClickUpTasks so the AI context can reuse the task list
 * without paying for the status lookups it does not need.
 *
 * @param {string} token - ClickUp Personal API Token
 * @param {boolean} includeClosed - Whether to include closed/done tasks
 * @returns {Object} {tasks, truncated, currentUser}
 */
function fetchClickUpTasksRaw(token, includeClosed) {
  const cacheKey = includeClosed ? 'cu_tasks_all' : 'cu_tasks_open';
  const cached = cacheGetJson(cacheKey);
  if (cached) return cached;

  // Step 1: Get current user info
  const userData = clickUpFetch('/user', token);
  const userId = userData.user.id;
  const currentUser = mapClickUpUser(userData.user);

  // Step 2: Get all workspaces (teams)
  const teamsData = clickUpFetch('/team', token);
  const teams = teamsData.teams || [];

  if (teams.length === 0) {
    return { tasks: [], truncated: false, currentUser: currentUser };
  }

  // Step 3: Fetch every page of tasks assigned to this user from each workspace
  var allTasks = [];
  var truncated = false;

  teams.forEach(function (team) {
    try {
      var page = 0;
      var lastPage = false;

      while (!lastPage && page < CLICKUP_MAX_PAGES) {
        var endpoint = '/team/' + team.id + '/task'
          + '?assignees[]=' + userId
          + '&subtasks=true'
          + '&include_closed=' + (includeClosed ? 'true' : 'false')
          + '&order_by=due_date'
          + '&reverse=false'
          + '&page=' + page;

        var tasksData = clickUpFetch(endpoint, token);
        var tasks = tasksData.tasks || [];

        tasks.forEach(function (task) {
          allTasks.push(mapClickUpTask(task, team.name));
        });

        lastPage = (tasksData.last_page === true) || tasks.length === 0;
        page++;
      }

      // Hit the page cap with more results still waiting — tell the UI instead of silently dropping them
      if (!lastPage) {
        truncated = true;
        Logger.log('Досягнуто ліміт сторінок для workspace ' + team.name);
      }
    } catch (e) {
      Logger.log('Помилка завантаження задач з workspace ' + team.name + ': ' + e.message);
    }
  });

  // Sort: tasks with due dates first (ascending), then tasks without
  allTasks.sort(function (a, b) {
    if (a.dueDate && b.dueDate) return a.dueDate - b.dueDate;
    if (a.dueDate) return -1;
    if (b.dueDate) return 1;
    return 0;
  });

  const result = { tasks: allTasks, truncated: truncated, currentUser: currentUser };
  cachePutJson(cacheKey, result, CACHE_TTL_CLICKUP_TASKS);
  return result;
}

/**
 * Fetches tasks assigned to the current user, plus the status set of every list
 * involved so the UI can offer an inline status dropdown.
 * @param {string} token - ClickUp Personal API Token (optional if stored)
 * @param {boolean} includeClosed - Whether to include closed/done tasks
 * @returns {Object} {tasks, statusesByList, truncated, currentUser}
 */
function getClickUpTasks(token, includeClosed) {
  const authToken = resolveClickUpToken(token);
  const raw = fetchClickUpTasksRaw(authToken, includeClosed);

  var listIdSet = {};
  raw.tasks.forEach(function (t) {
    if (t.listId) listIdSet[t.listId] = true;
  });

  var statusesByList = {};
  try {
    statusesByList = fetchClickUpStatusesForLists(Object.keys(listIdSet), authToken);
  } catch (e) {
    // Without statuses the UI falls back to read-only badges — not worth failing the whole load
    Logger.log('Не вдалося завантажити статуси списків: ' + e.message);
  }

  return {
    tasks: raw.tasks,
    statusesByList: statusesByList,
    truncated: raw.truncated,
    currentUser: raw.currentUser
  };
}

/**
 * Formats ClickUp tasks for the AI context. Returns an empty string when
 * ClickUp is not connected, so the section simply does not appear.
 */
function getClickUpContext() {
  const token = getClickUpToken();
  if (!token) return '';

  try {
    const raw = fetchClickUpTasksRaw(token, false);
    if (!raw.tasks || raw.tasks.length === 0) {
      return 'CLICKUP: активних задач немає.\n';
    }

    const shown = raw.tasks.slice(0, AI_CTX_MAX_CLICKUP_TASKS);
    let context = `ЗАДАЧІ З CLICKUP (${shown.length} з ${raw.tasks.length}, від найближчого дедлайну):\n`;

    const priorityNames = { 1: 'Urgent', 2: 'High', 3: 'Normal', 4: 'Low' };

    shown.forEach(function (t, index) {
      context += `${index + 1}. [${t.status.name}] "${t.name}" | ClickUp ID: ${t.id}\n`;

      const attrs = [];
      if (t.dueDate) attrs.push('дедлайн: ' + formatDateShort(new Date(t.dueDate).toISOString()));
      if (t.priority && priorityNames[t.priority]) attrs.push('пріоритет: ' + priorityNames[t.priority]);
      if (t.path) attrs.push('список: ' + t.path);
      if (t.creator && t.creator.name) attrs.push('створив: ' + t.creator.name);
      if (attrs.length > 0) context += '   ' + attrs.join(' | ') + '\n';
    });

    if (raw.tasks.length > shown.length) {
      context += `…та ще ${raw.tasks.length - shown.length} задач у ClickUp.\n`;
    }

    return context;
  } catch (e) {
    Logger.log('Не вдалося отримати контекст ClickUp: ' + e.message);
    return '';
  }
}

/**
 * Fetches full details for a single task — full description, assignees,
 * time tracking, checklists and subtasks.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} Detailed task object
 */
function getClickUpTaskDetails(taskId, token) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const authToken = resolveClickUpToken(token);
  const task = clickUpFetch('/task/' + taskId + '?include_subtasks=true', authToken);
  const details = mapClickUpTask(task, '');

  // Full, untruncated description
  details.description = task.text_content || task.description || '';
  details.hasMoreDescription = false;

  details.dateCreated = task.date_created ? parseInt(task.date_created, 10) : null;
  details.dateUpdated = task.date_updated ? parseInt(task.date_updated, 10) : null;

  details.checklists = (task.checklists || []).map(function (cl) {
    return {
      name: cl.name || '',
      items: (cl.items || []).map(function (item) {
        return { name: item.name || '', resolved: !!item.resolved };
      })
    };
  });

  details.subtasks = (task.subtasks || []).map(function (st) {
    const stStatus = st.status || {};
    return {
      id: st.id,
      name: st.name,
      status: {
        name: stStatus.status || '',
        color: stStatus.color || '#808080',
        type: stStatus.type || ''
      }
    };
  });

  return details;
}

/**
 * Updates a task's status, priority and/or due date.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @param {Object} updates - {status?: string, priority?: number|null, dueDate?: number|null}
 * @returns {Object} The updated field values as ClickUp stored them
 */
function updateClickUpTask(taskId, token, updates) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }
  if (!updates) {
    throw new Error('Немає полів для оновлення.');
  }

  const payload = {};

  if (updates.status) {
    payload.status = updates.status;
  }

  // priority: 1-4, or null to clear it
  if (updates.hasOwnProperty('priority')) {
    payload.priority = (updates.priority === null || updates.priority === '')
      ? null
      : parseInt(updates.priority, 10);
  }

  // dueDate: unix milliseconds, or null to clear it
  if (updates.hasOwnProperty('dueDate')) {
    payload.due_date = (updates.dueDate === null || updates.dueDate === '')
      ? null
      : parseInt(updates.dueDate, 10);
    payload.due_date_time = false;
  }

  if (Object.keys(payload).length === 0) {
    throw new Error('Немає полів для оновлення.');
  }

  const task = clickUpFetch('/task/' + taskId, resolveClickUpToken(token), 'put', payload);
  const statusObj = task.status || {};

  // The cached task list and AI context both just went stale
  invalidateClickUpCache();
  invalidateAiContextCache();

  return {
    id: task.id,
    status: {
      name: statusObj.status || '',
      color: statusObj.color || '#808080',
      type: statusObj.type || ''
    },
    priority: (task.priority && task.priority.id) ? parseInt(task.priority.id, 10) : null,
    dueDate: task.due_date ? parseInt(task.due_date, 10) : null
  };
}

/**
 * Makes an HTTP POST request to the ClickUp API v2.
 * @param {string} endpoint - API endpoint path (e.g. "/task/{id}/comment")
 * @param {string} token - ClickUp Personal API Token
 * @param {Object} payload - JSON payload to send
 * @returns {Object} Parsed JSON response
 */
function clickUpApiPost(endpoint, token, payload) {
  return clickUpFetch(endpoint, token, 'post', payload);
}

/**
 * Fetches comments for a specific ClickUp task.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Array} Array of formatted comment objects
 */
function getClickUpTaskComments(taskId, token) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const data = clickUpApiRequest('/task/' + taskId + '/comment', resolveClickUpToken(token));
  var comments = data.comments || [];
  
  return comments.map(function(comment) {
    // Extract plain text from comment_text array
    var textParts = [];
    if (comment.comment_text) {
      textParts.push(comment.comment_text);
    } else if (comment.comment && Array.isArray(comment.comment)) {
      comment.comment.forEach(function(part) {
        if (part.text) {
          textParts.push(part.text);
        }
      });
    }
    
    var userName = '';
    if (comment.user) {
      userName = comment.user.username || comment.user.email || '';
    }
    
    var userInitials = '';
    if (userName) {
      var parts = userName.split(' ');
      userInitials = parts.map(function(p) { return p.charAt(0).toUpperCase(); }).join('').substring(0, 2);
    }
    
    var userColor = comment.user && comment.user.color ? comment.user.color : '#7c4dff';
    
    return {
      id: comment.id,
      text: textParts.join(''),
      date: comment.date ? parseInt(comment.date) : null,
      userName: userName,
      userInitials: userInitials,
      userColor: userColor,
      userId: comment.user ? comment.user.id : null
    };
  });
}

/**
 * Posts a new comment to a ClickUp task.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @param {string} commentText - The comment text to post
 * @returns {Object} The created comment data
 */
function postClickUpTaskComment(taskId, token, commentText) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }
  if (!commentText || commentText.trim() === '') {
    throw new Error('Текст коментаря не може бути порожнім.');
  }

  var payload = {
    comment_text: commentText.trim(),
    notify_all: true
  };

  const result = clickUpApiPost('/task/' + taskId + '/comment', resolveClickUpToken(token), payload);
  // The awaiting-reply cache would otherwise keep showing this task for up
  // to CLICKUP_AWAITING_REPLY_CACHE_SECONDS after we just replied to it.
  invalidateClickUpCache();
  return result;
}

const CLICKUP_AWAITING_REPLY_SCAN_LIMIT = 15;
const CLICKUP_AWAITING_REPLY_CACHE_SECONDS = 600;

/**
 * Finds open ClickUp tasks whose last comment is not from us — i.e. tasks
 * effectively waiting on a reply. Bounded to the top N tasks (in the same
 * due-date order as the task list) since checking comments costs one extra
 * API call per task; tasks with no comments at all have nothing to reply to
 * and are skipped.
 * @returns {Array} Array of {id, name, url, lastCommenter, lastCommentDate}
 */
function getClickUpAwaitingReplyTasks() {
  const token = getClickUpToken();
  if (!token) return [];

  const cached = cacheGetJson('cu_awaiting_reply');
  if (cached) return cached;

  let raw;
  try {
    raw = fetchClickUpTasksRaw(token, false);
  } catch (e) {
    Logger.log('Не вдалося перевірити задачі, що чекають відповіді: ' + e.message);
    return [];
  }

  const currentUserId = raw.currentUser ? raw.currentUser.id : null;
  const candidates = raw.tasks.slice(0, CLICKUP_AWAITING_REPLY_SCAN_LIMIT);

  const results = [];
  candidates.forEach(function (t) {
    try {
      const comments = getClickUpTaskComments(t.id, token);
      if (comments.length === 0) return;

      const last = comments[0]; // ClickUp returns newest-first
      if (currentUserId && last.userId === currentUserId) return; // we already have the last word

      results.push({
        id: t.id,
        name: t.name,
        url: t.url,
        lastCommenter: last.userName,
        lastCommentDate: last.date
      });
    } catch (e) {
      Logger.log('Не вдалося перевірити коментарі задачі ' + t.id + ': ' + e.message);
    }
  });

  cachePutJson('cu_awaiting_reply', results, CLICKUP_AWAITING_REPLY_CACHE_SECONDS);
  return results;
}

/**
 * Summarizes a ClickUp task together with its whole comment thread.
 * Most useful on tasks where the actual state lives in 30 comments rather
 * than in the description.
 * @param {string} taskId - The ClickUp task ID
 * @returns {string} Brief in Markdown format
 */
function aiClickUpTaskBrief(taskId) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const token = resolveClickUpToken(null);
  const task = getClickUpTaskDetails(taskId, token);

  var comments = [];
  try {
    comments = getClickUpTaskComments(taskId, token);
  } catch (e) {
    Logger.log('Не вдалося завантажити коментарі для брифу: ' + e.message);
  }

  const priorityNames = { 1: 'Urgent', 2: 'High', 3: 'Normal', 4: 'Low' };

  let taskBlock = `Назва: "${task.name}"\n`;
  taskBlock += `Статус: ${task.status.name}\n`;
  if (task.priority && priorityNames[task.priority]) taskBlock += `Пріоритет: ${priorityNames[task.priority]}\n`;
  if (task.dueDate) taskBlock += `Дедлайн: ${formatDateShort(new Date(task.dueDate).toISOString())}\n`;
  if (task.path) taskBlock += `Розташування: ${task.path}\n`;
  if (task.creator && task.creator.name) taskBlock += `Створив: ${task.creator.name}\n`;
  if (task.assignees && task.assignees.length > 0) {
    taskBlock += `Виконавці: ${task.assignees.map(function (a) { return a.name; }).join(', ')}\n`;
  }
  if (task.timeEstimate) taskBlock += `Оцінка: ${Math.round(task.timeEstimate / 60000)} хв\n`;
  if (task.timeSpent) taskBlock += `Витрачено: ${Math.round(task.timeSpent / 60000)} хв\n`;
  taskBlock += `\nОпис:\n${truncateForContext(task.description, 2000) || 'Немає'}\n`;

  if (task.checklists && task.checklists.length > 0) {
    taskBlock += '\nЧеклісти:\n';
    task.checklists.forEach(function (cl) {
      const done = cl.items.filter(function (i) { return i.resolved; }).length;
      taskBlock += `- ${cl.name} [${done}/${cl.items.length}]\n`;
      cl.items.forEach(function (item) {
        taskBlock += `  ${item.resolved ? '☑' : '☐'} ${item.name}\n`;
      });
    });
  }

  if (task.subtasks && task.subtasks.length > 0) {
    taskBlock += '\nПідзадачі:\n';
    task.subtasks.forEach(function (st) {
      taskBlock += `- [${st.status.name}] ${st.name}\n`;
    });
  }

  // ClickUp returns comments newest-first; reverse so the model reads the thread chronologically
  let commentsBlock = 'Коментарів немає.';
  if (comments.length > 0) {
    const ordered = comments.slice().reverse();
    commentsBlock = ordered.map(function (c) {
      const when = c.date ? new Date(c.date).toLocaleString('uk-UA') : '';
      return `[${when}] ${c.userName}: ${truncateForContext(c.text, 600)}`;
    }).join('\n');
  }

  const prompt = `ЗАДАЧА:\n${taskBlock}\n\nОБГОВОРЕННЯ (від найстарішого до найновішого):\n${commentsBlock}`;

  const systemInstruction = `Ти — асистент TaskApp. Тобі дано задачу з ClickUp разом із усім обговоренням у коментарях.
Поточна дата: ${new Date().toISOString().split('T')[0]}.

Зроби стислий бриф, щоб людина за 30 секунд зрозуміла стан справ, не читаючи всю переписку.

Бриф має містити:
1. **Суть** — чого від нас хочуть, одним-двома реченнями.
2. **Що вже зроблено** — на основі коментарів та чеклістів.
3. **Що блокує** — відкриті питання, очікування на когось, ризики. Якщо нічого не блокує — так і напиши.
4. **Наступний крок** — одна конкретна дія, яку варто зробити далі.

Правила:
- Відповідай українською у форматі Markdown.
- Будь конкретним: імена, дати, цифри з обговорення.
- Не вигадуй того, чого немає в наданих даних.
- Якщо коментарів немає — скажи це прямо і будуй бриф лише з опису.`;

  return callGemini({ prompt: prompt, systemInstruction: systemInstruction });
}

