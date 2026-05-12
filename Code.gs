const BOT_TOKEN = '8725305732:AAHNJ08hpQ6SUR_zVgns4EG1-58VPTcRUpA';
const SHEET_ID = '1dYzU7SBxal9jE4gqhoZyXoBM4wFxtVU1Z2wzm2O1PZo';
const GEMINI_API_KEY = 'AIzaSyCZWVJ3bdgtlWnqSJyiHIFEZcqAELYV9';
const MINI_APP_URL = 'https://self-improvment.vercel.app'; // Update this if your Vercel URL is different!
const TELEGRAM_API_URL = `https://api.telegram.org/bot${BOT_TOKEN}`;

// ==========================================
// ENDPOINTS
// ==========================================

function doPost(e) {
  try {
    if (e.postData && e.postData.contents) {
      const contents = JSON.parse(e.postData.contents);
      
      // Deduplication: Prevent Telegram from retrying the same message
      if (contents.update_id && isDuplicateUpdate(contents.update_id)) {
        return ContentService.createTextOutput('OK');
      }

      if (contents.message) {
        handleTelegramMessage(contents.message);
        return ContentService.createTextOutput('OK');
      }
    }
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString());
  }

  const action = e.parameter.action;
  const userId = e.parameter.userId;
  const username = e.parameter.username || 'User';

  if (action === 'done' && userId) {
    const nextTask = handleTaskCompleted(userId, username);
    return createJsonResponse(nextTask);
  }

  return createJsonResponse({ error: "Invalid POST request" });
}

function doGet(e) {
  const action = e.parameter.action;
  const userId = e.parameter.userId;
  const username = e.parameter.username || 'User';

  if (action === 'getTask' && userId) {
    const task = handleGetTask(userId, username);
    return createJsonResponse(task);
  } else if (action === 'progress' && userId) {
    const progress = handleGetProgress(userId);
    return createJsonResponse(progress);
  }

  return createJsonResponse({ error: "Invalid GET request" });
}

// ==========================================
// CORE LOGIC & MEMORY
// ==========================================

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('Tasks');
}

// Check if we've already seen this Telegram message ID
function isDuplicateUpdate(updateId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastUpdateId = scriptProperties.getProperty('LAST_UPDATE_ID');
  if (lastUpdateId === String(updateId)) return true;
  scriptProperties.setProperty('LAST_UPDATE_ID', String(updateId));
  return false;
}

function getUserHistory(userId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const history = [];
  let pendingTask = null;
  let pendingRowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      const task = {
        taskNumber: data[i][2],
        pillar: data[i][3],
        type: data[i][4],
        title: data[i][5],
        description: data[i][6],
        resource: data[i][7],
        whyNext: data[i][8],
        status: data[i][9],
        dateAssigned: data[i][10],
        dateCompleted: data[i][11]
      };
      
      if (task.status === 'Pending') {
        pendingTask = task;
        pendingRowIndex = i + 1;
      } else if (task.status === 'Done') {
        history.push(task);
      }
    }
  }
  return { history, pendingTask, pendingRowIndex };
}

function handleGetTask(userId, username) {
  const { history, pendingTask } = getUserHistory(userId);
  if (pendingTask) return pendingTask;
  return generateNextTask(userId, username, history);
}

function handleTaskCompleted(userId, username) {
  const { history, pendingTask, pendingRowIndex } = getUserHistory(userId);
  if (pendingTask && pendingRowIndex > -1) {
    const sheet = getSheet();
    sheet.getRange(pendingRowIndex, 10).setValue('Done');
    sheet.getRange(pendingRowIndex, 12).setValue(new Date().toISOString());
    history.push(pendingTask);
  }
  return generateNextTask(userId, username, history);
}

function handleGetProgress(userId) {
  const { history } = getUserHistory(userId);
  const progress = {
    total: history.length,
    pillars: { "Marketing": 0, "Thinking": 0, "Discipline": 0, "Self-Growth": 0 },
    streak: calculateStreak(history)
  };
  history.forEach(task => { if (progress.pillars[task.pillar] !== undefined) progress.pillars[task.pillar]++; });
  return progress;
}

function calculateStreak(history) {
  if (history.length === 0) return 0;
  const dates = history.map(t => new Date(t.dateCompleted).toDateString());
  const uniqueDates = [...new Set(dates)].sort((a, b) => new Date(b) - new Date(a));
  let streak = 0;
  let currentDate = new Date();
  currentDate.setHours(0,0,0,0);
  for (let i = 0; i < uniqueDates.length; i++) {
    const taskDate = new Date(uniqueDates[i]);
    taskDate.setHours(0,0,0,0);
    const diffDays = Math.floor((currentDate - taskDate) / (1000 * 60 * 60 * 24));
    if (diffDays === 0) { if (streak === 0) streak = 1; }
    else if (diffDays === 1) { streak++; currentDate = taskDate; }
    else if (diffDays > 1 && i === 0 && streak === 0) return 0;
    else break;
  }
  return streak;
}

// ==========================================
// AI LOGIC (GEMINI)
// ==========================================

function generateNextTask(userId, username, history) {
  let prompt = `You are an elite personal growth coach. Create a hyper-personalized task for ${username}. 
Rules:
1. USE GOOGLE SEARCH to find a REAL, HIGH-QUALITY resource (article, youtube video, or podcast).
2. Never repeat anything from history: ${JSON.stringify(history.map(t => t.title))}
3. Pillars: Marketing, Thinking, Discipline, Self-Growth.
4. Types: Read, Listen, Watch, Do.
5. Return ONLY clean JSON in this structure:
{
  "pillar": "...",
  "type": "...",
  "title": "...",
  "description": "2 sentences",
  "resource": "URL",
  "whyNext": "1 sentence"
}`;

  const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ googleSearch: {} }],
    generationConfig: { temperature: 0.7, responseMimeType: "application/json" }
  };

  try {
    const response = UrlFetchApp.fetch(geminiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    if (data.candidates && data.candidates[0].content.parts[0].text) {
      const newTask = JSON.parse(data.candidates[0].content.parts[0].text);
      
      const sheet = getSheet();
      const taskNumber = history.length + 1;
      const dateAssigned = new Date().toISOString();
      
      sheet.appendRow([userId, username, taskNumber, newTask.pillar, newTask.type, newTask.title, newTask.description, newTask.resource, newTask.whyNext, 'Pending', dateAssigned, '']);
      
      newTask.taskNumber = taskNumber;
      newTask.status = 'Pending';
      return newTask;
    }
  } catch (e) {
    Logger.log("Gemini Error: " + e.toString());
  }

  return {
    pillar: "Self-Growth", type: "Do", title: "Review Your Intentions",
    description: "Take 5 minutes to write down exactly why you want to grow this week.",
    resource: "", whyNext: "Setting clear intentions is the foundation of change.", status: "Pending"
  };
}

// ==========================================
// TELEGRAM BOT LOGIC
// ==========================================

function handleTelegramMessage(message) {
  const chatId = message.chat.id;
  const text = message.text ? message.text.toLowerCase() : '';
  const userId = message.from.id;
  const username = message.from.username || message.from.first_name || 'User';

  if (text === '/start') {
    sendMessage(chatId, "Welcome to your Personal Growth Companion 🌱\n\nI curate a hyper-personalized growth journey for you, one task at a time.", {
      inline_keyboard: [[{ text: "Open Growth App 🌱", web_app: { url: MINI_APP_URL } }]]
    });
  } else if (text === 'done' || text === '/done') {
    sendMessage(chatId, "Processing completion... 🧠");
    const nextTask = handleTaskCompleted(userId, username);
    sendTaskToChat(chatId, nextTask);
  }
}

function sendTaskToChat(chatId, task) {
  let msg = `📌 *${task.pillar}* | *${task.type}*\n\n*${task.title}*\n\n${task.description}\n\n🔗 ${task.resource}`;
  sendMessage(chatId, msg, { inline_keyboard: [[{ text: "Open App ✅", web_app: { url: MINI_APP_URL } }]] });
}

function sendMessage(chatId, text, replyMarkup = null) {
  const payload = { chat_id: chatId, text: text, parse_mode: 'Markdown' };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  UrlFetchApp.fetch(TELEGRAM_API_URL + '/sendMessage', { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function setupWebhook() {
  const webAppUrl = "PASTE_YOUR_WEB_APP_URL_HERE"; 
  UrlFetchApp.fetch(`${TELEGRAM_API_URL}/setWebhook?url=${webAppUrl}`);
}
