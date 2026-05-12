const BOT_TOKEN = '8725305732:AAHNJ08hpQ6SUR_zVgns4EG1-58VPTcRUpA';
const SHEET_ID = '1dYzU7SBxal9jE4gqhoZyXoBM4wFxtVU1Z2wzm2O1PZo';
const GEMINI_API_KEY = 'AIzaSyCZWVJ3bdgtlWnqSJyiHIFEZcqAELYV9';
const MINI_APP_URL = 'https://self-improvment.vercel.app'; 
const TELEGRAM_API_URL = `https://api.telegram.org/bot${BOT_TOKEN}`;
const TEST_MODE = true; // 🛠️ Set to false to enable the 24h lock for users

function doPost(e) {
  try {
    if (e.postData && e.postData.contents) {
      const contents = JSON.parse(e.postData.contents);
      
      // 🛡️ DEDUPLICATION: Stop Telegram from retrying if the AI is slow
      if (contents.update_id) {
        if (isDuplicateUpdate(contents.update_id)) {
          return ContentService.createTextOutput('OK');
        }
      }

      if (contents.message) {
        handleTelegramMessage(contents.message);
        return ContentService.createTextOutput('OK');
      }
    }
  } catch (error) {
    Logger.log("doPost Error: " + error.toString());
  }

  const action = e.parameter.action;
  const userId = e.parameter.userId;
  if (action === 'done' && userId) {
    const nextTask = handleTaskCompleted(userId, e.parameter.username || 'User');
    return createJsonResponse(nextTask);
  }
  return createJsonResponse({ error: "No action" });
}

// 🛡️ Deduplication Helper
function isDuplicateUpdate(updateId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastUpdateId = scriptProperties.getProperty('LAST_UPDATE_ID');
  if (lastUpdateId === String(updateId)) return true;
  scriptProperties.setProperty('LAST_UPDATE_ID', String(updateId));
  return false;
}

function doGet(e) {
  const action = e.parameter.action;
  const userId = e.parameter.userId;
  const username = e.parameter.username || 'User';

  if (action === 'getTask' && userId) {
    return createJsonResponse(handleGetTask(userId, username));
  } else if (action === 'done' && userId) {
    return createJsonResponse(handleTaskCompleted(userId, username));
  } else if (action === 'progress' && userId) {
    return createJsonResponse(handleGetProgress(userId));
  }
  return createJsonResponse({ error: "Action not found: " + action });
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tasks');
}

function getUserHistory(userId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const history = [];
  let pendingTask = null;
  let pendingRowIndex = -1;
  let lastCompletionDate = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      const task = {
        pillar: data[i][3], type: data[i][4], title: data[i][5],
        description: data[i][6], resource: data[i][7], whyNext: data[i][8],
        status: data[i][9], dateAssigned: data[i][10], dateCompleted: data[i][11]
      };
      if (task.status === 'Pending') { 
        pendingTask = task; 
        pendingRowIndex = i + 1; 
      } else { 
        history.push(task);
        if (task.dateCompleted) lastCompletionDate = new Date(task.dateCompleted).toDateString();
      }
    }
  }

  // 🛡️ LOCK LOGIC: Locked if assigned today OR completed another task today
  let isLocked = false;
  if (!TEST_MODE && pendingTask) {
    const today = new Date().toDateString();
    const assignedDate = new Date(pendingTask.dateAssigned).toDateString();
    if (assignedDate === today || lastCompletionDate === today) {
      isLocked = true;
    }
  }

  return { history, pendingTask, pendingRowIndex, isLocked };
}

function handleGetTask(userId, username) {
  const { history, pendingTask, isLocked } = getUserHistory(userId);
  const task = pendingTask || generateNextTask(userId, username, history);
  task.isLocked = isLocked;
  return task;
}

function handleTaskCompleted(userId, username) {
  const { history, pendingTask, pendingRowIndex, isLocked } = getUserHistory(userId);
  
  if (isLocked && !TEST_MODE) {
    return { error: "Growth takes time. Focus on your current task for today." };
  }

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
  const pillars = { "Marketing": 0, "Thinking": 0, "Discipline": 0, "Self-Growth": 0 };
  
  history.forEach(t => {
    if (pillars[t.pillar] !== undefined) pillars[t.pillar]++;
  });

  return {
    total: history.length,
    pillars: pillars,
    streak: calculateStreak(history),
    history: history.slice(-5).reverse() // Last 5 tasks for the UI
  };
}

function calculateStreak(history) {
  if (history.length === 0) return 0;
  
  // Get unique dates of completion
  const dates = history
    .filter(t => t.dateCompleted)
    .map(t => new Date(t.dateCompleted).toDateString());
  
  const uniqueDates = [...new Set(dates)].sort((a, b) => new Date(b) - new Date(a));
  
  if (uniqueDates.length === 0) return 0;

  let streak = 0;
  let today = new Date().toDateString();
  let yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  let yesterdayStr = yesterday.toDateString();

  // If the last task wasn't today or yesterday, streak is broken
  if (uniqueDates[0] !== today && uniqueDates[0] !== yesterdayStr) return 0;

  let currentCheck = new Date(uniqueDates[0]);
  streak = 1;

  for (let i = 1; i < uniqueDates.length; i++) {
    let prevDate = new Date(uniqueDates[i]);
    let diff = (currentCheck - prevDate) / (1000 * 60 * 60 * 24);
    
    if (diff === 1) {
      streak++;
      currentCheck = prevDate;
    } else {
      break;
    }
  }
  
  return streak;
}

function generateNextTask(userId, username, history) {
  // Use history titles to avoid repeats
  const historyTitles = history.map(t => t.title).join(", ");
  
  const prompt = "You are an elite personal growth coach. Create a hyper-personalized growth task for " + username + ". " +
                 "Pillars: Marketing, Thinking, Discipline, Self-Growth. " +
                 "Types: Read, Watch, Listen, Do. " +
                 "History (DO NOT REPEAT): " + historyTitles + ". " +
                 "Use Google Search to find a high-quality article or video. " +
                 "Return ONLY clean JSON in this structure: {\"pillar\":\"\",\"type\":\"\",\"title\":\"\",\"description\":\"\",\"resource\":\"\",\"whyNext\":\"\"}";
  
  const model = "gemini-2.5-flash"; 
  const geminiUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + GEMINI_API_KEY;

  try {
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      tools: [{ googleSearch: {} }],
      generationConfig: { temperature: 0.7 }
    };

    const response = UrlFetchApp.fetch(geminiUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    const resText = response.getContentText();
    const data = JSON.parse(resText);
    
    if (data.candidates && data.candidates[0].content.parts[0].text) {
      let jsonText = data.candidates[0].content.parts[0].text;
      jsonText = jsonText.replace(/```json/g, "").replace(/```/g, "").trim();
      
      const newTask = JSON.parse(jsonText);
      const sheet = getSheet();
      sheet.appendRow([userId, username, history.length + 1, newTask.pillar, newTask.type, newTask.title, newTask.description, newTask.resource, newTask.whyNext, 'Pending', new Date().toISOString(), '']);
      newTask.status = 'Pending';
      return newTask;
    } else {
      throw new Error("AI Error: " + resText);
    }
  } catch (e) {
    return {
      pillar: "Self-Growth", type: "Do", title: "Review Your Intentions",
      description: "AI Error: " + e.toString(),
      resource: "", whyNext: "Intentions are key.", status: "Pending"
    };
  }
}

function handleTelegramMessage(message) {
  const chatId = message.chat.id;
  const text = (message.text || "").toLowerCase();
  if (text === '/start') {
    sendMessage(chatId, "Welcome 🌱", {
      inline_keyboard: [[{ text: "Open Growth App 🌱", web_app: { url: MINI_APP_URL } }]]
    });
  }
}

function sendMessage(chatId, text, replyMarkup = null) {
  const payload = { chat_id: chatId, text: text, parse_mode: 'Markdown' };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  UrlFetchApp.fetch(TELEGRAM_API_URL + '/sendMessage', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload)
  });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function setupWebhook() {
  const webAppUrl = "PASTE_YOUR_WEB_APP_URL_HERE"; 
  UrlFetchApp.fetch(`${TELEGRAM_API_URL}/setWebhook?url=${webAppUrl}`);
}
