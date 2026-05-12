const BOT_TOKEN = '8725305732:AAHNJ08hpQ6SUR_zVgns4EG1-58VPTcRUpA';
const SHEET_ID = '1dYzU7SBxal9jE4gqhoZyXoBM4wFxtVU1Z2wzm2O1PZo';
const GEMINI_API_KEY = 'AIzaSyCZWVJ3bdgtlWnqSJyiHIFEZcqAELYV9';
const MINI_APP_URL = 'https://elias2025new.github.io/Self-Improvment/';
const TELEGRAM_API_URL = `https://api.telegram.org/bot${BOT_TOKEN}`;

// ==========================================
// ENDPOINTS
// ==========================================

// Handle POST requests (Telegram Webhooks & Mini App Actions)
function doPost(e) {
  try {
    // 1. Check if this is a Telegram Webhook
    if (e.postData && e.postData.contents) {
      const contents = JSON.parse(e.postData.contents);
      if (contents.message) {
        handleTelegramMessage(contents.message);
        return ContentService.createTextOutput('OK');
      }
    }
  } catch (error) {
    // Ignore JSON parse errors, move on to check if it's a Mini App action
  }

  // 2. Handle Mini App POST Actions (e.g., mark as done)
  // We use URL query parameters for actions to avoid CORS preflight issues with JSON bodies
  const action = e.parameter.action;
  const userId = e.parameter.userId;
  const username = e.parameter.username || 'User';

  if (action === 'done' && userId) {
    const nextTask = handleTaskCompleted(userId, username);
    return createJsonResponse(nextTask);
  }

  return createJsonResponse({ error: "Invalid POST request" });
}

// Handle GET requests (Mini App data fetching)
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
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tasks');
}

function getUserHistory(userId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const history = [];
  let pendingTask = null;
  let pendingRowIndex = -1;
  
  // Start from 1 to skip headers
  for (let i = 1; i < data.length; i++) {
    // Force string comparison for IDs
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
        pendingRowIndex = i + 1; // Apps Script rows are 1-indexed
      } else if (task.status === 'Done') {
        history.push(task);
      }
    }
  }
  return { history, pendingTask, pendingRowIndex };
}

function handleGetTask(userId, username) {
  const { history, pendingTask } = getUserHistory(userId);
  
  if (pendingTask) {
    return pendingTask;
  }
  
  // If no pending task exists, generate one
  return generateNextTask(userId, username, history);
}

function handleTaskCompleted(userId, username) {
  const { history, pendingTask, pendingRowIndex } = getUserHistory(userId);
  
  if (pendingTask && pendingRowIndex > -1) {
    const sheet = getSheet();
    // Update status to Done
    sheet.getRange(pendingRowIndex, 10).setValue('Done');
    // Update Date Completed
    sheet.getRange(pendingRowIndex, 12).setValue(new Date().toISOString());
    
    // Add current to history so Gemini knows it's done immediately
    history.push(pendingTask);
  }
  
  // Generate and return the next task
  return generateNextTask(userId, username, history);
}

function handleGetProgress(userId) {
  const { history } = getUserHistory(userId);
  
  const progress = {
    total: history.length,
    pillars: {
      "Marketing": 0,
      "Thinking": 0,
      "Discipline": 0,
      "Self-Growth": 0
    },
    streak: calculateStreak(history),
    recent: history.slice(-10).reverse() // Last 10 tasks
  };
  
  history.forEach(task => {
    if (progress.pillars[task.pillar] !== undefined) {
      progress.pillars[task.pillar]++;
    }
  });
  
  return progress;
}

function calculateStreak(history) {
  if (history.length === 0) return 0;
  
  // Group by local date string
  const dates = history.map(t => new Date(t.dateCompleted).toDateString());
  const uniqueDates = [...new Set(dates)].sort((a, b) => new Date(b) - new Date(a));
  
  let streak = 0;
  let currentDate = new Date();
  currentDate.setHours(0,0,0,0);
  
  for (let i = 0; i < uniqueDates.length; i++) {
    const taskDate = new Date(uniqueDates[i]);
    taskDate.setHours(0,0,0,0);
    
    const diffDays = Math.floor((currentDate - taskDate) / (1000 * 60 * 60 * 24));
    
    if (diffDays === 0) {
      if (streak === 0) streak = 1; // Done today
    } else if (diffDays === 1) {
      streak++;
      currentDate = taskDate; // Shift reference to yesterday
    } else if (diffDays > 1 && i === 0 && streak === 0) {
      // Missed yesterday and today
      return 0; 
    } else {
      break;
    }
  }
  return streak;
}

// ==========================================
// AI LOGIC (GEMINI)
// ==========================================

function generateNextTask(userId, username, history) {
  let prompt = "";
  
  if (history.length === 0) {
    prompt = `You are an elite personal growth coach talking to a new student. 
Create the VERY FIRST foundational task for this user to start their growth journey.
It must belong to one of these 4 pillars: Marketing, Thinking, Discipline, Self-Growth.
It must be one of these types: Read, Listen, Watch, Do.
Task should take 15-30 minutes.

Return ONLY valid JSON in this exact structure with no markdown formatting:
{
  "pillar": "Thinking",
  "type": "Watch",
  "title": "Task title here",
  "description": "2-3 sentence description of what to do and why it matters",
  "resource": "https://specific-url.com or exact title and creator",
  "whyNext": "One sentence explaining why this is a great starting point"
}`;
  } else {
    const historyClean = history.map(t => ({ title: t.title, pillar: t.pillar, type: t.type }));
    const lastPillars = history.slice(-2).map(t => t.pillar);
    const lastTypes = history.slice(-2).map(t => t.type);
    
    prompt = `You are an elite personal growth coach. Here is the user's completed task history in order:
${JSON.stringify(historyClean)}

Create the NEXT task for this user. 
Rules:
1. USE GOOGLE SEARCH GROUNDING to find a REAL, CURRENT, AND HIGH-QUALITY resource (article, youtube video, podcast, etc).
2. Never repeat anything from their history.
3. Go one level deeper or build upon what they just did.
4. Rotate pillars. Do not use the same pillar more than twice in a row. Recent pillars: ${lastPillars.join(', ')}. Allowed: Marketing, Thinking, Discipline, Self-Growth.
5. Rotate types. Do not use the same type twice in a row. Recent types: ${lastTypes.join(', ')}. Allowed: Read, Listen, Watch, Do.

Return ONLY valid JSON in this exact structure with no markdown formatting:
{
  "pillar": "...",
  "type": "...",
  "title": "...",
  "description": "...",
  "resource": "https://specific-url.com or exact title",
  "whyNext": "..."
}`;
  }

  const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ googleSearch: {} }],
    generationConfig: {
      temperature: 0.7,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(geminiUrl, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.candidates && data.candidates[0].content.parts[0].text) {
      let jsonText = data.candidates[0].content.parts[0].text;
      const newTask = JSON.parse(jsonText);
      
      // Save to sheet
      const sheet = getSheet();
      const taskNumber = history.length + 1;
      const dateAssigned = new Date().toISOString();
      
      sheet.appendRow([
        userId,
        username,
        taskNumber,
        newTask.pillar,
        newTask.type,
        newTask.title,
        newTask.description,
        newTask.resource,
        newTask.whyNext,
        'Pending',
        dateAssigned,
        ''
      ]);
      
      // Assign the missing fields before returning to Mini App
      newTask.taskNumber = taskNumber;
      newTask.status = 'Pending';
      newTask.dateAssigned = dateAssigned;
      newTask.dateCompleted = '';
      
      return newTask;
    } else {
      throw new Error("Invalid Gemini response format");
    }
  } catch (e) {
    Logger.log("Error calling Gemini: " + e.toString());
    // Fallback task if Gemini fails to respond or rate limit
    return {
      pillar: "Self-Growth",
      type: "Do",
      title: "System Maintenance Pause",
      description: "Our AI brain is taking a breather. For now, take 5 deep breaths and reflect on the progress you've made so far.",
      resource: "",
      whyNext: "A moment of pause is essential for sustainable growth.",
      status: "Pending"
    };
  }
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
    sendMessage(chatId, "Welcome to your Personal Growth Companion 🌱\n\nI curate a hyper-personalized growth journey for you, one task at a time. Tap the button below to open your dashboard.", {
      inline_keyboard: [[{
        text: "Open Growth App 🌱",
        web_app: { url: MINI_APP_URL }
      }]]
    });
  } else if (text === 'done' || text === '/done') {
    sendMessage(chatId, "Processing your completion... 🧠");
    const nextTask = handleTaskCompleted(userId, username);
    sendTaskToChat(chatId, nextTask);
  } else if (text === '/current') {
    const task = handleGetTask(userId, username);
    sendTaskToChat(chatId, task);
  } else if (text === '/progress') {
    const progress = handleGetProgress(userId);
    let msg = `📊 *Your Progress*\n\n`;
    msg += `Total Tasks Completed: ${progress.total}\n`;
    msg += `Current Streak: ${progress.streak} days 🔥\n\n`;
    msg += `*Pillars:*\n`;
    msg += `📣 Marketing: ${progress.pillars['Marketing']}\n`;
    msg += `🧠 Thinking: ${progress.pillars['Thinking']}\n`;
    msg += `💪 Discipline: ${progress.pillars['Discipline']}\n`;
    msg += `🌿 Self-Growth: ${progress.pillars['Self-Growth']}\n`;
    
    sendMessage(chatId, msg);
  }
}

function sendTaskToChat(chatId, task) {
  const icons = {
    "Marketing": "📣", "Thinking": "🧠", "Discipline": "💪", "Self-Growth": "🌿",
    "Read": "📖", "Listen": "🎧", "Watch": "📺", "Do": "✍️"
  };
  
  const pIcon = icons[task.pillar] || "📌";
  const tIcon = icons[task.type] || "🔹";
  
  let msg = `${pIcon} *${task.pillar}* | ${tIcon} *${task.type}*\n\n`;
  msg += `*${task.title}*\n\n`;
  msg += `${task.description}\n\n`;
  if (task.resource) {
    if (task.resource.startsWith('http')) {
      msg += `🔗 [Resource Link](${task.resource})\n\n`;
    } else {
      msg += `🔍 Search for: *${task.resource}*\n\n`;
    }
  }
  msg += `💡 _Why this next:_ ${task.whyNext}`;

  sendMessage(chatId, msg, {
    inline_keyboard: [[{
      text: "Open App to Complete ✅",
      web_app: { url: MINI_APP_URL }
    }]]
  });
}

function sendMessage(chatId, text, replyMarkup = null) {
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown'
  };
  
  if (replyMarkup) {
    payload.reply_markup = JSON.stringify(replyMarkup);
  }

  UrlFetchApp.fetch(TELEGRAM_API_URL + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

// ==========================================
// HELPERS
// ==========================================

function createJsonResponse(data) {
  // Return JSON data without strict CORS headers since we use GET/URL queries for POST
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ⚠️ ONLY RUN THIS FUNCTION ONCE FROM THE EDITOR AFTER YOU PUBLISH AS WEB APP
function setupWebhook() {
  // Replace this with your exact published Google Apps Script Web App URL!
  const WEB_APP_URL = "YOUR_WEB_APP_URL_HERE"; 
  const response = UrlFetchApp.fetch(`${TELEGRAM_API_URL}/setWebhook?url=${WEB_APP_URL}`);
  Logger.log(response.getContentText());
}
