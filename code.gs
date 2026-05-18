/**
 * Finance Pro V9.0.0 (Master Simple Edition)
 * บัญชี + ปิดงาน Tasks + Google Calendar Sync + LINE Status (9 Categories)
 * ยกเลิกระบบลงทะเบียนล็อกอิน เพื่อความคล่องตัวสูงสุดในการใช้งานร่วมกัน
 */

const scriptProperties = PropertiesService.getScriptProperties();
const CHANNEL_ACCESS_TOKEN = scriptProperties.getProperty('LINE_TOKEN');
const SPREADSHEET_ID = scriptProperties.getProperty('SS_ID');
const MY_USER_ID = scriptProperties.getProperty('MY_USER_ID'); 

/**
 * โหลดหน้าเว็บ Dashboard
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Executive Dashboard V9')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// 1. ระบบจัดการ GOOGLE CALENDAR
// ==========================================

function syncToCalendar(data) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const start = new Date(data.date + 'T' + data.time);
    const end = new Date(start.getTime() + (60 * 60 * 1000)); // นัด 1 ชม.
    
    const event = calendar.createEvent(data.event, start, end);
    event.addPopupReminder(30); // แจ้งเตือนล่วงหน้า 30 นาทีบนมือถือ
    return true;
  } catch (e) { 
    console.error("Calendar Sync Error: " + e.message);
    return false; 
  }
}

// ==========================================
// 2. ระบบจัดการ GOOGLE TASKS (งานค้าง)
// ==========================================

function completeTask(taskId) {
  try {
    const task = Tasks.Tasks.get("@default", taskId);
    task.status = "completed";
    Tasks.Tasks.update(task, "@default", taskId);
    return true;
  } catch (e) { return false; }
}

function saveToGoogleTasks(title, dueDate) {
  try {
    const task = { title: title, notes: "สร้างโดย Finance Pro Dashboard" };
    if (dueDate) {
      task.due = new Date(dueDate).toISOString();
    }
    Tasks.Tasks.insert(task, "@default");
    return true;
  } catch (e) { return false; }
}

function getActiveTasks() {
  try {
    const tasks = Tasks.Tasks.list("@default", { showCompleted: false, maxResults: 15 });
    return (tasks.items || []).map(t => ({ title: t.title, id: t.id }));
  } catch (e) { return []; }
}

// ==========================================
// 3. ระบบจัดการข้อมูล DASHBOARD
// ==========================================

function getFullDashboardData(month, yearBE) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const transSheet = ss.getSheetByName("Transactions");
    const transData = transSheet.getDataRange().getValues();
    
    const targetMonth = parseInt(month);
    const targetYear = parseInt(yearBE) - 543; // ค.ศ.
    
    let res = {
      income: 0, expense: 0, profit: 0,
      dailyProfit: [], 
      yearlyData: { labels: [], income: [], expense: [] },
      recentLogs: [],
      pendingTasks: getActiveTasks()
    };

    let runningProfit = 0;
    const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
    let dayMap = {};
    let yearlySum = Array.from({length: 12}, () => ({inc: 0, exp: 0}));

    for (let i = 1; i < transData.length; i++) {
      const row = transData[i];
      const tDate = new Date(row[0]);
      if (isNaN(tDate.getTime())) continue;

      // ประมวลผลรายปี
      if (tDate.getFullYear() === targetYear) {
        const m = tDate.getMonth();
        if (row[2] === "รายรับ") yearlySum[m].inc += Number(row[3]);
        else yearlySum[m].exp += Number(row[3]);
      }

      // ประมวลผลเดือนปัจจุบัน
      if (tDate.getMonth() === targetMonth && tDate.getFullYear() === targetYear) {
        const amt = Number(row[3]) || 0;
        const day = tDate.getDate();
        if (row[2] === "รายรับ") { res.income += amt; runningProfit += amt; } 
        else { res.expense += amt; runningProfit -= amt; }
        dayMap[day] = runningProfit;
      }
    }
    res.profit = res.income - res.expense;

    // เตรียมกราฟเส้นรายวัน
    let lastVal = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      if (dayMap[d] !== undefined) lastVal = dayMap[d];
      res.dailyProfit.push(lastVal);
    }

    // เตรียมกราฟแท่งรายเดือน
    const mNames = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
    yearlySum.forEach((m, idx) => {
      res.yearlyData.labels.push(mNames[idx]);
      res.yearlyData.income.push(m.inc);
      res.yearlyData.expense.push(m.exp);
    });

    // ดึงประวัติ 15 รายการล่าสุด
    res.recentLogs = transData.slice(1).map(r => ({
      date: Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM"),
      item: r[1], type: r[2], amount: r[3], cat: r[4]
    })).reverse().slice(0, 15);

    return res;
  } catch(e) { return { error: e.message }; }
}

// ==========================================
// 4. บันทึกข้อมูลจากการกรอกหน้าเว็บ
// ==========================================

function saveSchedule(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName("Schedules").appendRow([new Date(data.date), data.time, data.event, "รอดำเนินการ", "Web Dashboard"]);
  syncToCalendar(data); // ลงปฏิทินหลัก
  
  if (MY_USER_ID) {
    sendPush(MY_USER_ID, `📅 จดนัดหมายใหม่ลงปฏิทินแล้วครับ!\n⏰ เวลา: ${data.time}\n📋 เรื่อง: ${data.event}`);
  }
  return true;
}

function saveFromWeb(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName("Transactions").appendRow([new Date(), formData.item, formData.type, Number(formData.amount), formData.category, "Web Dashboard"]);
  return true;
}

// ==========================================
// 5. ระบบ LINE Webhook (แยกวิเคราะห์ 9 หมวดหมู่)
// ==========================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.events[0];
    if (!event || event.type !== 'message') return;
    const userText = event.message.text.trim();

    // ดูงานค้างผ่าน LINE
    if (userText === "งานค้าง" || userText === "task" || userText === "งาน") {
      const tasks = getActiveTasks();
      let msg = "📝 งานที่ค้างอยู่ในระบบของบอส:\n";
      if (tasks.length === 0) msg = "บอสทำเสร็จครบทุกงานแล้วครับ! ยอดเยี่ยมมาก 👍";
      else {
        tasks.forEach((t, i) => { msg += `${i+1}. ${t.title}\n`; });
        msg += "\nสามารถติ๊กจบงานได้ง่ายๆ ที่หน้าเว็บเลยครับ";
      }
      replyMessage(event.replyToken, msg);
      return;
    }

    // ดูนัดหมาย
    if (userText === "นัดหมาย" || userText === "ตารางงาน") {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const scheds = ss.getSheetByName("Schedules").getDataRange().getValues();
      const now = new Date();
      let msg = "📅 งานนัดหมายเร็วๆ นี้:\n";
      let count = 0;
      for (let i = scheds.length - 1; i >= 1; i--) {
        const d = new Date(scheds[i][0]);
        if (d >= new Date(now.setHours(0,0,0,0))) {
          msg += `\n📌 ${Utilities.formatDate(d, "GMT+7", "dd/MM")} เวลา ${scheds[i][1]}\n💬 ${scheds[i][2]}\n`;
          count++;
          if (count >= 5) break;
        }
      }
      replyMessage(event.replyToken, count > 0 ? msg : "ไม่มีนัดหมายเร็วๆ นี้ครับบอส");
      return;
    }

    // จดงานลง Google Tasks ด้วย #งาน หรือ #task
    if (userText.includes("#งาน") || userText.includes("#task")) {
      const taskTitle = userText.replace(/#งาน|#task/g, "").trim();
      saveToGoogleTasks(taskTitle);
      replyMessage(event.replyToken, `📝 เพิ่มงานเข้า Google Tasks เรียบร้อย!\n👉 ${taskTitle}`);
      return;
    }

    // บันทึกรายจ่ายอัตโนมัติ (9 หมวดหมู่)
    const amountMatch = userText.match(/\d+(\.\d+)?/);
    if (amountMatch) {
      const amount = parseFloat(amountMatch[0]);
      let rawItem = userText.replace(amountMatch[0], '').trim();
      let category = "อื่นๆ"; let type = "รายจ่าย"; let icon = "☁️";
      const item = rawItem.toLowerCase();

      // ระบบตรวจจับรายรับ
      if (/(เงินเดือน|รายได้|salary|โบนัส)/.test(item)) { category = "อื่นๆ"; type = "รายรับ"; icon = "💰"; }
      
      // 1. ค่าอาหาร
      else if (/(อาหาร|ข้าว|กิน|food|เหนียว|หมูกระทะ|ชาบู|ส้มตำ|ก๋วยเตี๋ยว)/.test(item)) { category = "ค่าอาหาร"; icon = "🍱"; }
      
      // 2. ค่าน้ำ/ไฟ/TEL/NET
      else if (/(น้ำ|ไฟ|ประปา|ไฟฟ้า|โทรศัพท์|มือถือ|เน็ต|ais|true|dtac|tel|net|wifi)/.test(item) && !/(น้ำดื่ม|น้ำเปล่า)/.test(item)) { category = "ค่าน้ำ/ไฟ/TEL/NET"; icon = "⚡"; }
      
      // 3. ค่าทางด่วน
      else if (/(ทางด่วน|easy pass|easypass|m-pass|mpass|tollway)/.test(item)) { category = "ค่าทางด่วน"; icon = "🛣️"; }
      
      // 4. ค่าน้ำมัน
      else if (/(น้ำมัน|gasoline|diesel|เติมน้ำมัน)/.test(item)) { category = "ค่าน้ำมัน"; icon = "⛽"; }
      
      // 5. ค่าแก๊ส
      else if (/(แก๊ส|เติมแก๊ส|lpg)/.test(item)) { category = "ค่าแก๊ส"; icon = "🔥"; }
      
      // 6. ค่าครอบครัว
      else if (/(ครอบครัว|บ้าน|ลูก|ขนมลูก|ของเล่น|เมีย|ซื้อของเข้าบ้าน|family)/.test(item)) { category = "ค่าครอบครัว"; icon = "🏠"; }
      
      // 7. ค่ากาแฟ
      else if (/(กาแฟ|coffee|ชา|สตาร์บัคส์|คาเฟ่)/.test(item)) { category = "ค่ากาแฟ"; icon = "☕"; }
      
      // 8. ค่าเตะบอล
      else if (/(เตะบอล|ฟุตบอล|บอล|สนามบอล|soccer|football)/.test(item)) { category = "ค่าเตะบอล"; icon = "⚽"; }

      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      ss.getSheetByName("Transactions").appendRow([new Date(), rawItem, type, amount, category, "LINE Bot"]);
      replyMessage(event.replyToken, `บันทึกเรียบร้อย! ${icon}\n📌 ${rawItem}: ${amount.toLocaleString()} บ.\n📁 หมวดหมู่: ${category}`);
    }
  } catch (err) {}
}

function replyMessage(token, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'method': 'post', 'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    'payload': JSON.stringify({ 'replyToken': token, 'messages': [{'type': 'text', 'text': text}] })
  });
}

function sendPush(to, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    'method': 'post', 'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    'payload': JSON.stringify({ 'to': to, 'messages': [{'type': 'text', 'text': text}] })
  });
}
