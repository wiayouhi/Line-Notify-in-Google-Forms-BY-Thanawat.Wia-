function GoogleFormToLine() {
    var sheet = SpreadsheetApp.getActiveSheet(); 
    var row = sheet.getLastRow();                
    var column = sheet.getLastColumn();          
    var headers = sheet.getRange(1, 1, 1, column).getValues()[0]; 
    var lastRowData = sheet.getRange(row, 1, 1, column).getValues()[0]; 
    var message = "\n📢 แจ้งการลา 📢\n\n---------------------------------------------\n\n "; // เริ่มข้อความแจ้งเตือนพร้อมหัวข้อใหญ่
    var linkMessage = "📎 **ลิงก์ Google Sheets:**\nhttps://docs.google.com/spreadsheets/d/1zYijOIllUhAEIlRddD2tcBhDMzLHdwTkuqCmyAgNv7w/edit?resourcekey=&gid=1021450421#gid=1021450421\n"; // ลิงก์ของ Google Sheets
    var medicalFileLink = ""; // ลิงก์ไฟล์แนบใบรับรองทางการแพทย์
    var leaveType = ""; // ประเภทการลา
    var gifUrl = ""; // URL ของภาพ GIF
  
    // สร้างข้อความจากข้อมูลในแถวล่าสุด
    for (var i = 0; i < column; i++) {
      var item = headers[i];
      var value = lastRowData[i];
  
      if (!value) continue; // ข้ามช่องว่าง
  
      switch (item) {
        case "ประทับเวลา":
        case "วันที่":
        case "วันที่แจ้งลา":
        case "ตั้งแต่วันที่":
        case "จนถึงวันที่":
          try {
            value = Utilities.formatDate(new Date(value), "GMT+7", "dd/MM/yyyy(E) HH:mm:ss");
            message += ` 📅 ${item}:\n    ${value}\n\n`;
          } catch (e) {
            message += ` 📅 ${item}:\n    ไม่สามารถแปลงเวลาได้\n\n`;
          }
          break;
        case "ชื่อ-นามสกุล":
          message += ` 👤${item}:\n    - ${value}\n\n`;
          break;
        case "เลขที่ประจำตัว":
          message += ` 🆔 ${item}:\n    - ${value}\n\n`;
          break;
        case "ชั้น":
          message += ` 🏫 ${item}:\n    - ${value}\n\n`;
          break;
        case "เลขที่":
          message += ` #️⃣ ${item}:\n    - ${value}\n\n`;
          break;
        case "ลาประเภทใด":
          leaveType = value; // เก็บประเภทการลา
          message += ` 📌 ${item}:\n    - ${value}\n\n`;
          // เลือกภาพ GIF ตามประเภทการลา
          if (value === "ลาป่วย") {
            gifUrl = "https://cdn.discordapp.com/attachments/1064441940583661590/1314464678834999397/1.png?ex=6753de21&is=67528ca1&hm=d564de2a3b5cfe0fc668a87f315ca0146e413ca2e29001184e36b1c8dca0a724&"; // ลิงก์ GIF สำหรับลาป่วย
          } else if (value === "ลากิจ") {
            gifUrl = "https://cdn.discordapp.com/attachments/1064441940583661590/1314464678558044230/2.png?ex=6753de20&is=67528ca0&hm=ba7d84c88065fcb377887a54f556259ef9b98e8a04b7777e2b25b19da0759b08&"; // ลิงก์ GIF สำหรับลากิจ
          }
          break;
        case "เหตุผล":
          message += ` 📝 ${item}:\n    - ${value}\n\n`;
          break;
        case "จำนวนวันลา":
          message += ` 📊 ${item}:\n    - ${value}\n\n`;
          break;
        case "รักษาที่ไหน":
          message += ` 🏥 ${item}:\n    - ${value}\n\n`;       
          break;
        case "อาการป่วย":
          message += ` 😷 ${item}:\n    - ${value}\n\n`;
          break;  
        case "เบอร์โทรของผู้ปกครอง":
          message += ` 📞 ${item}:\n    - ${value}\n\n`;
          break;
        case "เบอร์โทรของนักเรียน":
          message += ` 📞 ${item}:\n    - ${value}\n\n`;
          break; 
        case "ยืนยันว่าข้อมูลนี้เป็นจริง":
          message += ` 🟢 ${item};\n    - ${value}\n\n`;
          break;
        case "ที่อยู่อีเมล":
          message += ` ✉️ ${item};\n    - ${value}\n\n`;
          break;    
        case "ไฟล์แนบใบรับรองทางการแพทย์": // กรณีเป็นลิงก์
          medicalFileLink += `📎 ลิงก์ไฟล์แนบใบรับรองทางการแพทย์ 📁: \n${value}\n`; // เก็บลิงก์ไฟล์
          break;
        default:
          message += `- 🪪 ข้อมูลเพิ่มเติม ${item}:\n    - ${value}\n\n`;
      }
    }
  
    // ส่งข้อความหลัก (ข้อมูลการลา)
    SendToLine(message);
  
    // ส่งลิงก์ Google Sheets
    SendToLine(linkMessage);
  
    // ถ้ามีลิงก์ไฟล์แนบใบรับรองทางการแพทย์ ให้ส่งแยกต่างหาก
    if (medicalFileLink) {
      SendToLine(medicalFileLink);
    }
  
    // ถ้ามี URL GIF ที่กำหนด ให้ส่งภาพ GIF
    if (gifUrl) {
      SendToLineImage(gifUrl);
    }
  }
  
  // ฟังก์ชันส่งข้อความไป LINE Notify
  function SendToLine(message) {
    var token = "OhK3NnC8EdMDftBX3cbUnWWmvgpwhaRFLcU7u6IKXIH"; // เปลี่ยนเป็น Token ของ LINE Notify
    var options = {
      method: "post",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: {
        message: message
      }
    };
  
    try {
      var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
      Logger.log(response.getContentText());
    } catch (e) {
      Logger.log("Error: " + e.message);
    }
  }
  
  // ฟังก์ชันส่งภาพ GIF ไป LINE Notify
  function GoogleFormToLine() {
    var sheet = SpreadsheetApp.getActiveSheet(); 
    var row = sheet.getLastRow();                
    var column = sheet.getLastColumn();          
    var headers = sheet.getRange(1, 1, 1, column).getValues()[0]; 
    var lastRowData = sheet.getRange(row, 1, 1, column).getValues()[0]; 
    var message = "\n📢 แจ้งการลา 📢\n\n---------------------------------------------\n\n "; // เริ่มข้อความแจ้งเตือนพร้อมหัวข้อใหญ่
    var linkMessage = "📎 **ลิงก์ Google Sheets:**\nhttps://docs.google.com/spreadsheets/d/1zYijOIllUhAEIlRddD2tcBhDMzLHdwTkuqCmyAgNv7w/edit?resourcekey=&gid=1021450421#gid=1021450421\n"; // ลิงก์ของ Google Sheets
    var medicalFileLink = ""; // ลิงก์ไฟล์แนบใบรับรองทางการแพทย์
    var leaveType = ""; // ประเภทการลา
    var gifUrl = ""; // URL ของภาพ GIF
  
    // สร้างข้อความจากข้อมูลในแถวล่าสุด
    for (var i = 0; i < column; i++) {
      var item = headers[i];
      var value = lastRowData[i];
  
      if (!value) continue; // ข้ามช่องว่าง
  
      switch (item) {
        case "ประทับเวลา":
        case "วันที่":
        case "วันที่แจ้งลา":
        case "ตั้งแต่วันที่":
        case "จนถึงวันที่":
          try {
            value = Utilities.formatDate(new Date(value), "GMT+7", "dd/MM/yyyy(E) HH:mm:ss");
            message += ` 📅 ${item}:\n    ${value}\n\n`;
          } catch (e) {
            message += ` 📅 ${item}:\n    ไม่สามารถแปลงเวลาได้\n\n`;
          }
          break;
        case "ชื่อ-นามสกุล":
          message += ` 👤${item}:\n    - ${value}\n\n`;
          break;
        case "เลขที่ประจำตัว":
          message += ` 🆔 ${item}:\n    - ${value}\n\n`;
          break;
        case "ชั้น":
          message += ` 🏫 ${item}:\n    - ${value}\n\n`;
          break;
        case "เลขที่":
          message += ` #️⃣ ${item}:\n    - ${value}\n\n`;
          break;
        case "ลาประเภทใด":
          leaveType = value; // เก็บประเภทการลา
          message += ` 📌 ${item}:\n    - ${value}\n\n`;
          // เลือกภาพ GIF ตามประเภทการลา
          if (value === "ลาป่วย") {
            gifUrl = "https://cdn.discordapp.com/attachments/1064441940583661590/1314464678834999397/1.png?ex=6761b5e1&is=67606461&hm=5d92e7753fce247ff89c8c06f2c824bdf4a13913391c24ff3b0887e69ce49abe&"; // ลิงก์ GIF สำหรับลาป่วย
          } else if (value === "ลากิจ") {
            gifUrl = "https://cdn.discordapp.com/attachments/1064441940583661590/1314464678558044230/2.png?ex=6761b5e0&is=67606460&hm=9d3eebb7feb1b39560e03ca619cd276dd457943dc2208aa419e343e050aa96b2&"; // ลิงก์ GIF สำหรับลากิจ
          }
          break;
        case "เหตุผล":
          message += ` 📝 ${item}:\n    - ${value}\n\n`;
          break;
        case "จำนวนวันลา":
          message += ` 📊 ${item}:\n    - ${value}\n\n`;
          break;
        case "รักษาที่ไหน":
          message += ` 🏥 ${item}:\n    - ${value}\n\n`;       
          break;
        case "อาการป่วย":
          message += ` 😷 ${item}:\n    - ${value}\n\n`;
          break;  
        case "เบอร์โทรของผู้ปกครอง":
          message += ` 📞 ${item}:\n    - ${value}\n\n`;
          break;
        case "เบอร์โทรของนักเรียน":
          message += ` 📞 ${item}:\n    - ${value}\n\n`;
          break; 
        case "ยืนยันว่าข้อมูลนี้เป็นจริง":
          message += ` 🟢 ${item};\n    - ${value}\n\n`;
          break;
        case "ที่อยู่อีเมล":
          message += ` ✉️ ${item};\n    - ${value}\n\n`;
          break; 
        case "ห้อง":
          message += ` 👨‍🎓 ${item};\n    - ${value}\n\n`;
          break  
        case "มีใบรับรองทางการแพทย์หรือไม่":
          message += ` 🧾 ${item};\n    - ${value}\n\n`;
          break     
        case "ไฟล์แนบใบรับรองทางการแพทย์": // กรณีเป็นลิงก์
          medicalFileLink += `📎 ลิงก์ไฟล์แนบใบรับรองทางการแพทย์ 📁: \n${value}\n`; // เก็บลิงก์ไฟล์
          break;
        default:
          message += `- 🪪 ข้อมูลเพิ่มเติม ${item}:\n    - ${value}\n\n`;
      }
    }
  
    // ส่งข้อความหลัก (ข้อมูลการลา)
    SendToLine(message);
  
    // ส่งลิงก์ Google Sheets
    SendToLine(linkMessage);
  
    // ถ้ามีลิงก์ไฟล์แนบใบรับรองทางการแพทย์ ให้ส่งแยกต่างหาก
    if (medicalFileLink) {
      SendToLine(medicalFileLink);
    }
  
    // ถ้ามี URL GIF ที่กำหนด ให้ส่งภาพ GIF
    if (gifUrl) {
      SendToLineImage(gifUrl);
    }
  }
  
  // ฟังก์ชันส่งข้อความไป LINE Notify
  function SendToLine(message) {
    var token = "token line notify"; // เปลี่ยนเป็น Token ของ LINE Notify
    var options = {
      method: "post",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: {
        message: message
      }
    };
  
    try {
      var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
      Logger.log(response.getContentText());
    } catch (e) {
      Logger.log("Error: " + e.message);
    }
  }
  
  // ฟังก์ชันส่งภาพ GIF ไป LINE Notify
  // ฟังก์ชันส่งภาพ GIF ไป LINE Notify
  function SendToLineImage(imageUrl) {
    var token = "token line notify"; // เปลี่ยนเป็น Token ของ LINE Notify
  
    // ส่งข้อความแรก "กำลังส่งรูป..."
    var optionsMessage = {
      method: "post",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: {
        message: "📢 กำลังโหลดรูป...ดึงภาพจากฐานข้อมูล 🔔", // ข้อความบรรยายสำหรับภาพ
      }
    };
  
    try {
      // ส่งข้อความแรก
      var responseMessage = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", optionsMessage);
      Logger.log(responseMessage.getContentText());
      
      // รอ 2 วินาที (สามารถปรับเวลาตามความเหมาะสม)
      Utilities.sleep(3000); 
  
      // ส่งข้อความว่า "🟢 succeed 🟢" พร้อมกับการคลูดาว
      var optionsSuccessMessage = {
        method: "post",
        headers: {
          Authorization: "Bearer " + token
        },
        payload: {
          message: "🟢 succeed 🟢", // ข้อความบรรยายสำหรับภาพ
        }
      };  
      var responseSuccessMessage = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", optionsSuccessMessage);
      Logger.log(responseSuccessMessage.getContentText());
      
      // ส่งภาพหลังจากข้อความ
      var optionsImage = {
        method: "post",
        headers: {
          Authorization: "Bearer " + token
        },
        payload: {
          message: "📸 นี่คือภาพ!", // ข้อความบรรยายสำหรับภาพ
          imageFullsize: imageUrl,  // ขนาดเต็มของภาพ
          imageThumbnail: imageUrl // ขนาดย่อของภาพ
        }
      };  
  
      // ส่งภาพ
      var responseImage = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", optionsImage);
      Logger.log(responseImage.getContentText());
    } catch (e) {
      Logger.log("Error: " + e.message);
    }
}


  
  
  
