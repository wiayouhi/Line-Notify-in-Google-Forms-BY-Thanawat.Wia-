# 📢 Google Form to LINE Notify Integration 🚀
**เปลี่ยน Google Sheets ให้เป็นระบบแจ้งเตือน LINE Notify อย่างสวยงาม พร้อมด้วย GIF อนิเมะสุดน่ารัก!**

![Anime Banner GIF](https://media.tenor.com/P7fsIu00v04AAAAM/the-boys-osrs.gif)

---

## 🛠️ ฟีเจอร์หลัก

### 🕒 แจ้งเตือนแบบเรียลไทม์
- ส่งข้อความไปยัง LINE ทันทีเมื่อมีข้อมูลใหม่ใน Google Sheets

### 🎞️ เลือก GIF อนิเมะอัตโนมัติ
- เลือกภาพ GIF ตามประเภทการลา เช่น
  - 🤒 *ลาป่วย* ![Sick Anime GIF](https://media.tenor.com/qJFIZnigq9oAAAAM/holo-wise-wolf.gif)
  - 🏡 *ลากิจ* ![Vacation Anime GIF](https://media.tenor.com/S_bAl5e1ZEYAAAAM/luffy-breakdown-one-piece.gif)

### 📝 ข้อความปรับแต่งได้
- แสดงข้อมูล เช่น ชื่อ, เหตุผล, วันที่ และอีกมากมาย พร้อมอิโมจิน่ารัก ๆ

---

## 🔧 การตั้งค่าโค้ด

### 1️⃣ เปลี่ยน Token ของ LINE Notify
ในโค้ด ให้ค้นหาและแก้ไข Token ตรงนี้:
```javascript
var token = "YOUR_LINE_NOTIFY_TOKEN"; // 🔄 เปลี่ยนเป็น Token ของคุณ
```

### 2️⃣ เพิ่มประเภทการลาและ GIF
ในส่วนของ `switch (item)` เพิ่ม GIF ใหม่ถ้าจำเป็น:
```javascript
case "ลาพักร้อน":
    gifUrl = "https://media.giphy.com/media/l46Cy1rHbQ92uuLXa/giphy.gif"; // 🔄 ใส่ URL GIF ที่ต้องการ
    break;
```

### 3️⃣ ลิงก์ Google Sheet
แก้ไขลิงก์ของ Google Sheets ตรงนี้:
```javascript
var linkMessage = "📎 **ลิงก์ Google Sheets:**\nhttps://docs.google.com/spreadsheets/d/your-spreadsheet-id"; // 🔄 ใส่ลิงก์ชีตของคุณ
```

---

## 📝 ตัวอย่างโค้ดที่ปรับปรุงแล้ว
```javascript
function GoogleFormToLine() {
    var sheet = SpreadsheetApp.getActiveSheet(); 
    var row = sheet.getLastRow();                
    var column = sheet.getLastColumn();          
    var headers = sheet.getRange(1, 1, 1, column).getValues()[0]; 
    var lastRowData = sheet.getRange(row, 1, 1, column).getValues()[0]; 
    var message = "\n📢 แจ้งการลา 📢\n\n---------------------------------------------\n\n";
    var linkMessage = "📎 **ลิงก์ Google Sheets:**\nhttps://docs.google.com/spreadsheets/d/your-spreadsheet-id"; 
    var gifUrl = "";

    for (var i = 0; i < column; i++) {
        var item = headers[i];
        var value = lastRowData[i];
        if (!value) continue;

        switch (item) {
            case "ลาประเภทใด":
                if (value === "ลาป่วย") {
                    gifUrl = "https://media.giphy.com/media/Lpnmiofq4MxFhzf3a8/giphy.gif";
                } else if (value === "ลากิจ") {
                    gifUrl = "https://media.giphy.com/media/l3q2IpBz5x1PgT0p6/giphy.gif";
                }
                break;
            // เพิ่มประเภทการลาอื่น ๆ ที่นี่
        }

        message += `📌 ${item}: ${value}\n`;
    }

    SendToLine(message);
    if (gifUrl) {
        SendToLineImage(gifUrl);
    }
}
```

---

## 🎥 ตัวอย่างการทำงาน
![Demo Animation](https://media.giphy.com/media/3oriO0OEd9QIDdllqo/giphy.gif)

✨ ตัวอย่างข้อความที่ส่งไป LINE Notify:
```
📢 แจ้งการลา 📢  
👤 ชื่อ: คุณสมชาย  
📅 วันที่: 05/01/2025  
📌 ประเภทการลา: ลาป่วย  
📝 เหตุผล: ไม่สบาย  
```
- รูปภาพที่ส่ง:  
  ![Sick Anime GIF](https://media.tenor.com/G2orKp98rJMAAAAM/alya.gif)

---

## 🌟 ติดตามเราได้ที่
📸 **Instagram**: [@thanawat.wia](https://instagram.com/thanawat.wia)  
🐙 **GitHub**: [@wiayouhi](https://github.com/wiayouhi)

