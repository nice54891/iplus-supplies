// 📄 ไฟล์ employees.js (เวอร์ชันอัปเดตข้อมูลพนักงานและสรุปการลาสะสม ม.ค.-พ.ค. 2569 ครบถ้วน)
const excelEmployees = [
    { id: "1007", firstName: "ณิชนันทน์", lastName: "เอื้อเวชนิชกุล", nickname: "พีช", department: "ขายและการตลาด", position: "ผู้ช่วยผู้จัดการฝ่ายขายและการตลาด", leaveStats: { sick: 2, personal: 0, maternity: 0, vacation: 4 } },
    { id: "1015", firstName: "ญาติกา", lastName: "พงษ์สินมณี", nickname: "เชอรี่", department: "ขายและการตลาด", position: "ผู้ช่วยผู้จัดการฝ่ายขายและการตลาด", leaveStats: { sick: 2, personal: 0, maternity: 0, vacation: 2 } },
    { id: "1018", firstName: "รุ่งทิวา", lastName: "สิงห์พันธ์", nickname: "จิ๋ว", department: "ขายและการตลาด", position: "ผู้จัดการฝ่ายขายและการตลาด", leaveStats: { sick: "3ชม.", personal: 0, maternity: 0, vacation: 1 } },
    { id: "1033", firstName: "ชฎาพร", lastName: "วรรณเกษม", nickname: "แอน", department: "ขายและการตลาด", position: "เจ้าหน้าที่ประสานงานขายอาวุโส", leaveStats: { sick: 0, personal: 1, maternity: 0, vacation: 6 } },
    { id: "1041", firstName: "สิรัญญา", lastName: "นัดดา", nickname: "ส้ม", department: "ขายและการตลาด", position: "sales Co", leaveStats: { sick: 1, personal: 0, maternity: 0, vacation: "1 วัน 5 ชม." } },
    { id: "1043", firstName: "เมธาวี", lastName: "สมบัติ", nickname: "ครีม", department: "ขายและการตลาด", position: "ประสานงานขาย", leaveStats: { sick: 1, personal: "3 วัน 1.5 ชม.", maternity: 0, vacation: 1 } },
    { id: "1053", firstName: "ณัฐกฤตา", lastName: "คำวงษ์", nickname: "บีม", department: "ขายและการตลาด", position: "เจ้าหน้าที่อาวุโส Sales", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: "1 วัน 30 น." } },
    { id: "1063", firstName: "แพรวพรรณ", lastName: "มีคำ", nickname: "แพรว", department: "ขายและการตลาด", position: "เซลล์ซัพพอร์ต", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 0 } },
    { id: "1076", firstName: "กิ่งแก้ว", lastName: "ฟ้องเสียง", nickname: "กิ่ง", department: "ขายและการตลาด", position: "พนักงานขาย", leaveStats: { sick: "3 วัน 6 ชม.", personal: 0, maternity: 0, vacation: 1 } },
    { id: "1085", firstName: "เลา", lastName: "นามแสน", nickname: "เลา", department: "ขายและการตลาด", position: "พนักงานขายต่างจังหวัด", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 0 } },
    { id: "1087", firstName: "จีรนันท์", lastName: "โคตมณี", nickname: "ริน", department: "ขายและการตลาด", position: "Admin", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: "4 วัน 4.5 ชม." } },
    { id: "1095", firstName: "กรรณิกา", lastName: "บุญโพธิ์กอง", nickname: "แพน", department: "ขายและการตลาด", position: "Admin", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 2 } },
    { id: "1116", firstName: "ปิยธิดา", lastName: "ตันลือนาม", nickname: "เตย", department: "ขายและการตลาด", position: "พนักงานขาย", leaveStats: { sick: 2, personal: 0, maternity: 0, vacation: 1 } },
    { id: "1127", firstName: "ทักษ์ดนัย", lastName: "สารวัน", nickname: "ปีเตอร์แพน", department: "ขายและการตลาด", position: "วิเคราะห์ข้อมูล", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 1 } },
    { id: "1138", firstName: "วิศรุต", lastName: "พิมเพราะ", nickname: "จิ๊บ", department: "ขายและการตลาด", position: "พนักงานขายต่างจังหวัด", leaveStats: { sick: 4, personal: 0, maternity: 0, vacation: 0 } },
    { id: "1140", firstName: "รุ่งเกียรติ", lastName: "สิงห์พันธ์", nickname: "ฮ่องเต้", department: "ขายและการตลาด", position: "พนักงานขาย", leaveStats: { sick: "2 วัน 2 ชม.", personal: 0, maternity: 0, vacation: 0 } },
    
    { id: "1012", firstName: "อินทุอร", lastName: "ลีระพานิช", nickname: "อิน", department: "จัดซื้อ", position: "ผู้จัดการแผนกจัดซื้อ", leaveStats: { sick: 0, personal: "7 วัน 4 ชม.", maternity: 0, vacation: "2 วัน 1.5 ชม." } },
    { id: "1070", firstName: "ยงยุทธ", lastName: "ศักดารัตน์", nickname: "บอย", department: "จัดซื้อ", position: "Shipping Officer", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: "1 ชม. 30 น." } },
    { id: "1094", firstName: "จีรนันท์", lastName: "กรณ์จิราลักษณ์", nickname: "น้ำ", department: "จัดซื้อ", position: "จัดซื้อต่างประเทศ", leaveStats: { sick: 1, personal: "1 วัน 2 ชม.", maternity: 0, vacation: 2 } },
    { id: "1102", firstName: "ณิชาภา", lastName: "เหลียวดูดี", nickname: "สา", department: "จัดซื้อ", position: "Shipping Officer", leaveStats: { sick: 0, personal: "2 ชม. 30 น.", maternity: 0, vacation: 0 } },
    { id: "1139", firstName: "ศศิภา", lastName: "แตงโช", nickname: "ออย", department: "จัดซื้อ", position: "หัวหน้าแผนกจัดซื้อในประเทศ", leaveStats: { sick: 0, personal: 1, maternity: 0, vacation: 0 } },
    { id: "1146", firstName: "สุวดี", lastName: "ตุ้มเพ็ชร", nickname: "หญิง", department: "จัดซื้อ", position: "เจ้าหน้าที่จัดซื้อ", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 0 } },

    { id: "1009", firstName: "สุชาติ", lastName: "นนทฤทธิ์", nickname: "ป๊อก", department: "บัญชีและการเงิน", position: "พนักงานส่งเอกสารบัญชี", leaveStats: { sick: 25, personal: 0, maternity: 0, vacation: 2 } },
    { id: "1031", firstName: "เบญจมาศ", lastName: "ประยูรวงษ์", nickname: "เบญ", department: "บัญชีและการเงิน", position: "พนักงานบัญชี", leaveStats: { sick: 0, personal: 0, maternity: 80, vacation: 0 } },
    { id: "1107", firstName: "พรกนก", lastName: "สายเปา", nickname: "ดาว", department: "บัญชีและการเงิน", position: "พนักงานบัญชี", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 0 } },
    { id: "1109", firstName: "สุภาวดี", lastName: "จ้างไธสง", nickname: "ดาว (ผจก.)", department: "บัญชีและการเงิน", position: "ผู้จัดการบัญชีการเงิน", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: "2 วัน 30 น." } },
    { id: "1112", firstName: "นฤมล", lastName: "หมื่นสุนทร", nickname: "น้ำฝน", department: "บัญชีและการเงิน", position: "พนักงานบัญชี", leaveStats: { sick: 0, personal: 1, maternity: 0, vacation: 4 } },
    { id: "1114", firstName: "เมธินี", lastName: "โหมดเถื่อน", nickname: "ใบเตย", department: "บัญชีและการเงิน", position: "พนักงานบัญชี", leaveStats: { sick: 1, personal: 0, maternity: 0, vacation: "30 นาที" } },

    { id: "1060", firstName: "ชินาภา", lastName: "จันทวงศ์", nickname: "กุ้ง", department: "QMR/DCC/QC", position: "เจ้าหน้าที่ Support เอกสาร", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 0 } },
    { id: "1019", firstName: "ธนษา", lastName: "ธารารัตนภูดล", nickname: "เอ็ม", department: "QMR/DCC/QC", position: "ผู้จัดการ QMR /คลังสินค้าและขนส่ง/ HR", leaveStats: { sick: "4 ชม.", personal: 0, maternity: 0, vacation: "1 วัน 6 ชม." } },
    { id: "1058", firstName: "ธัญพร", lastName: "ศรีพิสุทธิ์", nickname: "ตาล", department: "QMR/DCC/QC", position: "เจ้าหน้าที่ควบคุมคุณภาพ", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 1 } },
    { id: "1088", firstName: "วิทยา", lastName: "ปักษาสุข", nickname: "ตี๋", department: "QMR/DCC/QC", position: "เช็คเกอร์", leaveStats: { sick: "2 ชม.", personal: 0, maternity: 0, vacation: 0 } },
    { id: "1096", firstName: "กันยณัฐ", lastName: "น้ำทอง", nickname: "อัญ", department: "QMR/DCC/QC", position: "Document Control", leaveStats: { sick: 2, personal: 0, maternity: 0, vacation: "1 วัน 7 ชม. 30 น." } },
    { id: "1125", firstName: "ศดานันท์", lastName: "จินดา", nickname: "แน็ต", department: "QMR/DCC/QC", position: "เช็คเกอร์ (ออกระหว่างปี)", leaveStats: { sick: 2, personal: "1 วัน 4ชม.", maternity: 0, vacation: 0 } },
    { id: "1143", firstName: "ปุญญาภา", lastName: "สร้อยทอง", nickname: "ตูน", department: "QMR/DCC/QC", position: "เช็คเกอร์", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 2 } },
    
    { id: "1108", firstName: "Khin Naing", lastName: "Htwe", nickname: "ทวย", department: "คลังสินค้าและขนส่ง", position: "พนักงานยกของ", leaveStats: { sick: 0, personal: 0, maternity: 0, vacation: 1 } }
];
