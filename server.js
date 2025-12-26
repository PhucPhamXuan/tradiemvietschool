const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Cấu hình upload file - Lưu ngay tại thư mục gốc
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, __dirname); // Lưu vào thư mục hiện tại (root)
    },
    filename: function (req, file, cb) {
        cb(null, 'data.xlsx'); // Ghi đè tên file thành data.xlsx
    }
});
const upload = multer({ storage: storage });

// Middleware
app.use(express.json());
// Không dùng express.static('public') nữa vì không còn folder public

// --- HELPER FUNCTIONS ---

const normalizeHeader = (header) => {
    if (!header) return null;
    const clean = header.toString().trim().toLowerCase();
    if (clean.includes('toan') || clean.includes('toán')) return 'Toán';
    if (clean.includes('van') || clean.includes('văn')) return 'Văn';
    if (clean.includes('anh') || clean.includes('tiếng anh')) return 'Anh';
    if (clean.includes('ly') || clean.includes('lý') || clean.includes('vat li')) return 'Lý';
    if (clean.includes('hoa') || clean.includes('hóa')) return 'Hóa';
    if (clean.includes('sinh')) return 'Sinh';
    if (clean.includes('su') || clean.includes('sử') || clean.includes('lich su')) return 'Sử';
    if (clean.includes('dia') || clean.includes('địa')) return 'Địa';
    if (clean.includes('gdcd') || clean.includes('ktpl') || clean.includes('phap luat')) return 'KTPL';
    if (clean.includes('tin') || clean.includes('tin hoc')) return 'Tin';
    if (clean.includes('cn') || clean.includes('cong nghe')) return 'CN';
    return null;
};

const readExcelData = () => {
    const filePath = path.join(__dirname, 'data.xlsx');
    
    if (!fs.existsSync(filePath)) {
        return { error: "Chưa có file dữ liệu (data.xlsx)" };
    }

    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet);

        if (!jsonData || jsonData.length === 0) return { data: [], headers: [] };

        const firstRowKeys = Object.keys(jsonData[0]);
        const subjectMap = {}; 
        const subjects = [];

        firstRowKeys.forEach(key => {
            const norm = normalizeHeader(key);
            if (norm) {
                subjectMap[key] = norm;
                if (!subjects.includes(norm)) subjects.push(norm);
            }
        });

        const processedData = jsonData.map(row => {
            const scores = {};
            let totalScore = 0;
            let subjectCount = 0;

            const findValue = (keywords) => {
                const key = firstRowKeys.find(k => keywords.some(kw => k.toLowerCase().includes(kw)));
                return key ? row[key] : '';
            };

            const sbd = findValue(['sbd', 'số báo danh', 'so bao danh']) || row['SBD'] || '';
            const name = findValue(['họ tên', 'ho ten', 'hoten', 'tên']) || row['HoTen'] || '';
            const lop = findValue(['lớp', 'lop']) || row['Lop'] || '';

            Object.keys(subjectMap).forEach(rawKey => {
                const subjectName = subjectMap[rawKey];
                const val = row[rawKey];
                if (val !== undefined && val !== null && val !== '' && !isNaN(parseFloat(val))) {
                    const numVal = parseFloat(val);
                    scores[subjectName] = numVal;
                    totalScore += numVal;
                    subjectCount++;
                } else {
                    scores[subjectName] = null;
                }
            });

            return {
                SBD: sbd, HoTen: name, Lop: lop, _scores: scores,
                DTB: subjectCount > 0 ? (totalScore / subjectCount).toFixed(2) : 0
            };
        });

        return { data: processedData, headers: subjects };
    } catch (error) {
        console.error("Lỗi đọc file:", error);
        return { error: "File data.xlsx bị lỗi format." };
    }
};

// --- API ENDPOINTS ---

app.get('/api/data', (req, res) => {
    const result = readExcelData();
    if (result.error) return res.status(404).json(result);
    res.json(result);
});

app.post('/api/upload', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('Không có file.');
    res.send('File uploaded successfully');
});

// Serve Frontend (Trỏ trực tiếp vào file index.html cùng cấp)
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
