const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Cấu hình upload file
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, __dirname);
    },
    filename: function (req, file, cb) {
        cb(null, 'data.xlsx');
    }
});
const upload = multer({ storage: storage });

app.use(express.json());

// --- LOGIC MỚI: CHUẨN HÓA VÀ QUÉT CỘT ---

const normalizeHeader = (header) => {
    if (!header) return null;
    const clean = header.toString().trim().toLowerCase();
    // Logic nhận diện từ khóa (đã bao gồm Sinh, KTPL, GDCD)
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
        
        // Dùng header: 1 để lấy mảng mảng (array of arrays) thay vì JSON object ngay
        // Điều này giúp ta lấy được chính xác dòng tiêu đề gốc
        const rawData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        if (!rawData || rawData.length < 2) return { data: [], headers: [] };

        // Dòng 0 là tiêu đề gốc
        const headerRow = rawData[0]; 
        
        // 1. Tạo Map để mapping từ Index cột -> Tên chuẩn hóa (Ví dụ: cột 3 -> 'Toán')
        const colIndexMap = {}; 
        const subjects = [];

        headerRow.forEach((rawHeader, index) => {
            const norm = normalizeHeader(rawHeader);
            if (norm) {
                colIndexMap[index] = norm; // Lưu lại index của cột môn học
                if (!subjects.includes(norm)) subjects.push(norm); // Danh sách môn duy nhất
            }
        });

        // 2. Duyệt qua các dòng dữ liệu (từ dòng 1 trở đi)
        const processedData = [];
        
        // Helper tìm index của các cột thông tin (SBD, Tên, Lớp)
        const findColIndex = (keywords) => {
            return headerRow.findIndex(h => h && keywords.some(kw => h.toString().toLowerCase().includes(kw)));
        };

        const idxSBD = findColIndex(['sbd', 'số báo danh', 'so bao danh']);
        const idxName = findColIndex(['họ tên', 'ho ten', 'hoten', 'tên']);
        const idxLop = findColIndex(['lớp', 'lop']);

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;

            const scores = {};
            let totalScore = 0;
            let subjectCount = 0;

            // Lấy điểm dựa trên index đã map
            subjects.forEach(sub => {
                // Tìm index cột gốc tương ứng với môn này (trong map)
                // Lưu ý: Có thể có trường hợp file excel 1 môn có 2 cột (ít gặp), code này lấy cột cuối cùng khớp.
                // Để chính xác: Duyệt qua keys của colIndexMap
                scores[sub] = null; // Mặc định là null
            });

            Object.keys(colIndexMap).forEach(colIdx => {
                const subjectName = colIndexMap[colIdx];
                const val = row[colIdx];
                
                if (val !== undefined && val !== null && val !== '' && !isNaN(parseFloat(val))) {
                    const numVal = parseFloat(val);
                    scores[subjectName] = numVal;
                    // Logic tính điểm TB (chỉ tính môn có điểm)
                    totalScore += numVal;
                    subjectCount++;
                }
            });

            // Lấy thông tin
            const sbd = idxSBD !== -1 ? row[idxSBD] : '';
            const name = idxName !== -1 ? row[idxName] : '';
            const lop = idxLop !== -1 ? row[idxLop] : '';

            // Chỉ push nếu có thông tin cơ bản
            if (name || sbd) {
                processedData.push({
                    SBD: sbd,
                    HoTen: name,
                    Lop: lop,
                    _scores: scores,
                    DTB: subjectCount > 0 ? (totalScore / subjectCount).toFixed(2) : 0
                });
            }
        }

        return { data: processedData, headers: subjects };

    } catch (error) {
        console.error("Lỗi đọc file:", error);
        return { error: "File data.xlsx bị lỗi format hoặc đang được mở." };
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

app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
