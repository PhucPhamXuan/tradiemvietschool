const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Helper: Lấy tên file dựa trên khối
const getFileNameByGrade = (grade) => {
    if (grade == 11) return 'data1.xlsx';
    if (grade == 10) return 'data0.xlsx';
    return 'data.xlsx'; // Mặc định khối 12
};

app.use(express.json());

// --- LOGIC MỚI: CHUẨN HÓA VÀ QUÉT CỘT ---

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

const readExcelData = (grade) => {
    const fileName = getFileNameByGrade(grade);
    const filePath = path.join(__dirname, fileName);
    
    if (!fs.existsSync(filePath)) {
        return { error: `Chưa có file dữ liệu cho khối ${grade} (${fileName}) trên hệ thống.` };
    }

    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const rawData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        if (!rawData || rawData.length < 2) return { data: [], headers: [] };

        const headerRow = rawData[0]; 
        const colIndexMap = {}; 
        const subjects = [];

        headerRow.forEach((rawHeader, index) => {
            const norm = normalizeHeader(rawHeader);
            if (norm) {
                colIndexMap[index] = norm;
                if (!subjects.includes(norm)) subjects.push(norm);
            }
        });

        const processedData = [];
        
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

            subjects.forEach(sub => scores[sub] = null);

            Object.keys(colIndexMap).forEach(colIdx => {
                const subjectName = colIndexMap[colIdx];
                const val = row[colIdx];
                
                if (val !== undefined && val !== null && val !== '' && !isNaN(parseFloat(val))) {
                    const numVal = parseFloat(val);
                    scores[subjectName] = numVal;
                    totalScore += numVal;
                    subjectCount++;
                }
            });

            const sbd = idxSBD !== -1 ? row[idxSBD] : '';
            const name = idxName !== -1 ? row[idxName] : '';
            const lop = idxLop !== -1 ? row[idxLop] : '';

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
        return { error: "File hệ thống bị lỗi format." };
    }
};

// --- API ENDPOINTS ---

// API lấy dữ liệu theo khối
app.get('/api/data', (req, res) => {
    const grade = req.query.grade || '12'; // Lấy param grade
    const result = readExcelData(grade);
    if (result.error) return res.status(404).json(result);
    res.json(result);
});

// ĐÃ XÓA API POST /api/upload

app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
