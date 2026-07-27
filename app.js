// Đảm bảo Office.js khởi tạo hoàn tất
Office.onReady((info) => {
    console.log("Office ready info:", info);
    if (info.host === Office.HostType.Word) {
        loadAbbrevRules();
    }
});

// Chuyển Tab
function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    const activeTabObj = document.getElementById(tabId);
    if (activeTabObj) activeTabObj.classList.add('active');

    if (window.event && window.event.currentTarget) {
        window.event.currentTarget.classList.add('active');
    }
}

// -------------------------------------------------------------
// 1. TÍNH NĂNG CHỈNH SỬA ĐỊNH DẠNG NGÀY THÁNG
// -------------------------------------------------------------
async function formatDates() {
    const statusDiv = document.getElementById('dateStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = 'Đang xử lý...';

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();

            let fullText = body.text || "";
            let replaceCount = 0;

            // Mẫu 1: ngày DD tháng MM năm YYYY
            const regex1 = /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/gi;
            let m;
            let list1 = [];
            while ((m = regex1.exec(fullText)) !== null) {
                list1.push({ src: m[0], rep: `${m[1]}/${m[2]}/${m[3]}` });
            }

            for (const item of list1) {
                const results = body.search(item.src, { matchCase: false });
                results.load("items");
                await context.sync();

                for (let i = 0; i < results.items.length; i++) {
                    results.items[i].insertText(item.rep, Word.InsertLocation.replace);
                    replaceCount++;
                }
            }

            // Mẫu 2: ngày DD tháng MM
            body.load("text");
            await context.sync();
            fullText = body.text || "";
            const regex2 = /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})/gi;
            let list2 = [];
            while ((m = regex2.exec(fullText)) !== null) {
                list2.push({ src: m[0], rep: `ngày ${m[1]}/${m[2]}` });
            }
            for (const item of list2) {
                const results = body.search(item.src, { matchCase: false });
                results.load("items");
                await context.sync();

                for (let i = 0; i < results.items.length; i++) {
                    results.items[i].insertText(item.rep, Word.InsertLocation.replace);
                    replaceCount++;
                }
            }

            // Mẫu 3: tháng MM năm YYYY
            body.load("text");
            await context.sync();
            fullText = body.text || "";
            const regex3 = /tháng\s+(\d{1,2})\s+năm\s+(\d{4})/gi;
            let list3 = [];
            while ((m = regex3.exec(fullText)) !== null) {
                list3.push({ src: m[0], rep: `T${m[1]}/${m[2]}` });
            }
            for (const item of list3) {
                const results = body.search(item.src, { matchCase: false });
                results.load("items");
                await context.sync();

                for (let i = 0; i < results.items.length; i++) {
                    results.items[i].insertText(item.rep, Word.InsertLocation.replace);
                    replaceCount++;
                }
            }

            // Mẫu 4: tháng MM
            body.load("text");
            await context.sync();
            fullText = body.text || "";
            const regex4 = /tháng\s+(\d{1,2})/gi;
            let list4 = [];
            while ((m = regex4.exec(fullText)) !== null) {
                list4.push({ src: m[0], rep: `T${m[1]}` });
            }
            for (const item of list4) {
                const results = body.search(item.src, { matchCase: false });
                results.load("items");
                await context.sync();

                for (let i = 0; i < results.items.length; i++) {
                    results.items[i].insertText(item.rep, Word.InsertLocation.replace);
                    replaceCount++;
                }
            }

            await context.sync();
            statusDiv.className = 'status-msg success';
            statusDiv.innerText = `Đã chuyển đổi thành công ${replaceCount} vị trí!`;
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = 'Lỗi: ' + (err.message || err);
    }
}




// Helper thay thế bằng Regex trong Word
async function replaceWithRegex(context, parentRange, regex, replacementFn) {
    parentRange.load('text');
    await context.sync();

    let matches = [];
    let match;
    const fullText = parentRange.text || "";
    
    while ((match = regex.exec(fullText)) !== null) {
        matches.push({
            text: match[0],
            replacement: replacementFn(...match)
        });
    }

    let replaceCount = 0;
    for (const item of matches) {
        const searchResults = parentRange.search(item.text, { matchCase: false, matchWholeWord: false });
        searchResults.load('items');
        await context.sync();

        for (let i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].insertText(item.replacement, Word.InsertLocation.replace);
            replaceCount++;
        }
    }
    return replaceCount;
}


// -------------------------------------------------------------
// 2. TÍNH NĂNG THAY THẾ TỪ VIẾT TẮT
// -------------------------------------------------------------
let abbrevRules = [
    { full: "Cộng hòa Xã hội Chủ nghĩa Việt Nam", short: "CHXHCNVN" },
    { full: "Thành phố", short: "TP." },
    { full: "Ủy ban nhân dân", short: "UBND" },
    { full: "Hội đồng nhân dân", short: "HĐND" },
    { full: "Trách nhiệm hữu hạn", short: "TNHH" }
];

function loadAbbrevRules() {
    const saved = localStorage.getItem('abbrevRules');
    if (saved) {
        try { abbrevRules = JSON.parse(saved); } catch(e){}
    }
    renderAbbrevTable();
}

function saveAbbrevRules() {
    localStorage.setItem('abbrevRules', JSON.stringify(abbrevRules));
    renderAbbrevTable();
}

function renderAbbrevTable() {
    const tbody = document.getElementById('abbrevList');
    tbody.innerHTML = '';

    abbrevRules.forEach((rule, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${rule.full}</td>
            <td><strong>${rule.short}</strong></td>
            <td><button class="btn-del" onclick="removeAbbrevRule(${idx})">Xóa</button></td>
        `;
        tbody.appendChild(tr);
    });
}

function addAbbrevRule() {
    const fullInput = document.getElementById('fullTerm');
    const shortInput = document.getElementById('shortTerm');

    const full = fullInput.value.trim();
    const short = shortInput.value.trim();

    if (!full || !short) {
        alert("Vui lòng nhập đầy đủ thông tin!");
        return;
    }

    abbrevRules.push({ full, short });
    saveAbbrevRules();

    fullInput.value = '';
    shortInput.value = '';
}

function removeAbbrevRule(idx) {
    abbrevRules.splice(idx, 1);
    saveAbbrevRules();
}

async function replaceAbbreviations() {
    const statusDiv = document.getElementById('abbrevStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = 'Đang thay thế từ viết tắt...';

    const scope = document.querySelector('input[name="abbrevScope"]:checked').value;

    try {
        await Word.run(async (context) => {
            let range = scope === 'selection' ? context.document.getSelection() : context.document.body;
            let totalReplaced = 0;

            for (const rule of abbrevRules) {
                const searchResults = range.search(rule.full, { matchCase: false, matchWholeWord: false });
                searchResults.load('items');
                await context.sync();

                for (let i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].insertText(rule.short, Word.InsertLocation.replace);
                    totalReplaced++;
                }
            }

            statusDiv.className = 'status-msg success';
            statusDiv.innerText = `Đã thay thế ${totalReplaced} từ viết tắt!`;
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = 'Lỗi xử lý: ' + err.message;
    }
}

// -------------------------------------------------------------
// 3. TÍNH NĂNG RÀ SOÁT CHÍNH TẢ TIẾNG VIỆT
// -------------------------------------------------------------
const commonVietnameseErrors = [
    // Phụ âm đầu / Chính tả từ ghép phổ biến
    { error: /\bsản xuất\b/gi, correct: "sản xuất", desc: "Đúng chính tả" },
    { error: /\bxản xuất\b/gi, correct: "sản xuất", desc: "Sai chính tả 'x/s'" },
    { error: /\bsơ xuất\b/gi, correct: "sơ suất", desc: "Sai chính tả 'xuất/suất'" },
    { error: /\bsơ xuất\b/gi, correct: "sơ suất", desc: "Sai chính tả 'xuất/suất'" },
    { error: /\bchỉn chu\b/gi, correct: "chỉn chu", desc: "Đúng chính tả" },
    { error: /\bchỉnh chu\b/gi, correct: "chỉn chu", desc: "Sai chính tả 'chỉnh chu -> chỉn chu'" },
    { error: /\bsuôn sẻ\b/gi, correct: "suôn sẻ", desc: "Đúng chính tả" },
    { error: /\bsuôn sẽ\b/gi, correct: "suôn sẻ", desc: "Sai dấu hỏi/nã 'sẽ -> sẻ'" },
    { error: /\bthấu đáo\b/gi, correct: "thấu đáo", desc: "Đúng chính tả" },
    { error: /\bđáo để\b/gi, correct: "đáo để", desc: "Đúng chính tả" },
    { error: /\bđã đành\b/gi, correct: "đã đành", desc: "Đúng chính tả" },
    { error: /\bcụm từ\b/gi, correct: "cụm từ", desc: "Đúng" },
    { error: /\bcủng cố\b/gi, correct: "củng cố", desc: "Đúng" },
    { error: /\bcũng cố\b/gi, correct: "củng cố", desc: "Sai dấu hỏi/ngã 'cũng -> củng'" },
    { error: /\bđột xuất\b/gi, correct: "đột xuất", desc: "Đúng" },
    { error: /\bđột suất\b/gi, correct: "đột xuất", desc: "Sai chính tả 'suất -> xuất'" },
    { error: /\bxác suất\b/gi, correct: "xác suất", desc: "Đúng" },
    { error: /\bxác xuất\b/gi, correct: "xác suất", desc: "Sai chính tả 'xuất -> suất'" },
    { error: /\bnăng suất\b/gi, correct: "năng suất", desc: "Đúng" },
    { error: /\bnăng xuất\b/gi, correct: "năng suất", desc: "Sai chính tả 'xuất -> suất'" },
    { error: /\btham quan\b/gi, correct: "tham quan", desc: "Đúng" },
    { error: /\btham quan\b/gi, correct: "tham quan", desc: "Đúng" },
    { error: /\bthăm quan\b/gi, correct: "tham quan", desc: "Sai chính tả 'thăm quan -> tham quan'" },
    { error: /\bchẩn đoán\b/gi, correct: "chẩn đoán", desc: "Đúng" },
    { error: /\bchuẩn đoán\b/gi, correct: "chẩn đoán", desc: "Sai chính tả 'chuẩn -> chẩn'" },
    { error: /\btrút bỏ\b/gi, correct: "trút bỏ", desc: "Đúng" },
    { error: /\btrút giận\b/gi, correct: "trút giận", desc: "Đúng" },
    { error: /\brút kinh nghiệm\b/gi, correct: "rút kinh nghiệm", desc: "Đúng" },
    { error: /\btrút kinh nghiệm\b/gi, correct: "rút kinh nghiệm", desc: "Sai chính tả 'trút -> rút'" },
    { error: /\bxuất sắc\b/gi, correct: "xuất sắc", desc: "Đúng" },
    { error: /\bsuất sắc\b/gi, correct: "xuất sắc", desc: "Sai chính tả 'suất -> xuất'" },
    { error: /\bxuất xắc\b/gi, correct: "xuất sắc", desc: "Sai chính tả 'xắc -> sắc'" },
    { error: /\bsuất xắc\b/gi, correct: "xuất sắc", desc: "Sai chính tả 'suất xắc -> xuất sắc'" }
];

async function checkSpelling() {
    const resultsDiv = document.getElementById('spellResults');
    resultsDiv.innerHTML = '<div class="status-msg">Đang rà soát chính tả...</div>';

    const scope = document.querySelector('input[name="spellScope"]:checked').value;

    try {
        await Word.run(async (context) => {
            let range = scope === 'selection' ? context.document.getSelection() : context.document.body;
            range.load('text');
            await context.sync();

            const text = range.text;
            let findings = [];

            // Quét các quy tắc sai chính tả
            for (const item of commonVietnameseErrors) {
                // Chỉ lấy những quy tắc là lỗi sai
                if (item.error.source.includes('chuẩn đoán') || 
                    item.error.source.includes('thăm quan') ||
                    item.error.source.includes('chỉnh chu') ||
                    item.error.source.includes('sơ xuất') ||
                    item.error.source.includes('suôn sẽ') ||
                    item.error.source.includes('cũng cố') ||
                    item.error.source.includes('đột suất') ||
                    item.error.source.includes('xác xuất') ||
                    item.error.source.includes('năng xuất') ||
                    item.error.source.includes('trút kinh nghiệm') ||
                    item.error.source.includes('suất sắc') ||
                    item.error.source.includes('xuất xắc') ||
                    item.error.source.includes('suất xắc') ||
                    item.error.source.includes('xản xuất')) {
                    
                    let match;
                    const regex = new RegExp(item.error, 'gi');
                    while ((match = regex.exec(text)) !== null) {
                        findings.push({
                            original: match[0],
                            correct: item.correct,
                            desc: item.desc
                        });
                    }
                }
            }

            renderSpellFindings(findings);
        });
    } catch (err) {
        console.error(err);
        resultsDiv.innerHTML = '<div class="status-msg error">Lỗi rà soát: ' + err.message + '</div>';
    }
}

function renderSpellFindings(findings) {
    const resultsDiv = document.getElementById('spellResults');
    if (findings.length === 0) {
        resultsDiv.innerHTML = '<div class="status-msg success">🎉 Không phát hiện lỗi chính tả phổ biến nào!</div>';
        return;
    }

    resultsDiv.innerHTML = `<div class="status-msg error">Phát hiện ${findings.length} nghi vấn chính tả:</div>`;

    findings.forEach((item, idx) => {
        const div = document.createElement('div');
        div.className = 'spell-item';
        div.innerHTML = `
            <div>Từ phát hiện: <span class="original">${item.original}</span></div>
            <div>Gợi ý sửa: <span class="suggestion">${item.correct}</span> (${item.desc})</div>
            <div class="spell-actions">
                <button class="btn-sm" onclick="fixSingleSpell('${item.original}', '${item.correct}')">Sửa lỗi này</button>
            </div>
        `;
        resultsDiv.appendChild(div);
    });
}

async function fixSingleSpell(original, correct) {
    try {
        await Word.run(async (context) => {
            const scope = document.querySelector('input[name="spellScope"]:checked').value;
            let range = scope === 'selection' ? context.document.getSelection() : context.document.body;

            const searchResults = range.search(original, { matchCase: false });
            searchResults.load('items');
            await context.sync();

// Gắn hàm lên window để HTML gọi trực tiếp không bị lỗi scope
window.switchTab = switchTab;
window.formatDates = formatDates;
window.addAbbrevRule = addAbbrevRule;
window.removeAbbrevRule = removeAbbrevRule;
window.replaceAbbreviations = replaceAbbreviations;
window.checkSpelling = checkSpelling;
window.fixSingleSpell = fixSingleSpell;
window.fixAllSpelling = fixAllSpelling;

async function fixAllSpelling() {
    try {
        await Word.run(async (context) => {
            const scope = document.querySelector('input[name="spellScope"]:checked').value;
            let range = scope === 'selection' ? context.document.getSelection() : context.document.body;

            range.load('text');
            await context.sync();

            let count = 0;
            const text = range.text || "";

            for (const item of commonVietnameseErrors) {
                if (item.error.test(text)) {
                    const searchResults = range.search(item.error.source, { matchCase: false, matchWholeWord: false });
                    searchResults.load('items');
                    await context.sync();

                    for (let i = 0; i < searchResults.items.length; i++) {
                        searchResults.items[i].insertText(item.correct, Word.InsertLocation.replace);
                        count++;
                    }
                }
            }

            const resultsDiv = document.getElementById('spellResults');
            resultsDiv.innerHTML = `<div class="status-msg success">🎉 Đã tự động sửa thành công ${count} lỗi chính tả!</div>`;
        });
    } catch (err) {
        console.error(err);
        alert("Lỗi: " + (err.message || err));
    }
}

