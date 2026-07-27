// ============================================================
// Word Add-in: Tiện ích Văn bản Tiếng Việt
// Version: 1.1.0 - Fixed syntax errors, scope support, window bindings
// ============================================================

// Đảm bảo Office.js khởi tạo hoàn tất
Office.onReady(function(info) {
    console.log("Office ready info:", info);
    if (info.host === Office.HostType.Word) {
        loadAbbrevRules();
        console.log("Add-in initialized successfully.");
    }
});

// ============================================================
// HELPER: Lấy range theo scope
// ============================================================
function getRange(context, scopeName) {
    var scopeEl = document.querySelector('input[name="' + scopeName + '"]:checked');
    var value = scopeEl ? scopeEl.value : 'all';
    if (value === 'selection') {
        return context.document.getSelection();
    }
    return context.document.body;
}

// ============================================================
// CHUYỂN TAB
// ============================================================
function switchTab(tabId) {
    // Bỏ active tất cả tab buttons
    document.querySelectorAll('.tab-btn').forEach(function(btn) {
        btn.classList.remove('active');
    });
    // Bỏ active tất cả tab content
    document.querySelectorAll('.tab-content').forEach(function(content) {
        content.classList.remove('active');
    });

    // Đánh dấu active tab content
    var activeTabObj = document.getElementById(tabId);
    if (activeTabObj) activeTabObj.classList.add('active');

    // Đánh dấu active tab button tương ứng
    var btns = document.querySelectorAll('.tab-btn');
    btns.forEach(function(btn) {
        if (btn.getAttribute('data-tab') === tabId) {
            btn.classList.add('active');
        }
    });
}

// ============================================================
// 1. TÍNH NĂNG CHỈNH SỬA ĐỊNH DẠNG NGÀY THÁNG
// ============================================================
async function formatDates() {
    var statusDiv = document.getElementById('dateStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = 'Đang xử lý...';

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'dateScope');
            range.load('text');
            await context.sync();

            var text = range.text || '';
            var replaceCount = 0;

            // Thứ tự pattern quan trọng: dài nhất trước
            var patterns = [
                {
                    reg: /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/gi,
                    rep: function(m, d, mth, y) { return parseInt(d) + '/' + parseInt(mth) + '/' + y; }
                },
                {
                    reg: /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})/gi,
                    rep: function(m, d, mth) { return 'ngày ' + parseInt(d) + '/' + parseInt(mth); }
                },
                {
                    reg: /tháng\s+(\d{1,2})\s+năm\s+(\d{4})/gi,
                    rep: function(m, mth, y) { return 'T' + parseInt(mth) + '/' + y; }
                },
                {
                    reg: /tháng\s+(\d{1,2})/gi,
                    rep: function(m, mth) { return 'T' + parseInt(mth); }
                }
            ];

            for (var pi = 0; pi < patterns.length; pi++) {
                var p = patterns[pi];
                var matches = [];
                var match;
                while ((match = p.reg.exec(text)) !== null) {
                    matches.push({ src: match[0], dest: p.rep.apply(null, match) });
                }

                for (var mi = 0; mi < matches.length; mi++) {
                    var item = matches[mi];
                    var results = range.search(item.src, { matchCase: false });
                    results.load('items');
                    await context.sync();
                    for (var i = 0; i < results.items.length; i++) {
                        results.items[i].insertText(item.dest, Word.InsertLocation.replace);
                        replaceCount++;
                    }
                }

                // Reload text sau mỗi pattern vì nội dung đã thay đổi
                range.load('text');
                await context.sync();
                text = range.text || '';
            }

            await context.sync();
            statusDiv.className = 'status-msg success';
            statusDiv.innerText = 'Đã chuyển đổi thành công ' + replaceCount + ' vị trí!';
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = 'Lỗi: ' + (err.message || err);
    }
}

// ============================================================
// 2. TÍNH NĂNG THAY THẾ TỪ VIẾT TẮT
// ============================================================
var abbrevRules = [
    { full: "Cộng hòa Xã hội Chủ nghĩa Việt Nam", short: "CHXHCNVN" },
    { full: "Thành phố Hồ Chí Minh", short: "TP.HCM" },
    { full: "Thành phố", short: "TP." },
    { full: "Ủy ban nhân dân", short: "UBND" },
    { full: "Hội đồng nhân dân", short: "HĐND" },
    { full: "Trách nhiệm hữu hạn", short: "TNHH" },
    { full: "Bộ Tài chính", short: "BTC" },
    { full: "Ngân hàng Nhà nước", short: "NHNN" }
];

function loadAbbrevRules() {
    try {
        var saved = localStorage.getItem('abbrevRules');
        if (saved) {
            abbrevRules = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('Cannot load from localStorage:', e);
    }
    renderAbbrevTable();
}

function saveAbbrevRules() {
    try {
        localStorage.setItem('abbrevRules', JSON.stringify(abbrevRules));
    } catch (e) {
        console.warn('Cannot save to localStorage:', e);
    }
    renderAbbrevTable();
}

function renderAbbrevTable() {
    var tbody = document.getElementById('abbrevList');
    if (!tbody) return;
    tbody.innerHTML = '';

    abbrevRules.forEach(function(rule, idx) {
        var tr = document.createElement('tr');
        tr.innerHTML =
            '<td>' + rule.full + '</td>' +
            '<td><strong>' + rule.short + '</strong></td>' +
            '<td><button class="btn-del" onclick="removeAbbrevRule(' + idx + ')">Xóa</button></td>';
        tbody.appendChild(tr);
    });
}

function addAbbrevRule() {
    var fullInput = document.getElementById('fullTerm');
    var shortInput = document.getElementById('shortTerm');

    var full = fullInput.value.trim();
    var short_ = shortInput.value.trim();

    if (!full || !short_) {
        alert("Vui lòng nhập đầy đủ thông tin!");
        return;
    }

    abbrevRules.push({ full: full, short: short_ });
    saveAbbrevRules();

    fullInput.value = '';
    shortInput.value = '';
}

function removeAbbrevRule(idx) {
    abbrevRules.splice(idx, 1);
    saveAbbrevRules();
}

async function replaceAbbreviations() {
    var statusDiv = document.getElementById('abbrevStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = 'Đang thay thế từ viết tắt...';

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'abbrevScope');
            var totalReplaced = 0;

            for (var ri = 0; ri < abbrevRules.length; ri++) {
                var rule = abbrevRules[ri];
                var searchResults = range.search(rule.full, { matchCase: false, matchWholeWord: false });
                searchResults.load('items');
                await context.sync();

                for (var i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].insertText(rule.short, Word.InsertLocation.replace);
                    totalReplaced++;
                }
            }

            await context.sync();
            statusDiv.className = 'status-msg success';
            statusDiv.innerText = 'Đã thay thế ' + totalReplaced + ' từ viết tắt!';
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = 'Lỗi xử lý: ' + err.message;
    }
}

// ============================================================
// 3. TÍNH NĂNG RÀ SOÁT CHÍNH TẢ TIẾNG VIỆT
// ============================================================
var commonVietnameseErrors = [
    // Phụ âm x/s
    { err: 'xản xuất', fix: 'sản xuất', desc: "Sai 'x' → 's'" },
    { err: 'sơ xuất', fix: 'sơ suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'đột suất', fix: 'đột xuất', desc: "Sai 'suất' → 'xuất'" },
    { err: 'xác xuất', fix: 'xác suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'năng xuất', fix: 'năng suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'suất sắc', fix: 'xuất sắc', desc: "Sai 'suất' → 'xuất'" },
    { err: 'xuất xắc', fix: 'xuất sắc', desc: "Sai 'xắc' → 'sắc'" },
    { err: 'suất xắc', fix: 'xuất sắc', desc: "Sai 'suất xắc' → 'xuất sắc'" },

    // Phụ âm ch/tr, th
    { err: 'chuẩn đoán', fix: 'chẩn đoán', desc: "Sai 'chuẩn' → 'chẩn'" },
    { err: 'thăm quan', fix: 'tham quan', desc: "Sai 'thăm' → 'tham'" },
    { err: 'chỉnh chu', fix: 'chỉn chu', desc: "Sai 'chỉnh' → 'chỉn'" },
    { err: 'trút kinh nghiệm', fix: 'rút kinh nghiệm', desc: "Sai 'trút' → 'rút'" },

    // Dấu hỏi / ngã
    { err: 'suôn sẽ', fix: 'suôn sẻ', desc: "Sai dấu ngã → hỏi" },
    { err: 'cũng cố', fix: 'củng cố', desc: "Sai dấu ngã → hỏi" },
    { err: 'dể dàng', fix: 'dễ dàng', desc: "Sai dấu hỏi → ngã" },
    { err: 'bổ xung', fix: 'bổ sung', desc: "Sai 'xung' → 'sung'" },

    // Từ ghép sai phổ biến
    { err: 'vô hình chung', fix: 'vô hình trung', desc: "Sai 'chung' → 'trung'" },
    { err: 'tựu chung', fix: 'tựu trung', desc: "Sai 'chung' → 'trung'" },
    { err: 'chín mùi', fix: 'chín muồi', desc: "Sai 'mùi' → 'muồi'" },
    { err: 'khoảng khắc', fix: 'khoảnh khắc', desc: "Sai 'khoảng' → 'khoảnh'" }
];

async function checkSpelling() {
    var resultsDiv = document.getElementById('spellResults');
    resultsDiv.innerHTML = '<div class="status-msg">Đang rà soát chính tả...</div>';

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'spellScope');
            range.load('text');
            await context.sync();

            var text = range.text || '';
            var findings = [];

            for (var i = 0; i < commonVietnameseErrors.length; i++) {
                var item = commonVietnameseErrors[i];
                var regex = new RegExp(escapeRegex(item.err), 'gi');
                var match;
                while ((match = regex.exec(text)) !== null) {
                    findings.push({
                        original: match[0],
                        correct: item.fix,
                        desc: item.desc
                    });
                }
            }

            renderSpellFindings(findings);
        });
    } catch (err) {
        console.error(err);
        resultsDiv.innerHTML = '<div class="status-msg error">Lỗi rà soát: ' + err.message + '</div>';
    }
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function renderSpellFindings(findings) {
    var resultsDiv = document.getElementById('spellResults');
    if (findings.length === 0) {
        resultsDiv.innerHTML = '<div class="status-msg success">🎉 Không phát hiện lỗi chính tả phổ biến nào!</div>';
        return;
    }

    var html = '<div class="status-msg error">Phát hiện ' + findings.length + ' nghi vấn chính tả:</div>';
    html += '<button class="btn btn-primary" style="margin-top:6px;margin-bottom:8px" onclick="fixAllSpelling()">🔧 Sửa tất cả ' + findings.length + ' lỗi</button>';

    findings.forEach(function(item) {
        html += '<div class="spell-item">';
        html += '<div>Từ phát hiện: <span class="original">' + item.original + '</span></div>';
        html += '<div>Gợi ý sửa: <span class="suggestion">' + item.correct + '</span>';
        if (item.desc) html += ' (' + item.desc + ')';
        html += '</div>';
        html += '<div class="spell-actions">';
        html += "<button class=\"btn-sm\" onclick=\"fixSingleSpell('" + item.original.replace(/'/g, "\\'") + "', '" + item.correct.replace(/'/g, "\\'") + "')\">Sửa lỗi này</button>";
        html += '</div></div>';
    });

    resultsDiv.innerHTML = html;
}

async function fixSingleSpell(original, correct) {
    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'spellScope');
            var searchResults = range.search(original, { matchCase: false, matchWholeWord: false });
            searchResults.load('items');
            await context.sync();

            for (var i = 0; i < searchResults.items.length; i++) {
                searchResults.items[i].insertText(correct, Word.InsertLocation.replace);
            }
            await context.sync();

            // Quét lại sau khi sửa
            await checkSpelling();
        });
    } catch (err) {
        console.error(err);
        alert('Lỗi sửa chính tả: ' + (err.message || err));
    }
}

async function fixAllSpelling() {
    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'spellScope');
            var count = 0;

            for (var i = 0; i < commonVietnameseErrors.length; i++) {
                var item = commonVietnameseErrors[i];
                var searchResults = range.search(item.err, { matchCase: false, matchWholeWord: false });
                searchResults.load('items');
                await context.sync();

                for (var j = 0; j < searchResults.items.length; j++) {
                    searchResults.items[j].insertText(item.fix, Word.InsertLocation.replace);
                    count++;
                }
            }

            await context.sync();
            var resultsDiv = document.getElementById('spellResults');
            resultsDiv.innerHTML = '<div class="status-msg success">🎉 Đã tự động sửa thành công ' + count + ' lỗi chính tả!</div>';
        });
    } catch (err) {
        console.error(err);
        alert("Lỗi: " + (err.message || err));
    }
}

// ============================================================
// GẮN HÀM LÊN WINDOW ĐỂ HTML ONCLICK GỌI ĐƯỢC
// ============================================================
window.switchTab = switchTab;
window.formatDates = formatDates;
window.addAbbrevRule = addAbbrevRule;
window.removeAbbrevRule = removeAbbrevRule;
window.replaceAbbreviations = replaceAbbreviations;
window.checkSpelling = checkSpelling;
window.fixSingleSpell = fixSingleSpell;
window.fixAllSpelling = fixAllSpelling;
