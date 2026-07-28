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
                    rep: function(m, d, mth, y) { return parseInt(d, 10) + '/' + parseInt(mth, 10) + '/' + y; }
                },
                {
                    reg: /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})/gi,
                    rep: function(m, d, mth) {
                        var prefix = (m.charAt(0) === 'N') ? 'Ngày ' : 'ngày ';
                        return prefix + parseInt(d, 10) + '/' + parseInt(mth, 10);
                    }
                },
                {
                    reg: /tháng\s+(\d{1,2})\s+năm\s+(\d{4})/gi,
                    rep: function(m, mth, y) { return 'T' + parseInt(mth, 10) + '/' + y; }
                },
                {
                    reg: /tháng\s+(\d{1,2})/gi,
                    rep: function(m, mth) { return 'T' + parseInt(mth, 10); }
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

function sortAbbrevRulesAlphabetically() {
    if (abbrevRules && abbrevRules.length > 0) {
        abbrevRules.sort(function(a, b) {
            return a.full.localeCompare(b.full, 'vi', { sensitivity: 'base' });
        });
    }
}

function loadAbbrevRules() {
    try {
        var saved = localStorage.getItem('abbrevRules');
        if (saved) {
            abbrevRules = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('Cannot load from localStorage:', e);
    }
    sortAbbrevRulesAlphabetically();
    renderAbbrevTable();
}

function saveAbbrevRules() {
    sortAbbrevRulesAlphabetically();
    try {
        localStorage.setItem('abbrevRules', JSON.stringify(abbrevRules));
    } catch (e) {
        console.warn('Cannot save to localStorage:', e);
    }
    renderAbbrevTable();
}

var editingAbbrevIdx = -1;

function renderAbbrevTable() {
    var tbody = document.getElementById('abbrevList');
    if (!tbody) return;
    tbody.innerHTML = '';

    abbrevRules.forEach(function(rule, idx) {
        var tr = document.createElement('tr');
        if (idx === editingAbbrevIdx) {
            tr.innerHTML =
                '<td><input type="text" id="editFullTerm" class="input-edit" value="' + escapeHtml(rule.full) + '"></td>' +
                '<td><input type="text" id="editShortTerm" class="input-edit" value="' + escapeHtml(rule.short) + '"></td>' +
                '<td style="text-align: center;">' +
                '<button class="btn-edit" style="background:var(--success-color);" title="Lưu" onclick="saveEditAbbrevRule(' + idx + ')">💾</button>' +
                '<button class="btn-del" title="Hủy" onclick="cancelEditAbbrevRule()">✕</button>' +
                '</td>';
        } else {
            tr.innerHTML =
                '<td>' + escapeHtml(rule.full) + '</td>' +
                '<td><strong>' + escapeHtml(rule.short) + '</strong></td>' +
                '<td style="text-align: center;">' +
                '<button class="btn-edit" title="Sửa" onclick="editAbbrevRule(' + idx + ')">✏️</button>' +
                '<button class="btn-del" title="Xóa" onclick="removeAbbrevRule(' + idx + ')">✕</button>' +
                '</td>';
        }
        tbody.appendChild(tr);
    });
}

function editAbbrevRule(idx) {
    editingAbbrevIdx = idx;
    renderAbbrevTable();
}

function cancelEditAbbrevRule() {
    editingAbbrevIdx = -1;
    renderAbbrevTable();
}

function saveEditAbbrevRule(idx) {
    var fullInput = document.getElementById('editFullTerm');
    var shortInput = document.getElementById('editShortTerm');
    if (!fullInput || !shortInput) return;

    var full = fullInput.value.trim();
    var short_ = shortInput.value.trim();

    if (!full || !short_) {
        alert("Vui lòng nhập đầy đủ thông tin!");
        return;
    }

    if (idx >= 0 && idx < abbrevRules.length) {
        abbrevRules[idx] = { full: full, short: short_ };
        editingAbbrevIdx = -1;
        saveAbbrevRules();
    }
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

function exportAbbrevRules() {
    var statusDiv = document.getElementById('abbrevStatus');
    if (!abbrevRules || abbrevRules.length === 0) {
        statusDiv.className = 'status-msg status-error';
        statusDiv.innerText = '⚠️ Danh sách từ điển đang trống!';
        return;
    }
    try {
        var jsonStr = JSON.stringify(abbrevRules, null, 2);
        var blob = new Blob([jsonStr], { type: "application/json;charset=utf-8" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = "tu_dien_viet_tat.json";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        statusDiv.className = 'status-msg status-success';
        statusDiv.innerText = '✅ Đã xuất ' + abbrevRules.length + ' quy tắc ra tệp tu_dien_viet_tat.json!';
    } catch (err) {
        console.error('exportAbbrevRules error:', err);
        statusDiv.className = 'status-msg status-error';
        statusDiv.innerText = '❌ Lỗi xuất tệp: ' + (err.message || err);
    }
}

function importAbbrevRules(event) {
    var file = event.target.files[0];
    var statusDiv = document.getElementById('abbrevStatus');
    if (!file) return;

    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var importedData = JSON.parse(e.target.result);
            if (!Array.isArray(importedData)) {
                throw new Error('Cấu trúc tệp không hợp lệ (cần danh sách dạng mảng JSON).');
            }

            var addedCount = 0;
            var updatedCount = 0;

            for (var i = 0; i < importedData.length; i++) {
                var item = importedData[i];
                if (item && item.full && item.short) {
                    var f = item.full.trim();
                    var s = item.short.trim();
                    var existingIdx = -1;

                    for (var k = 0; k < abbrevRules.length; k++) {
                        if (abbrevRules[k].full.toLowerCase() === f.toLowerCase()) {
                            existingIdx = k;
                            break;
                        }
                    }

                    if (existingIdx >= 0) {
                        abbrevRules[existingIdx].short = s;
                        updatedCount++;
                    } else {
                        abbrevRules.push({ full: f, short: s });
                        addedCount++;
                    }
                }
            }

            saveAbbrevRules();

            statusDiv.className = 'status-msg status-success';
            statusDiv.innerText = '✅ Nhập thành công! (Thêm mới: ' + addedCount + ', Cập nhật: ' + updatedCount + ')';
        } catch (err) {
            console.error('importAbbrevRules error:', err);
            statusDiv.className = 'status-msg status-error';
            statusDiv.innerText = '❌ Lỗi nhập tệp: ' + (err.message || err);
        } finally {
            event.target.value = '';
        }
    };
    reader.readAsText(file);
}

async function replaceAbbreviations() {
    var statusDiv = document.getElementById('abbrevStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = 'Đang thay thế từ viết tắt...';

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'abbrevScope');
            var totalReplaced = 0;

            // Sắp xếp các cụm từ từ dài nhất đến ngắn nhất
            var sortedRules = abbrevRules.slice().sort(function(a, b) {
                return b.full.length - a.full.length;
            });

            for (var ri = 0; ri < sortedRules.length; ri++) {
                var rule = sortedRules[ri];
                var searchResults = range.search(rule.full, { matchCase: false, matchWholeWord: false });
                searchResults.load('items');
                await context.sync();

                for (var i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].insertText(rule.short, Word.InsertLocation.replace);
                    totalReplaced++;
                }
            }

            // Bước 2: Tự động loại bỏ các cụm lặp thừa trong ngoặc như "PPA (PPA)" -> "PPA" hoặc "Genco1 (Genco1)" -> "Genco1"
            range.load('text');
            await context.sync();
            var currentText = range.text || '';

            var dupRegex = /([^\s()]+(?:\s+[^\s()]+)*)\s*\(\s*\1\s*\)/gi;
            var dupMatches = [];
            var match;

            while ((match = dupRegex.exec(currentText)) !== null) {
                dupMatches.push({ fullMatch: match[0], keepText: match[1] });
            }

            var cleanCount = 0;
            for (var di = 0; di < dupMatches.length; di++) {
                var dupItem = dupMatches[di];
                var resDup = range.search(dupItem.fullMatch, { matchCase: false });
                resDup.load('items');
                await context.sync();

                for (var k = 0; k < resDup.items.length; k++) {
                    resDup.items[k].insertText(dupItem.keepText, Word.InsertLocation.replace);
                    cleanCount++;
                }
            }

            await context.sync();
            statusDiv.className = 'status-msg success';
            if (cleanCount > 0) {
                statusDiv.innerText = '✅ Đã thay thế ' + totalReplaced + ' cụm từ & loại bỏ ' + cleanCount + ' lặp thừa trong ngoặc!';
            } else {
                statusDiv.innerText = '✅ Đã thay thế ' + totalReplaced + ' cụm từ viết tắt!';
            }
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = 'Lỗi: ' + (err.message || err);
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

function escapeHtml(str) {
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(str || ''));
    return div.innerHTML;
}

function escapeAttr(str) {
    return (str || '').replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

// ============================================================
// GẮN HÀM LÊN WINDOW ĐỂ HTML ONCLICK GỌI ĐƯỢC
// ============================================================
window.switchTab = switchTab;
window.formatDates = formatDates;
window.addAbbrevRule = addAbbrevRule;
window.removeAbbrevRule = removeAbbrevRule;
window.editAbbrevRule = editAbbrevRule;
window.cancelEditAbbrevRule = cancelEditAbbrevRule;
window.saveEditAbbrevRule = saveEditAbbrevRule;
window.exportAbbrevRules = exportAbbrevRules;
window.importAbbrevRules = importAbbrevRules;
window.replaceAbbreviations = replaceAbbreviations;
window.checkSpelling = checkSpelling;
window.fixSingleSpell = fixSingleSpell;
window.fixAllSpelling = fixAllSpelling;
