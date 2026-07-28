// ============================================================
// Word Add-in: Tiện ích rút gọn văn bản
// Version: 1.2.0 - Multi-language support (VI/EN), renamed title
// ============================================================

// ============================================================
// INTERNATIONALIZATION (i18n)
// ============================================================
var translations = {
    vi: {
        app_title: "Tiện ích rút gọn văn bản",
        header_title: "📝 Tiện ích rút gọn văn bản",
        header_subtitle: "Tùy chọn xử lý & tối ưu hóa văn bản",
        tab_date: "📅 Ngày tháng",
        tab_abbrev: "🔤 Từ viết tắt",
        tab_spell: "🔍 Chính tả",
        date_title: "Rút ngắn định dạng ngày tháng",
        date_desc: "Tự động tìm và quy đổi ngày tháng dài sang định dạng viết tắt gọn gàng:",
        scope_selection: "Vùng đang chọn",
        scope_all: "Toàn bộ văn bản",
        date_btn: "Chuyển đổi ngay",
        date_processing: "⏳ Đang xử lý...",
        date_success: "✅ Đã chuyển đổi thành công {count} vị trí!",
        date_error: "❌ Lỗi: {error}",
        abbrev_title: "Thay thế từ khóa thành từ viết tắt",
        abbrev_desc: "Quản lý danh sách từ khóa và thay thế để rút ngắn văn bản.",
        btn_export: "📥 Xuất JSON",
        btn_import: "📤 Nhập JSON",
        ph_full: "Từ/Cụm từ đầy đủ (VD: Thành phố)",
        ph_short: "Từ viết tắt (VD: TP.)",
        btn_add_rule: "+ Thêm rule",
        th_full: "Cụm từ đầy đủ",
        th_short: "Từ viết tắt",
        th_action: "Thao tác",
        btn_edit_title: "Sửa",
        btn_del_title: "Xóa",
        btn_save_title: "Lưu",
        btn_cancel_title: "Hủy",
        abbrev_btn: "Thay thế viết tắt",
        abbrev_processing: "Đang thay thế từ viết tắt...",
        abbrev_success: "✅ Đã thay thế {replaced} cụm từ viết tắt!",
        abbrev_success_clean: "✅ Đã thay thế {replaced} cụm từ & loại bỏ {clean} lặp thừa trong ngoặc!",
        abbrev_empty_export: "⚠️ Danh sách từ điển đang trống!",
        abbrev_export_success: "✅ Đã xuất {count} quy tắc ra tệp tu_dien_viet_tat.json!",
        abbrev_import_success: "✅ Nhập thành công! (Thêm mới: {added}, Cập nhật: {updated})",
        abbrev_alert_empty: "Vui lòng nhập đầy đủ thông tin!",
        spell_title: "Rà soát chính tả Tiếng Việt",
        spell_desc: "Quét lỗi phụ âm đầu (n/l, ch/tr, x/s, c/k), dấu hỏi/ngã, từ ghép sai chính tả phổ biến.",
        spell_btn: "Rà soát chính tả",
        spell_checking: "Đang rà soát chính tả...",
        spell_no_errors: "🎉 Không phát hiện lỗi chính tả phổ biến nào!",
        spell_found_errors: "Phát hiện {count} nghi vấn chính tả:",
        spell_fix_all: "🔧 Sửa tất cả {count} lỗi",
        spell_found_word: "Từ phát hiện",
        spell_suggest_fix: "Gợi ý sửa",
        spell_fix_this: "Sửa lỗi này",
        spell_fix_all_success: "🎉 Đã tự động sửa thành công {count} lỗi chính tả!"
    },
    en: {
        app_title: "Text Shortener Utility",
        header_title: "📝 Text Shortener Utility",
        header_subtitle: "Text processing & optimization options",
        tab_date: "📅 Date Format",
        tab_abbrev: "🔤 Abbreviations",
        tab_spell: "🔍 Spell Check",
        date_title: "Shorten Date Format",
        date_desc: "Automatically convert long date formats to concise abbreviations:",
        scope_selection: "Selection",
        scope_all: "Entire document",
        date_btn: "Convert Now",
        date_processing: "⏳ Processing...",
        date_success: "✅ Successfully converted {count} occurrences!",
        date_error: "❌ Error: {error}",
        abbrev_title: "Replace Keywords with Abbreviations",
        abbrev_desc: "Manage keywords and replacements to shorten text.",
        btn_export: "📥 Export JSON",
        btn_import: "📤 Import JSON",
        ph_full: "Full phrase (e.g. City)",
        ph_short: "Abbreviation (e.g. City)",
        btn_add_rule: "+ Add Rule",
        th_full: "Full Phrase",
        th_short: "Abbreviation",
        th_action: "Actions",
        btn_edit_title: "Edit",
        btn_del_title: "Delete",
        btn_save_title: "Save",
        btn_cancel_title: "Cancel",
        abbrev_btn: "Replace Abbreviations",
        abbrev_processing: "Replacing abbreviations...",
        abbrev_success: "✅ Replaced {replaced} abbreviation phrases!",
        abbrev_success_clean: "✅ Replaced {replaced} phrases & removed {clean} redundant bracket duplicates!",
        abbrev_empty_export: "⚠️ Dictionary list is empty!",
        abbrev_export_success: "✅ Exported {count} rules to tu_dien_viet_tat.json!",
        abbrev_import_success: "✅ Import successful! (Added: {added}, Updated: {updated})",
        abbrev_alert_empty: "Please fill in all fields!",
        spell_title: "Vietnamese Spell Check",
        spell_desc: "Scan initial consonant errors (n/l, ch/tr, x/s, c/k), tone marks, and common typos.",
        spell_btn: "Check Spelling",
        spell_checking: "Checking spelling...",
        spell_no_errors: "🎉 No common spelling errors detected!",
        spell_found_errors: "Found {count} spelling issues:",
        spell_fix_all: "🔧 Fix all {count} errors",
        spell_found_word: "Detected word",
        spell_suggest_fix: "Suggestion",
        spell_fix_this: "Fix this issue",
        spell_fix_all_success: "🎉 Automatically fixed {count} spelling errors!"
    }
};

var currentLang = localStorage.getItem('app_lang') || 'vi';

function t(key, params) {
    var dict = translations[currentLang] || translations.vi;
    var str = dict[key] || translations.vi[key] || key;
    if (params) {
        Object.keys(params).forEach(function(k) {
            str = str.replace(new RegExp('\\{' + k + '\\}', 'g'), params[k]);
        });
    }
    return str;
}

function setLanguage(lang) {
    if (lang !== 'vi' && lang !== 'en') return;
    currentLang = lang;
    try {
        localStorage.setItem('app_lang', lang);
    } catch (e) {
        console.warn('Cannot save app_lang to localStorage:', e);
    }
    updateUILanguage();
}

function updateUILanguage() {
    var btnVi = document.getElementById('lang-vi');
    var btnEn = document.getElementById('lang-en');
    if (btnVi && btnEn) {
        if (currentLang === 'vi') {
            btnVi.classList.add('active');
            btnEn.classList.remove('active');
        } else {
            btnEn.classList.add('active');
            btnVi.classList.remove('active');
        }
    }

    document.querySelectorAll('[data-i18n]').forEach(function(el) {
        var key = el.getAttribute('data-i18n');
        if (key === 'app_title') {
            document.title = t(key);
        } else {
            el.innerText = t(key);
        }
    });

    document.querySelectorAll('[data-i18n-ph]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-ph');
        el.placeholder = t(key);
    });

    renderAbbrevTable();
}

// Đảm bảo Office.js khởi tạo hoàn tất
Office.onReady(function(info) {
    console.log("Office ready info:", info);
    loadAbbrevRules();
    updateUILanguage();
    console.log("Add-in initialized successfully.");
});

// If Office.onReady doesn't fire immediately in plain browser view
document.addEventListener('DOMContentLoaded', function() {
    updateUILanguage();
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
    document.querySelectorAll('.tab-btn').forEach(function(btn) {
        btn.classList.remove('active');
    });
    document.querySelectorAll('.tab-content').forEach(function(content) {
        content.classList.remove('active');
    });

    var activeTabObj = document.getElementById(tabId);
    if (activeTabObj) activeTabObj.classList.add('active');

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
    statusDiv.innerText = t('date_processing');

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'dateScope');
            range.load('text');
            await context.sync();

            var text = range.text || '';
            var replaceCount = 0;

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

                range.load('text');
                await context.sync();
                text = range.text || '';
            }

            await context.sync();
            statusDiv.className = 'status-msg success';
            statusDiv.innerText = t('date_success', { count: replaceCount });
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = t('date_error', { error: err.message || err });
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
                '<button class="btn-edit" style="background:var(--success-color);" title="' + t('btn_save_title') + '" onclick="saveEditAbbrevRule(' + idx + ')">💾</button>' +
                '<button class="btn-del" title="' + t('btn_cancel_title') + '" onclick="cancelEditAbbrevRule()">✕</button>' +
                '</td>';
        } else {
            tr.innerHTML =
                '<td>' + escapeHtml(rule.full) + '</td>' +
                '<td><strong>' + escapeHtml(rule.short) + '</strong></td>' +
                '<td style="text-align: center;">' +
                '<button class="btn-edit" title="' + t('btn_edit_title') + '" onclick="editAbbrevRule(' + idx + ')">✏️</button>' +
                '<button class="btn-del" title="' + t('btn_del_title') + '" onclick="removeAbbrevRule(' + idx + ')">✕</button>' +
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
        alert(t('abbrev_alert_empty'));
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
        alert(t('abbrev_alert_empty'));
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
        statusDiv.innerText = t('abbrev_empty_export');
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
        statusDiv.innerText = t('abbrev_export_success', { count: abbrevRules.length });
    } catch (err) {
        console.error('exportAbbrevRules error:', err);
        statusDiv.className = 'status-msg status-error';
        statusDiv.innerText = t('date_error', { error: err.message || err });
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
            statusDiv.innerText = t('abbrev_import_success', { added: addedCount, updated: updatedCount });
        } catch (err) {
            console.error('importAbbrevRules error:', err);
            statusDiv.className = 'status-msg status-error';
            statusDiv.innerText = t('date_error', { error: err.message || err });
        } finally {
            event.target.value = '';
        }
    };
    reader.readAsText(file);
}

async function replaceAbbreviations() {
    var statusDiv = document.getElementById('abbrevStatus');
    statusDiv.className = 'status-msg';
    statusDiv.innerText = t('abbrev_processing');

    try {
        await Word.run(async function(context) {
            var range = getRange(context, 'abbrevScope');
            var totalReplaced = 0;

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
                statusDiv.innerText = t('abbrev_success_clean', { replaced: totalReplaced, clean: cleanCount });
            } else {
                statusDiv.innerText = t('abbrev_success', { replaced: totalReplaced });
            }
        });
    } catch (err) {
        console.error(err);
        statusDiv.className = 'status-msg error';
        statusDiv.innerText = t('date_error', { error: err.message || err });
    }
}

// ============================================================
// 3. TÍNH NĂNG RÀ SOÁT CHÍNH TẢ TIẾNG VIỆT
// ============================================================
var commonVietnameseErrors = [
    { err: 'xản xuất', fix: 'sản xuất', desc: "Sai 'x' → 's'" },
    { err: 'sơ xuất', fix: 'sơ suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'đột suất', fix: 'đột xuất', desc: "Sai 'suất' → 'xuất'" },
    { err: 'xác xuất', fix: 'xác suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'năng xuất', fix: 'năng suất', desc: "Sai 'xuất' → 'suất'" },
    { err: 'suất sắc', fix: 'xuất sắc', desc: "Sai 'suất' → 'xuất'" },
    { err: 'xuất xắc', fix: 'xuất sắc', desc: "Sai 'xắc' → 'sắc'" },
    { err: 'suất xắc', fix: 'xuất sắc', desc: "Sai 'suất xắc' → 'xuất sắc'" },
    { err: 'chuẩn đoán', fix: 'chẩn đoán', desc: "Sai 'chuẩn' → 'chẩn'" },
    { err: 'thăm quan', fix: 'tham quan', desc: "Sai 'thăm' → 'tham'" },
    { err: 'chỉnh chu', fix: 'chỉn chu', desc: "Sai 'chỉnh' → 'chỉn'" },
    { err: 'trút kinh nghiệm', fix: 'rút kinh nghiệm', desc: "Sai 'trút' → 'rút'" },
    { err: 'suôn sẽ', fix: 'suôn sẻ', desc: "Sai dấu ngã → hỏi" },
    { err: 'cũng cố', fix: 'củng cố', desc: "Sai dấu ngã → hỏi" },
    { err: 'dể dàng', fix: 'dễ dàng', desc: "Sai dấu hỏi → ngã" },
    { err: 'bổ xung', fix: 'bổ sung', desc: "Sai 'xung' → 'sung'" },
    { err: 'vô hình chung', fix: 'vô hình trung', desc: "Sai 'chung' → 'trung'" },
    { err: 'tựu chung', fix: 'tựu trung', desc: "Sai 'chung' → 'trung'" },
    { err: 'chín mùi', fix: 'chín muồi', desc: "Sai 'mùi' → 'muồi'" },
    { err: 'khoảng khắc', fix: 'khoảnh khắc', desc: "Sai 'khoảng' → 'khoảnh'" }
];

async function checkSpelling() {
    var resultsDiv = document.getElementById('spellResults');
    resultsDiv.innerHTML = '<div class="status-msg">' + t('spell_checking') + '</div>';

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
        resultsDiv.innerHTML = '<div class="status-msg error">' + t('date_error', { error: err.message || err }) + '</div>';
    }
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function renderSpellFindings(findings) {
    var resultsDiv = document.getElementById('spellResults');
    if (findings.length === 0) {
        resultsDiv.innerHTML = '<div class="status-msg success">' + t('spell_no_errors') + '</div>';
        return;
    }

    var html = '<div class="status-msg error">' + t('spell_found_errors', { count: findings.length }) + '</div>';
    html += '<button class="btn btn-primary" style="margin-top:6px;margin-bottom:8px" onclick="fixAllSpelling()">' + t('spell_fix_all', { count: findings.length }) + '</button>';

    findings.forEach(function(item) {
        html += '<div class="spell-item">';
        html += '<div>' + t('spell_found_word') + ': <span class="original">' + item.original + '</span></div>';
        html += '<div>' + t('spell_suggest_fix') + ': <span class="suggestion">' + item.correct + '</span>';
        if (item.desc) html += ' (' + item.desc + ')';
        html += '</div>';
        html += '<div class="spell-actions">';
        html += "<button class=\"btn-sm\" onclick=\"fixSingleSpell('" + item.original.replace(/'/g, "\\'") + "', '" + item.correct.replace(/'/g, "\\'") + "')\">" + t('spell_fix_this') + "</button>";
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

            await checkSpelling();
        });
    } catch (err) {
        console.error(err);
        alert(t('date_error', { error: err.message || err }));
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
            resultsDiv.innerHTML = '<div class="status-msg success">' + t('spell_fix_all_success', { count: count }) + '</div>';
        });
    } catch (err) {
        console.error(err);
        alert(t('date_error', { error: err.message || err }));
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

// Gắn lên window
window.setLanguage = setLanguage;
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
