// Manage the local vocabulary list and the two directions of practice.
const $ = selector => document.querySelector(selector);
const state = { rows: [], cards: [], index: 0, correct: 0, attempted: 0, revealed: false, chineseColumn: 1 };
const defaultRows = window.DEFAULT_ROWS;
const usesDesktopKeyboard = () => window.matchMedia('(hover: hover) and (pointer: fine)').matches;

function normalizeAnswer(value) {
    return value.normalize('NFKC').toLocaleLowerCase('en').replace(/[‘’]/g, "'")
        .replace(/[“”]/g, '"').replace(/\s+/g, ' ').trim().replace(/[.!?。！？]+$/, '').trim();
}

function updateStatus(message, error = false) {
    $('#fileStatus').textContent = message;
    $('#fileStatus').classList.toggle('error', error);
}

function detectChineseColumn(rows) {
    const sample = rows.slice(0, 20);
    const count = column => sample.reduce((total, row) => total + (/[\u3400-\u9fff]/.test(row[column]) ? 1 : 0), 0);
    return count(0) > count(1) ? 0 : 1;
}

function isHeader(row) {
    const labels = row.map(value => value.trim().toLowerCase());
    return labels.some(value => /^(english|英文|英语|单词|句子)$/.test(value))
        && labels.some(value => /^(chinese|中文|汉语|释义|翻译)$/.test(value));
}

function shuffle(items) {
    const result = [...items];
    for (let i = result.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [result[i], result[j]] = [result[j], result[i]];
    }
    return result;
}

function loadRows(rows, status, isDefault = false) {
    state.rows = rows.map(row => [...row]);
    state.chineseColumn = detectChineseColumn(state.rows);
    $('#columnSelect').value = String(state.chineseColumn);
    $('#options').hidden = false;
    $('#clearButton').hidden = false;
    $('#restoreButton').hidden = isDefault;
    updateStatus(status);
    startPractice();
}

function startPractice() {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const cards = state.rows.flatMap(row => {
        const chinese = row[state.chineseColumn].trim();
        const english = row[1 - state.chineseColumn].trim();
        if (!chinese || !english) return [];
        if (mode === 'write') return [{ chinese, english, direction: 'write' }];
        if (mode === 'reveal') return [{ chinese, english, direction: 'reveal' }];
        return [{ chinese, english, direction: 'write' }, { chinese, english, direction: 'reveal' }];
    });
    state.cards = shuffle(cards);
    state.index = 0;
    state.correct = 0;
    state.attempted = 0;
    render();
}

function render() {
    const hasCards = state.cards.length > 0;
    $('#emptyState').hidden = hasCards;
    $('#practice').hidden = !hasCards;
    if (!hasCards) return;

    const card = state.cards[state.index];
    state.revealed = false;
    $('#progressText').textContent = `${state.index + 1} / ${state.cards.length}`;
    $('#progressBar').style.width = `${(state.index + 1) / state.cards.length * 100}%`;
    $('#scoreText').textContent = `拼写正确 ${state.correct} / ${state.attempted}`;
    $('#direction').textContent = card.direction === 'write' ? '看中文，写英文' : '看英文，想中文';
    $('#prompt').textContent = card.direction === 'write' ? card.chinese : card.english;
    $('#prompt').classList.toggle('englishPrompt', card.direction === 'reveal');
    $('#writeForm').hidden = card.direction !== 'write';
    $('#revealButton').hidden = card.direction !== 'reveal';
    $('#answerInput').value = card.result?.input ?? card.draft ?? '';
    $('#answerInput').disabled = false;
    $('#submitButton').disabled = false;
    $('#feedback').hidden = true;
    $('#nextButton').hidden = true;
    $('#previousButton').disabled = state.index === 0;
    if (card.result) {
        showAnswer(card.result.correct);
        return;
    }
    if (card.direction === 'write' && usesDesktopKeyboard()) {
        $('#answerInput').focus();
    }
}

function showAnswer(isCorrect = null) {
    if (state.revealed) return;
    state.revealed = true;
    const card = state.cards[state.index];
    card.result = { correct: isCorrect, input: $('#answerInput').value };
    const feedback = $('#feedback');
    feedback.hidden = false;
    feedback.className = isCorrect === true ? 'feedback correct' : isCorrect === false ? 'feedback incorrect' : 'feedback';
    $('#resultLabel').textContent = isCorrect === true ? '答对了' : isCorrect === false ? '再记一次' : '参考答案';
    $('#answerText').textContent = card.direction === 'write' ? card.english : card.chinese;
    $('#answerInput').disabled = true;
    $('#submitButton').disabled = true;
    $('#revealButton').hidden = true;
    $('#nextButton').hidden = false;
    $('#nextButton').textContent = state.index === state.cards.length - 1 ? '再练一轮' : '下一题';
    $('#scoreText').textContent = `拼写正确 ${state.correct} / ${state.attempted}`;
    $('#nextButton').focus();
}

$('#fileInput').addEventListener('change', async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    updateStatus('正在读取…');
    try {
        if (!/\.xlsx$/i.test(file.name)) throw new Error('请选择 .xlsx 格式的 Excel 文件。旧版 .xls 请先另存为 .xlsx。');
        if (file.size > 20 * 1024 * 1024) throw new Error('文件超过 20 MB，请拆分后再导入。');
        let rows = await window.readXlsx(file);
        if (rows.length && isHeader(rows[0])) rows = rows.slice(1);
        rows = rows.filter(row => row[0].trim() && row[1].trim());
        if (!rows.length) throw new Error('没有找到同时包含英文和中文的两列内容。');
        loadRows(rows, `已导入 ${rows.length} 条 · ${file.name}`);
    } catch (error) {
        updateStatus(`导入失败：${error instanceof Error ? error.message : '文件读取失败。'} 当前内容未改变。`, true);
    }
    $('#fileInput').value = '';
});

$('#clearButton').addEventListener('click', () => {
    state.rows = [];
    state.cards = [];
    state.correct = 0;
    state.attempted = 0;
    $('#options').hidden = true;
    $('#clearButton').hidden = true;
    $('#restoreButton').hidden = false;
    $('#fileInput').value = '';
    updateStatus('当前题库已清空。可导入文件或恢复默认内容。');
    render();
});
$('#restoreButton').addEventListener('click', () => {
    $('#fileInput').value = '';
    loadRows(defaultRows, `已加载默认题库 · ${defaultRows.length} 条`, true);
});

$('#columnSelect').addEventListener('change', event => {
    state.chineseColumn = Number(event.target.value);
    startPractice();
});
document.querySelectorAll('input[name="mode"]').forEach(input => input.addEventListener('change', startPractice));
$('#writeForm').addEventListener('submit', event => {
    event.preventDefault();
    if (state.revealed) return;
    if (!$('#answerInput').value.trim()) {
        showAnswer();
        return;
    }
    state.attempted++;
    const correct = normalizeAnswer($('#answerInput').value) === normalizeAnswer(state.cards[state.index].english);
    if (correct) state.correct++;
    showAnswer(correct);
});
$('#revealButton').addEventListener('click', () => showAnswer());
$('#previousButton').addEventListener('click', () => {
    if (state.index === 0) return;
    state.cards[state.index].draft = $('#answerInput').value;
    state.index--;
    render();
});
$('#nextButton').addEventListener('click', () => {
    if (state.index === state.cards.length - 1) startPractice();
    else {
        state.index++;
        render();
    }
});

document.addEventListener('keydown', event => {
    if (event.key !== 'Enter' || event.isComposing || event.keyCode === 229
        || event.ctrlKey || event.altKey || event.metaKey || event.shiftKey
        || !usesDesktopKeyboard() || !state.cards.length) return;
    const control = event.target.closest('input, select, textarea, button, a, [contenteditable]');
    if (control?.id === 'previousButton') return;
    if (control && !$('#practice').contains(control)) return;
    event.preventDefault();
    if (event.repeat) return;
    if (state.revealed) $('#nextButton').click();
    else if (state.cards[state.index].direction === 'write') $('#writeForm').requestSubmit();
    else showAnswer();
});

loadRows(defaultRows, `已加载默认题库 · ${defaultRows.length} 条`, true);
