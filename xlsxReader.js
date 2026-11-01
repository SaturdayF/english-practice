// Read the first worksheet of a two-column .xlsx file without uploading it.
const decoder = new TextDecoder('utf-8');
const MAX_XML_BYTES = 32 * 1024 * 1024;

function readUint16(view, offset) {
    return view.getUint16(offset, true);
}

function readUint32(view, offset) {
    return view.getUint32(offset, true);
}

function normalizePath(path) {
    const parts = [];
    for (const part of path.replaceAll('\\', '/').split('/')) {
        if (part === '..') parts.pop();
        else if (part && part !== '.') parts.push(part);
    }
    return parts.join('/');
}

function zipEntries(buffer) {
    const view = new DataView(buffer);
    const bytes = new Uint8Array(buffer);
    let end = -1;
    for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset--) {
        if (readUint32(view, offset) === 0x06054b50) {
            end = offset;
            break;
        }
    }
    if (end < 0) throw new Error('这不是有效的 .xlsx 文件。');

    const count = readUint16(view, end + 10);
    let offset = readUint32(view, end + 16);
    const entries = new Map();
    for (let index = 0; index < count; index++) {
        if (readUint32(view, offset) !== 0x02014b50) throw new Error('Excel 文件结构不完整。');
        const method = readUint16(view, offset + 10);
        const packedSize = readUint32(view, offset + 20);
        const size = readUint32(view, offset + 24);
        const nameLength = readUint16(view, offset + 28);
        const extraLength = readUint16(view, offset + 30);
        const commentLength = readUint16(view, offset + 32);
        const localOffset = readUint32(view, offset + 42);
        const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
        entries.set(normalizePath(name), { method, packedSize, size, localOffset });
        offset += 46 + nameLength + extraLength + commentLength;
    }
    return { view, bytes, entries };
}

async function readEntry(zip, path) {
    const entry = zip.entries.get(normalizePath(path));
    if (!entry) throw new Error(`Excel 文件缺少 ${path}。`);
    if (entry.size > MAX_XML_BYTES) throw new Error('Excel 工作表过大。');
    const offset = entry.localOffset;
    if (readUint32(zip.view, offset) !== 0x04034b50) throw new Error('Excel 文件结构不完整。');
    const start = offset + 30 + readUint16(zip.view, offset + 26) + readUint16(zip.view, offset + 28);
    const compressed = zip.bytes.subarray(start, start + entry.packedSize);
    if (entry.method === 0) return decoder.decode(compressed);
    if (entry.method !== 8 || typeof DecompressionStream === 'undefined') {
        throw new Error('当前浏览器无法解压这个 Excel 文件，请用新版 Chrome 或 Edge。');
    }
    const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    const data = await new Response(stream).arrayBuffer();
    if (data.byteLength > MAX_XML_BYTES) throw new Error('Excel 工作表过大。');
    return decoder.decode(data);
}

function parseXml(source) {
    const doc = new DOMParser().parseFromString(source, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Excel 文件中的 XML 无法读取。');
    return doc;
}

function elements(parent, name) {
    return [...parent.getElementsByTagNameNS('*', name)];
}

function columnIndex(reference) {
    const letters = reference.match(/^[A-Z]+/i)?.[0]?.toUpperCase() ?? 'A';
    return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

window.readXlsx = async function readXlsx(file) {
    const zip = zipEntries(await file.arrayBuffer());
    const workbook = parseXml(await readEntry(zip, 'xl/workbook.xml'));
    const sheet = elements(workbook, 'sheet')[0];
    if (!sheet) throw new Error('Excel 文件中没有工作表。');
    const relationId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const relations = parseXml(await readEntry(zip, 'xl/_rels/workbook.xml.rels'));
    const relation = elements(relations, 'Relationship').find(item => item.getAttribute('Id') === relationId);
    if (!relation) throw new Error('找不到第一个工作表。');
    const target = relation.getAttribute('Target') ?? '';
    if (relation.getAttribute('TargetMode') === 'External') throw new Error('不支持外部工作表。');
    const sheetPath = normalizePath(target.startsWith('/') ? target : `xl/${target}`);

    let strings = [];
    if (zip.entries.has('xl/sharedStrings.xml')) {
        const shared = parseXml(await readEntry(zip, 'xl/sharedStrings.xml'));
        strings = elements(shared, 'si').map(item => elements(item, 't').map(text => text.textContent ?? '').join(''));
    }

    const worksheet = parseXml(await readEntry(zip, sheetPath));
    const rows = [];
    for (const row of elements(worksheet, 'row')) {
        const values = ['', ''];
        for (const cell of elements(row, 'c')) {
            const index = columnIndex(cell.getAttribute('r') ?? 'A');
            if (index > 1) continue;
            const type = cell.getAttribute('t');
            const value = elements(cell, 'v')[0]?.textContent ?? '';
            values[index] = type === 's'
                ? strings[Number(value)] ?? ''
                : type === 'inlineStr'
                    ? elements(cell, 't').map(text => text.textContent ?? '').join('')
                    : value;
        }
        if (values.some(value => value.trim())) rows.push(values);
    }
    return rows;
};
