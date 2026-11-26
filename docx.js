// 生成 Word 文档（.docx）。不依赖任何库：手写 OOXML，打进一个 stored ZIP。
const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const FONT = '微软雅黑';
const INK = '1C1C1E';
const MUTED = '7A7A7A';
const GREEN = '1F7A52';
const OT = '9A5B00';
const PAY = '2F6F8F';
const LINE = 'C9D3CD';
const HEAD_FILL = 'EAF2EC';
const DETAIL_FILL = 'F6F8F6';
const BALANCE_FILL = 'EAF2EC';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const encoder = new TextEncoder();

function escapeXml(text) {
  return String(text)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function run(text, options = {}) {
  const { bold = false, size = 18, color = INK } = options;
  return (
    '<w:r><w:rPr>' +
    `<w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}"/>` +
    (bold ? '<w:b/>' : '') +
    (color ? `<w:color w:val="${color}"/>` : '') +
    `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` +
    '</w:rPr>' +
    `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`
  );
}

function paragraph(content, options = {}) {
  const {
    align = 'left',
    before = 0,
    after = 0,
    line = 240,
    keepNext = false,
    shade,
    indent,
  } = options;
  const runs = Array.isArray(content) ? content.join('') : content;
  return (
    '<w:p><w:pPr>' +
    (keepNext ? '<w:keepNext/>' : '') +
    (shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` : '') +
    `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>` +
    (indent ? `<w:ind w:left="${indent}" w:right="${indent}"/>` : '') +
    `<w:jc w:val="${align}"/></w:pPr>` +
    runs +
    '</w:p>'
  );
}

function cell(content, width, options = {}) {
  const { fill, span } = options;
  const paragraphs = Array.isArray(content) ? content.join('') : content;
  return (
    '<w:tc><w:tcPr>' +
    `<w:tcW w:w="${width}" w:type="dxa"/>` +
    (span ? `<w:gridSpan w:val="${span}"/>` : '') +
    (fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : '') +
    '<w:vAlign w:val="center"/>' +
    '</w:tcPr>' +
    paragraphs +
    '</w:tc>'
  );
}

function row(cells, options = {}) {
  const { header = false } = options;
  return (
    '<w:tr><w:trPr>' +
    '<w:cantSplit/>' +
    (header ? '<w:tblHeader/>' : '') +
    '</w:trPr>' +
    cells.join('') +
    '</w:tr>'
  );
}

function table(widths, rows) {
  const total = widths.reduce((sum, width) => sum + width, 0);
  const borders =
    `<w:top w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>` +
    `<w:left w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>` +
    `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>`;
  return (
    '<w:tbl><w:tblPr>' +
    `<w:tblW w:w="${total}" w:type="dxa"/>` +
    '<w:jc w:val="center"/>' +
    `<w:tblBorders>${borders}</w:tblBorders>` +
    '<w:tblLayout w:type="fixed"/>' +
    '<w:tblCellMar>' +
    '<w:top w:w="70" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>' +
    '<w:bottom w:w="70" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>' +
    '</w:tblCellMar>' +
    '</w:tblPr>' +
    `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>` +
    rows.join('') +
    '</w:tbl>'
  );
}

const TABLE_WIDTH = 9600;

/// 一天的明细表：日期 / 状态 / 工日 / 加班 / 工资，加班时间和备注折成第二行。
function rowsTable(rows, fields = {}) {
  const show = {
    days: fields.days !== false,
    overtime: fields.overtime !== false,
    wage: fields.wage !== false,
    note: fields.note !== false,
  };
  const all = [
    { key: 'date', label: '日期', width: 2500 },
    { key: 'status', label: '状态', width: 1800 },
    { key: 'days', label: '工日', width: 1600 },
    { key: 'overtime', label: '加班', width: 1800 },
    { key: 'wage', label: '工资', width: 1900 },
  ];
  const cols = all.filter((col) => col.key === 'date' || col.key === 'status' || show[col.key]);
  const totalBase = cols.reduce((sum, col) => sum + col.width, 0);
  const widths = cols.map((col) => Math.round((col.width / totalBase) * TABLE_WIDTH));
  const span = cols.length;

  const head = row(
    cols.map((col, index) =>
      cell(
        paragraph(run(col.label, { bold: true, size: 18, color: '1F3A2D' }), { align: 'center' }),
        widths[index],
        { fill: HEAD_FILL }
      )
    ),
    { header: true }
  );

  const body = [];
  for (const item of rows) {
    const isPayment = item.note === '发工资' || item.note === '生活费';
    const noteColor = isPayment ? PAY : INK;

    const cells = [];
    cols.forEach((col, index) => {
      const width = widths[index];
      if (col.key === 'date') {
        cells.push(cell(paragraph(run(item.date, { size: 18 }), { align: 'left' }), width));
      } else if (col.key === 'status') {
        cells.push(
          cell(
            paragraph(
              run(item.status, {
                size: 18,
                bold: item.status === '上班',
                color: item.status === '上班' ? GREEN : item.status === '离场' ? OT : MUTED,
              }),
              { align: 'center' }
            ),
            width
          )
        );
      } else if (col.key === 'days') {
        cells.push(cell(paragraph(run(item.days, { size: 18, bold: true }), { align: 'center' }), width));
      } else if (col.key === 'overtime') {
        cells.push(
          cell(
            paragraph(
              item.hours
                ? run(`${item.hours} 小时`, { size: 18, bold: true, color: OT })
                : run('—', { size: 18, color: 'A8A8A8' }),
              { align: 'center' }
            ),
            width
          )
        );
      } else if (col.key === 'wage') {
        cells.push(cell(paragraph(run(item.wage, { size: 18, bold: true }), { align: 'right' }), width));
      }
    });
    body.push(row(cells));

    const detail = [];
    if (item.hours && show.overtime) {
      detail.push(run('加班时间 ', { size: 16, color: MUTED }));
      detail.push(run(item.overtime, { size: 16, color: INK }));
    }
    if (item.note !== '无' && show.note) {
      if (detail.length) detail.push(run('　·　', { size: 16, color: 'B8B8B8' }));
      // 收款那行只写「生活费 ¥X」，不加「备注」这类多余的词。
      if (!isPayment) detail.push(run('备注 ', { size: 16, color: MUTED }));
      detail.push(run(item.note, { size: 16, color: noteColor }));
      if (item.paid) {
        detail.push(run(`　${item.paid}`, { size: 16, color: PAY, bold: true }));
      }
    }

    if (detail.length) {
      body.push(
        row([
          cell(paragraph(detail, { align: 'left' }), TABLE_WIDTH, {
            span,
            fill: DETAIL_FILL,
          }),
        ])
      );
    }
  }

  const tableNote = row([
    cell(
      paragraph(
        run(
          fields.wage !== false || fields.paid !== false
            ? '工资 = 工日 × 日薪 + 加班小时 × 加班时薪　·　工资余额 = 项目累计工资 − 项目累计支账　·　支账即生活费。'
            : '工日按 0.5 / 1 天记，加班按小时记，只统计工时。',
          { size: 15, color: MUTED }
        ),
        { align: 'left' }
      ),
      TABLE_WIDTH,
      { span, fill: 'FBFCFB' }
    ),
  ]);

  return table(widths, [head, ...body, tableNote]);
}

function divider() {
  return (
    '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="31785A"/></w:pBdr>' +
    '<w:spacing w:before="260" w:after="100"/></w:pPr></w:p>'
  );
}

function pageBreak() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

/// 每个项目的「总结」：本期工日/加班/出勤、工资与支账、项目累计与余额。
function summaryXml(section, periodWord, fields = {}) {
  const month = section.month;
  const totals = section.totals;
  const show = {
    days: fields.days !== false,
    overtime: fields.overtime !== false,
    wage: fields.wage !== false,
    paid: fields.paid !== false,
  };
  const headParts = [];
  if (show.days) headParts.push(run(`${periodWord} ${month.days} 个工日`, { size: 19, bold: true }));
  const tail = [];
  if (show.overtime) tail.push(`加班 ${month.hours} 小时`);
  tail.push(`上班 ${month.worked} 天 · 休息 ${month.rested} 天`);
  if (month.leaves > 0) tail.push(`离场 ${month.leaves} 天`);
  headParts.push(run(`${show.days ? ' · ' : ''}${tail.join(' · ')}`, { size: 19, color: INK }));

  const moneyParts = [];
  if (show.wage) {
    moneyParts.push(run(`${periodWord}工资 `, { size: 20, color: MUTED }));
    moneyParts.push(run(month.wage, { size: 20, bold: true }));
  }
  if (show.paid) {
    moneyParts.push(run(`${moneyParts.length ? '　　' : ''}${periodWord}支账 `, { size: 20, color: MUTED }));
    moneyParts.push(run(month.paid, { size: 20, bold: true }));
  }

  const totalParts = [];
  if (show.wage) {
    totalParts.push(run('项目累计工资 ', { size: 20, color: MUTED }));
    totalParts.push(run(totals.wage, { size: 20, bold: true }));
  }
  if (show.paid) {
    totalParts.push(run(`${totalParts.length ? '　　' : ''}累计支账 `, { size: 20, color: MUTED }));
    totalParts.push(run(totals.paid, { size: 20, bold: true }));
  }

  return (
    paragraph(run('总结', { bold: true, size: 27, color: '1F3A2D' }), {
      after: 90,
      keepNext: true,
    }) +
    paragraph(headParts, { after: 70 }) +
    (moneyParts.length
      ? paragraph(moneyParts, { after: 40 })
      : '') +
    (totalParts.length
      ? paragraph(totalParts, { after: 110 })
      : '') +
    (fields.balance !== false
      ? paragraph(
          [
            run(`${totals.balanceLabel} `, { size: 24, bold: true, color: '1F3A2D' }),
            run(totals.balance, {
              size: 30,
              bold: true,
              color: totals.negative ? OT : GREEN,
            }),
          ],
          { shade: BALANCE_FILL, indent: 120, before: 60, after: 140 }
        )
      : '')
  );
}

/// 多个项目时每节开头的项目名和这段时间。
function sectionHead(section) {
  return (
    paragraph(run(section.project, { bold: true, size: 30, color: '1F3A2D' }), {
      after: 40,
      keepNext: true,
    }) +
    paragraph(run(`${section.periodLabel} 记工明细`, { size: 20, color: MUTED }), {
      after: 160,
      keepNext: true,
    })
  );
}

/// 多项目导出最后的合计：这段时间的和、项目累计的，都列出来。
function grandXml(grand, count, periodWord, fields = {}) {
  const show = {
    wage: fields.wage !== false,
    paid: fields.paid !== false,
  };
  const periodParts = [];
  if (show.wage) {
    periodParts.push(run(`${periodWord}工资 `, { size: 20, color: MUTED }));
    periodParts.push(run(grand.periodWage, { size: 20, bold: true }));
  }
  if (show.paid) {
    periodParts.push(run(`${periodParts.length ? '　　' : ''}${periodWord}支账 `, { size: 20, color: MUTED }));
    periodParts.push(run(grand.periodPaid, { size: 20, bold: true }));
  }
  const totalParts = [];
  if (show.wage) {
    totalParts.push(run('项目累计工资 ', { size: 20, color: MUTED }));
    totalParts.push(run(grand.wage, { size: 20, bold: true }));
  }
  if (show.paid) {
    totalParts.push(run(`${totalParts.length ? '　　' : ''}累计支账 `, { size: 20, color: MUTED }));
    totalParts.push(run(grand.paid, { size: 20, bold: true }));
  }

  return (
    divider() +
    paragraph(run(`合计（${count} 个项目）`, { bold: true, size: 26, color: '1F3A2D' }), {
      after: 100,
      keepNext: true,
    }) +
    (periodParts.length ? paragraph(periodParts, { after: 60 }) : '') +
    (totalParts.length ? paragraph(totalParts, { after: 90 }) : '') +
    (fields.balance !== false
      ? paragraph(
          [
            run(`${grand.balanceLabel} `, { size: 24, bold: true, color: '1F3A2D' }),
            run(grand.balance, { size: 30, bold: true, color: grand.negative ? OT : GREEN }),
          ],
          { shade: BALANCE_FILL, indent: 120, before: 60, after: 140 }
        )
      : '')
  );
}

function documentXml(data) {
  const sections = data.sections || [];
  const multi = sections.length > 1;
  const periodWord = data.periodWord || '本月';

  const head =
    paragraph(run(data.title || '记工表', { bold: true, size: 36, color: '1F3A2D' }), {
      align: 'center',
      after: 80,
    }) +
    paragraph(run(data.subtitle || '', { size: 22, color: MUTED }), {
      align: 'center',
      after: 40,
    }) +
    paragraph(run(`导出时间：${data.generatedAt}`, { size: 17, color: '9A9A9A' }), {
      align: 'center',
      after: 200,
    });

  const parts = [];
  sections.forEach((section, index) => {
    if (index > 0) parts.push(pageBreak());
    if (multi) parts.push(sectionHead(section));
    parts.push(rowsTable(section.rows, data.fields));
    parts.push(divider());
    parts.push(summaryXml(section, periodWord, data.fields));
  });
  if (multi && data.grand) parts.push(grandXml(data.grand, sections.length, periodWord, data.fields));

  const footer = paragraph(run('由「记工」导出', { size: 15, color: '9A9A9A' }), {
    align: 'right',
    before: 260,
  });

  return (
    XML_HEAD +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' +
    head +
    parts.join('') +
    footer +
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>' +
    '</w:sectPr>' +
    '</w:body></w:document>'
  );
}

function stylesXml() {
  return (
    XML_HEAD +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `<w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}"/>` +
    '<w:sz w:val="21"/><w:szCs w:val="21"/>' +
    '<w:lang w:val="en-US" w:eastAsia="zh-CN"/>' +
    '</w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '</w:styles>'
  );
}

function contentTypesXml() {
  return (
    XML_HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'
  );
}

function rootRelsXml() {
  return (
    XML_HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>'
  );
}

function documentRelsXml() {
  return (
    XML_HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>'
  );
}

function coreXml(data) {
  return (
    XML_HEAD +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(data.title || '记工表')}</dc:title>` +
    '<dc:creator>记工</dc:creator><cp:lastModifiedBy>记工</cp:lastModifiedBy>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${data.iso}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${data.iso}</dcterms:modified>` +
    '</cp:coreProperties>'
  );
}

function appXml() {
  return (
    XML_HEAD +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>记工</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
    '</Properties>'
  );
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const u16 = (value) => [value & 0xff, (value >>> 8) & 0xff];
const u32 = (value) => [
  value & 0xff,
  (value >>> 8) & 0xff,
  (value >>> 16) & 0xff,
  (value >>> 24) & 0xff,
];

function zip(files) {
  const chunks = [];
  const directory = [];
  let offset = 0;

  for (const file of files) {
    const name = encoder.encode(file.name);
    const bytes = encoder.encode(file.content);
    const crc = crc32(bytes);
    const local = [
      ...u32(0x04034b50),
      ...u16(20),
      ...u16(0x0800),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(crc),
      ...u32(bytes.length),
      ...u32(bytes.length),
      ...u16(name.length),
      ...u16(0),
    ];
    chunks.push(new Uint8Array(local), name, bytes);
    directory.push({ name, crc, size: bytes.length, offset });
    offset += local.length + name.length + bytes.length;
  }

  const directoryStart = offset;
  for (const item of directory) {
    const header = [
      ...u32(0x02014b50),
      ...u16(20),
      ...u16(20),
      ...u16(0x0800),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(item.crc),
      ...u32(item.size),
      ...u32(item.size),
      ...u16(item.name.length),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(item.offset),
    ];
    chunks.push(new Uint8Array(header), item.name);
    offset += header.length + item.name.length;
  }

  const directorySize = offset - directoryStart;
  chunks.push(
    new Uint8Array([
      ...u32(0x06054b50),
      ...u16(0),
      ...u16(0),
      ...u16(directory.length),
      ...u16(directory.length),
      ...u32(directorySize),
      ...u32(directoryStart),
      ...u16(0),
    ])
  );

  return new Blob(chunks, { type: DOCX_MIME });
}

export function buildWorkDocx(data) {
  return zip([
    { name: '[Content_Types].xml', content: contentTypesXml() },
    { name: '_rels/.rels', content: rootRelsXml() },
    { name: 'docProps/core.xml', content: coreXml(data) },
    { name: 'docProps/app.xml', content: appXml() },
    { name: 'word/document.xml', content: documentXml(data) },
    { name: 'word/_rels/document.xml.rels', content: documentRelsXml() },
    { name: 'word/styles.xml', content: stylesXml() },
  ]);
}

export function safeFileName(text) {
  return String(text || '')
    .replace(/[\\/:*?"<>|]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export async function saveFile(blob, filename, mime) {
  if (typeof File === 'function' && navigator.canShare) {
    const file = new File([blob], filename, { type: mime || blob.type });
    try {
      if (navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: filename });
        return 'shared';
      }
    } catch (error) {
      if (error && error.name === 'AbortError') return 'cancelled';
    }
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.rel = 'noopener';
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 8000);
  return 'downloaded';
}

export async function saveDocx(blob, filename) {
  return saveFile(blob, filename, DOCX_MIME);
}
