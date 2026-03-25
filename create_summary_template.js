// Скрипт для создания шаблона template_summary.docx
// ГИБРИДНЫЙ ПОДХОД: шаблон содержит оформление + тег {TABLE}
// Запуск: node create_summary_template.js

const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");

// ── XML PARTS ──
const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const wordRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

function p(text, opts = {}) {
  const bold = opts.bold ? "<w:b/>" : "";
  const italic = opts.italic ? "<w:i/>" : "";
  const sz = opts.sz || 22;
  const jc = opts.jc || "center";
  const spacing = opts.spacingAfter != null ? `<w:spacing w:after="${opts.spacingAfter}"/>` : "";
  return `<w:p><w:pPr><w:jc w:val="${jc}"/>${spacing}</w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${bold}${italic}</w:rPr>
    <w:t xml:space="preserve">${text}</w:t></w:r>
  </w:p>`;
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${p("СВОДНАЯ ВЕДОМОСТЬ", { bold: true, sz: 28, spacingAfter: 0 })}
    ${p("успеваемости учащихся группы «{group}»", { bold: true, sz: 20, spacingAfter: 0 })}
    ${p("Специальность: «{profRu}»", { bold: true, italic: true, sz: 20, spacingAfter: 0 })}
    ${p("ПЛ № 3 за {monthLabel} {studyYear} гг.", { bold: true, sz: 20, spacingAfter: 120 })}
    ${p("{TABLE}", { sz: 2, spacingAfter: 0 })}
    ${p("", { sz: 10, spacingAfter: 200 })}
    ${p("Зам. директора по УПМР ___________________                                              Куратор группы {curatorFullName}", { sz: 20, jc: "left" })}
    <w:sectPr>
      <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
      <w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="0" w:footer="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

// ── CREATE DOCX ──
const zip = new PizZip();
zip.file("[Content_Types].xml", contentTypesXml);
zip.file("_rels/.rels", relsXml);
zip.file("word/_rels/document.xml.rels", wordRelsXml);
zip.file("word/document.xml", documentXml);
zip.file("word/settings.xml", settingsXml);

const buf = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
const outPath = path.join(__dirname, "templates", "template_summary.docx");
fs.writeFileSync(outPath, buf);
console.log(`✅ Шаблон создан: ${outPath}`);
console.log(`   Размер: ${buf.length} байт`);
console.log(`\n📋 ТЕГИ В ШАБЛОНЕ (можно менять в Word):`);
console.log(`   {group}          — Название группы`);
console.log(`   {profRu}         — Специальность`);
console.log(`   {monthLabel}     — Месяц`);
console.log(`   {studyYear}      — Учебный год`);
console.log(`   {curatorFullName}— ФИО куратора`);
console.log(`   {TABLE}          — АВТОМАТИЧЕСКАЯ ТАБЛИЦА (предметы + оценки)`);
console.log(`\n✨ В шаблоне можно менять:`);
console.log(`   - Шрифты, размеры, отступы, ориентацию`);
console.log(`   - Текст заголовков, подписей`);
console.log(`   - Порядок элементов`);
console.log(`   НО НЕ УДАЛЯТЬ теги в фигурных скобках!`);
