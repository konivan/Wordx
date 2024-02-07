const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  WidthType,
  HeightRule,
  BorderStyle,
  HeadingLevel,
  TextRun,
} = require("docx");

const createBtn = document.getElementById("createBtn");
const createRowBtn = document.getElementById("createRowBtn");
const wordsCount = document.getElementById("wordsCount");
createBtn.addEventListener("click", () => createTable());
createRowBtn.addEventListener("click", () => createNewRow());

const borderTemplate = {
  left: {
    style: BorderStyle.SINGLE,
    size: 9,
  },
  right: {
    style: BorderStyle.SINGLE,
    size: 9,
  },
  top: {
    style: BorderStyle.SINGLE,
    size: 9,
  },
  bottom: {
    style: BorderStyle.SINGLE,
    size: 9,
  },
};

const rows = [
  new TableRow({
    height: {
      value: 900,
      rule: HeightRule.EXACT,
    },
    children: [
      "Word/Phrase",
      "Transcription",
      "Definition(s)",
      "Translation",
      "Examples",
    ].map(
      (item) =>
        new TableCell({
          borders: borderTemplate,
          width: {
            size: 1900,
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_2,
              outlineLevel: 5,
              children: [
                new TextRun({
                  text: item,
                  bold: true,
                  color: "000000",
                }),
              ],
            }),
          ],
        })
    ),
  }),
];

function createTable() {
  const table = new Table({
    columnWidths: [1900, 1900, 1900, 1900, 1900],
    rows: rows,
  });

  const doc = new Document({
    sections: [
      {
        children: [table],
      },
    ],
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
  });
}

function createNewRow() {
  let rowCount = Number(wordsCount.innerText);
  const newWord = [];
  document
    .querySelectorAll("#word")
    .forEach((word) => newWord.push(word.value));

  const template = new TableRow({
    height: {
      value: 2300,
      rule: HeightRule.EXACT,
    },
    children: newWord.map(
      (item) =>
        new TableCell({
          borders: borderTemplate,
          width: {
            size: 1900,
            type: WidthType.DXA,
          },
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_3,
              children: [
                new TextRun({
                  text: `${item}`,
                  color: "000000",
                }),
              ],
            }),
          ],
        })
    ),
  });

  document.querySelectorAll("#word").forEach((word) => (word.value = ""));
  rowCount++;
  wordsCount.innerText = rowCount;
  return rows.push(template);
}
