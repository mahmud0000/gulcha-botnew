const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat
} = require('docx');
const fs = require('fs');

// Yordamchi funksiyalar va stillar
const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };

function spacing(n = 1) {
  return new Paragraph({ spacing: { before: n * 80, after: n * 80 }, children: [new TextRun("")] });
}

function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "AAAAAA" } },
    spacing: { before: 200, after: 200 },
    children: [new TextRun("")]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, color: "1A5276", font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, color: "2E86C1", font: "Arial" })]
  });
}

function para(text, color) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: color || "333333" })]
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function numbered(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    numbering: { reference: "numbers", level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function codeBox(lines) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        shading: { fill: "2C3E50", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: lines.map(line => new Paragraph({
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text: line, size: 20, font: "Courier New", color: "A9DFBF" })]
        }))
      })]
    })]
  });
}

function noteBox(text, color, bg) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: color },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: color },
          left: { style: BorderStyle.SINGLE, size: 8, color: color },
          right: { style: BorderStyle.SINGLE, size: 1, color: color },
        },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [new Paragraph({
          children: [new TextRun({ text, size: 22, font: "Arial", color: "333333" })]
        })]
      })]
    })]
  });
}

function stepTable(number, title, desc) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [800, 8200],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders,
          shading: { fill: "2E86C1", type: ShadingType.CLEAR },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: number, bold: true, size: 36, color: "FFFFFF", font: "Arial" })]
          })]
        }),
        new TableCell({
          borders,
          shading: { fill: "EBF5FB", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 200, right: 150 },
          children: [
            new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 24, font: "Arial", color: "1A5276" })] }),
            new Paragraph({ children: [new TextRun({ text: desc, size: 22, font: "Arial", color: "555555" })] })
          ]
        })
      ]
    })]
  });
}

// ASOSIY HUJJAT TUZILMASI
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // MUQOVA
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 1440, after: 200 },
        children: [new TextRun({ text: "GULCHA TAOM BOT", bold: true, size: 52, font: "Arial", color: "1A5276" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "GitHub va Render orqali deploy qilish", size: 28, font: "Arial", color: "555555" })]
      }),
      spacing(2),

      new Table({
        width: { size: 9000, type: WidthType.DXA },
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            shading: { fill: "EBF5FB", type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 300, right: 300 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Bu qo'llanmada nima bor?", bold: true, size: 26, font: "Arial", color: "1A5276" })] }),
              spacing(),
              bullet("Telegram bot uchun token olish"),
              bullet("GitHub da kod yuklash"),
              bullet("Render.com da bepul 24/7 server sozlash"),
            ]
          })]
        })]
      }),

      divider(),

      h1("1-QADAM: Tayyorgarlik"),
      stepTable("1", "Bot Token", "Telegram da @BotFather ni toping → /newbot"),
      spacing(),
      stepTable("2", "Admin ID", "@userinfobot ga /start yuboring"),

      divider(),

      h1("2-QADAM: Render.com muhit o'zgaruvchilari"),
      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [3000, 6000],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders,
                shading: { fill: "1A5276", type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [new TextRun({ text: "Key (Nom)", bold: true, color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders,
                shading: { fill: "1A5276", type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [new TextRun({ text: "Value (Qiymat)", bold: true, color: "FFFFFF" })] })]
              })
            ]
          }),
          ...[
            ["BOT_TOKEN", "8619295805:AAHAR0TuU-wzDmgZS2kqDlnaHUr1AYxfNoQ (BotFather dan)"],
            ["ADMIN_ID", "5594795335 (Sizning ID)"],
            ["CARD_NUMBER", "5614 6821 2328 4780 (Karta raqamingiz)"]
          ].map(([key, val]) => new TableRow({
            children: [
              new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: key, font: "Courier New" })] })] }),
              new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: val })] })] })
            ]
          }))
        ]
      }),

      spacing(2),
      h1("HAR KUNGI FOYDALANISH"),
      h2("Ertalab — menyu kiritish"),
      numbered("Botga /start yuboring"),
      numbered("'📋 Menyu kiritish' bosing"),
      bullet("Format: Taom nomi - narx"),
      numbered("/done yozing"),

      spacing(3),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Omad! Gulcha Taom rivojlansin! 🚀", bold: true, size: 28, color: "1A5276" })]
      })
    ]
  }]
});

// Faylni saqlash
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Gulcha_Taom_Deploy_Qollanma.docx", buffer);
  console.log("Hujjat muvaffaqiyatli yaratildi!");
}).catch(e => {
  console.error("Xatolik yuz berdi:", e);
});
