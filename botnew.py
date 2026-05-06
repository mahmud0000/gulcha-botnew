const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat
} = require('docx');
const fs = require('fs');

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
    columnWidths: [9000],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9000, type: WidthType.DXA },
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
    columnWidths: [9000],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: color },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: color },
          left: { style: BorderStyle.SINGLE, size: 8, color: color },
          right: { style: BorderStyle.SINGLE, size: 1, color: color },
        },
        width: { size: 9000, type: WidthType.DXA },
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
          width: { size: 800, type: WidthType.DXA },
          shading: { fill: "2E86C1", type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 100, right: 100 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: number, bold: true, size: 36, color: "FFFFFF", font: "Arial" })]
          })]
        }),
        new TableCell({
          borders,
          width: { size: 8200, type: WidthType.DXA },
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
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 32, bold: true, font: "Arial", color: "1A5276" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { size: 26, bold: true, font: "Arial", color: "2E86C1" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
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
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 800 },
        children: [new TextRun({ text: "Boshlang'ich foydalanuvchilar uchun to'liq qo'llanma", size: 24, font: "Arial", color: "999999", italics: true })]
      }),

      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [9000],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            width: { size: 9000, type: WidthType.DXA },
            shading: { fill: "EBF5FB", type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 300, right: 300 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Bu qo'llanmada nima bor?", bold: true, size: 26, font: "Arial", color: "1A5276" })] }),
              spacing(),
              bullet("Telegram bot uchun token olish"),
              bullet("GitHub da kod yuklash"),
              bullet("Render.com da bepul 24/7 server sozlash"),
              bullet("Botni sinab ko'rish"),
              bullet("Har kuni ishlatish ko'rsatmasi"),
            ]
          })]
        })]
      }),

      spacing(3),
      divider(),

      // QADAM 1
      h1("1-QADAM: Tayyorgarlik"),
      para("Boshlashdan avval quyidagi 3 ta narsani tayyor qiling:"),
      spacing(),

      stepTable("1", "Bot Token", "Telegram da @BotFather ni toping → /newbot yuboring → nom bering → token oling"),
      spacing(),
      stepTable("2", "Admin ID", "@userinfobot ga /start yuboring → u sizning ID raqamingizni beradi"),
      spacing(),
      stepTable("3", "Karta raqami", "Mijozlar karta orqali to'lovi uchun o'z karta raqamingiz"),

      spacing(2),
      noteBox("⚠️  Muhim: Agar oldin bot yaratgan bo'lsangiz, @BotFather → /revoke → yangi token oling. Eski token artiq ishlamaydi.", "F39C12", "FEF9E7"),

      divider(),

      // QADAM 2
      h1("2-QADAM: GitHub da fayllarni yangilash"),
      para("GitHub — kodingiz saqlanadigan joy. Render.com githubdan kodni o'qiydi."),
      spacing(),

      h2("bot.py faylini yangilash"),
      numbered("github.com ga kiring → ustaka akkauntiga login qiling"),
      numbered("ustaka/gulcha-bot repositoryni oching"),
      numbered("bot.py faylini bosing"),
      numbered("O'ng yuqoridagi '...' (uch nuqta) tugmasini bosing"),
      numbered("'Edit file' ni tanlang"),
      numbered("Ichidagi BARCHA matnni o'chiring"),
      para("   Telefonda: matnni bosib ushlab turing → 'Hammasini tanlash' → o'chiring", "777777"),
      numbered("Yangi bot.py kodi matnini to'liq ko'chiring va joylashtiring"),
      numbered("'Commit changes' tugmasini bosing → yana 'Commit changes'"),

      spacing(),
      noteBox("✅  bot.py muvaffaqiyatli yangilandi!", "27AE60", "EAFAF1"),

      spacing(),
      h2("requirements.txt faylini yangilash"),
      numbered("Xuddi shunday requirements.txt faylini ham oching va tahrirlang"),
      numbered("Ichini to'liq o'chirib, faqat quyidagini yozing:"),
      spacing(),
      codeBox(["python-telegram-bot==20.7"]),
      spacing(),
      numbered("'Commit changes' bosing"),
      spacing(),
      noteBox("✅  GitHub tayyor! Ikkala fayl ham yangilandi.", "27AE60", "EAFAF1"),

      divider(),

      // QADAM 3
      h1("3-QADAM: Render.com da sozlash"),
      para("Render — botingiz 24/7 to'xtovsiz ishlashi uchun bepul server."),
      spacing(),

      h2("Environment Variables (muhim sozlamalar)"),
      para("Bu yerda botga kerakli maxfiy ma'lumotlar kiritiladi:"),
      spacing(),

      numbered("dashboard.render.com ga kiring"),
      numbered("'My project' ni bosing → 'gulcha-bot' serviceni bosing"),
      numbered("Chap menyudan 'Environment' ni tanlang"),
      numbered("Quyidagi 3 ta variable ni kiriting:"),

      spacing(),
      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [2500, 6500],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders,
                width: { size: 2500, type: WidthType.DXA },
                shading: { fill: "1A5276", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "Key (Nom)", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders,
                width: { size: 6500, type: WidthType.DXA },
                shading: { fill: "1A5276", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "Value (Qiymat) — nima yozish kerak", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })]
              })
            ]
          }),
          ...[
TOKEN = os.getenv("8766748683:AAFJ6cKF56fTORAISb6uQzuojpHRTxRJVGI")
)"],
            ["ADMIN_ID", "5594795335
) — @userinfobot dan oling"],
            ["CARD_NUMBER", "Karta raqamingiz (masalan: 5614 6821 2328 4780)"],
          ].map(([key, val], i) => new TableRow({
            children: [
              new TableCell({
                borders,
                width: { size: 2500, type: WidthType.DXA },
                shading: { fill: i % 2 === 0 ? "F2F3F4" : "FDFEFE", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: key, size: 22, font: "Courier New", color: "922B21" })] })]
              }),
              new TableCell({
                borders,
                width: { size: 6500, type: WidthType.DXA },
                shading: { fill: i % 2 === 0 ? "F2F3F4" : "FDFEFE", type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: val, size: 22, font: "Arial" })] })]
              })
            ]
          }))
        ]
      }),

      spacing(),
      numbered("Har birini kiritib 'Save' bosing"),
      numbered("'Save Changes' bosing"),

      spacing(),
      h2("Start Command tekshirish"),
      numbered("Render da 'Settings' bo'limiga o'ting"),
      numbered("'Start Command' maydonida quyidagi bo'lishi kerak:"),
      spacing(),
      codeBox(["python bot.py"]),
      spacing(),
      noteBox("⚠️  Agar bo'sh bo'lsa yoki boshqacha bo'lsa — o'chiring va: python bot.py yozing", "F39C12", "FEF9E7"),

      spacing(),
      h2("Qayta deploy qilish"),
      numbered("Render da 'Manual Deploy' tugmasini bosing"),
      numbered("'Deploy latest commit' ni tanlang"),
      numbered("Loglarni kuting — quyidagi yozuv chiqsin:"),
      spacing(),
      codeBox(["✅ Gulcha Taom Bot ishga tushdi!"]),
      spacing(),
      noteBox("✅  Bot 24/7 ishlaydi! Render ni yopsangiz ham bot to'xtamaydi.", "27AE60", "EAFAF1"),

      divider(),

      // QADAM 4
      h1("4-QADAM: Botni sinab ko'rish"),

      h2("Admin sifatida tekshirish"),
      numbered("Telegram da @gulchataom_bot ni toping"),
      numbered("/start yuboring"),
      numbered("Quyidagi admin panel chiqishi kerak:"),
      spacing(),
      codeBox([
        "👋 Salom, Admin!",
        "",
        "  📋 Menyu kiritish     📦 Buyurtmalar",
        "  📢 Xabar yuborish     📊 Hisobot",
        "  👥 Mijozlar bazasi",
      ]),

      spacing(),
      h2("Mijoz sifatida sinash"),
      noteBox("Boshqa telefon yoki Telegram akkauntidan @gulchataom_bot ga kiring", "2E86C1", "EBF5FB"),
      spacing(),
      numbered("Botga /start yuboring"),
      numbered("Ism familiyangizni yozing"),
      numbered("Telefon raqam yuboring (tugma chiqadi)"),
      numbered("Lokatsiya yuboring (📍 Lokatsiyamni yuborish tugmasi chiqadi)"),
      numbered("'🍽 Buyurtma berish' bosing"),
      numbered("Taom tanlang → miqdor → Cola/Non/Salat taklifi → to'lov → tasdiqlang"),
      numbered("Sizning asosiy telefoningizda admin xabari kelishi kerak"),

      spacing(),
      h2("Buyurtmani boshqarish"),
      numbered("Admin xabarida tugmalar bo'ladi: 👨\u200D🍳 → 🚚 → ✅"),
      numbered("Har tugmani bosganingizda mijozga avtomatik xabar ketadi"),
      numbered("✅ Yetkazildi bosganda mijozga baholash tugmalari keladi"),

      divider(),

      // KUN TARTIBI
      h1("HAR KUNGI FOYDALANISH"),

      h2("Ertalab — menyu kiritish"),
      numbered("Botga /start yuboring (admin sifatida)"),
      numbered("'📋 Menyu kiritish' bosing"),
      numbered("Har bir taomni yuboring:"),
      bullet("Rasmsiz: Osh - 25000 (matn bilan)"),
      bullet("Rasmli: Rasmni yuboring, caption da: Osh - 25000"),
      numbered("/done yozing — menyu saqlandi va mijozlar buyurtma bera oladi"),

      spacing(),
      h2("Kun davomida"),
      bullet("Har yangi buyurtmada Telegram da xabar keladi"),
      bullet("Buyurtma holatini 👨\u200D🍳 Tayyorlanmoqda → 🚚 Yetkazilmoqda → ✅ Yetkazildi orqali o'zgartiring"),
      bullet("Mijozga har bosqichda avtomatik xabar ketadi"),

      spacing(),
      h2("Kech — hisobot"),
      numbered("'📊 Hisobot' bosing — kunlik tushum va buyurtmalar soni"),
      numbered("'👥 Mijozlar bazasi' — barcha mijozlar, GPS, VIP status"),

      spacing(),
      noteBox("⚠️  Menyu har kuni yarim tunda avtomatik tozalanadi. Har kuni ertalab yangi menyu kiritish kerak!", "F39C12", "FEF9E7"),

      divider(),

      // ADMIN BUYRUQLARI
      h1("ADMIN BUYRUQLARI JADVALI"),

      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [2500, 3500, 3000],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "1A5276", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "Buyruq / Tugma", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 3500, type: WidthType.DXA }, shading: { fill: "1A5276", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "Nima qiladi", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, shading: { fill: "1A5276", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: "Qachon", bold: true, size: 22, font: "Arial", color: "FFFFFF" })] })] }),
            ]
          }),
          ...[
            ["/start", "Admin panelni ochadi", "Har kuni birinchi kirganda"],
            ["📋 Menyu kiritish", "Bugungi menyuni kiritish", "Har kuni ertalab"],
            ["/done", "Menyu kiritishni yakunlash", "Barcha taomlar kiritilgach"],
            ["📦 Buyurtmalar", "Bugungi buyurtmalar ro'yxati", "Kun davomida"],
            ["📊 Hisobot", "Kunlik/haftalik statistika", "Kun oxirida"],
            ["👥 Mijozlar bazasi", "Barcha mijozlar + GPS", "Kerak bo'lganda"],
            ["📢 Xabar yuborish", "Barcha mijozlarga xabar", "Aksiya va yangiliklarда"],
            ["/cancel", "Istalgan amalni bekor qilish", "Noto'g'ri bosib qo'yganda"],
          ].map(([cmd, func, when], i) => new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: i%2===0 ? "F2F3F4" : "FDFEFE", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: cmd, size: 20, font: "Courier New", color: "922B21" })] })] }),
              new TableCell({ borders, width: { size: 3500, type: WidthType.DXA }, shading: { fill: i%2===0 ? "F2F3F4" : "FDFEFE", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: func, size: 20, font: "Arial" })] })] }),
              new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, shading: { fill: i%2===0 ? "F2F3F4" : "FDFEFE", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 },
                children: [new Paragraph({ children: [new TextRun({ text: when, size: 20, font: "Arial", color: "666666" })] })] }),
            ]
          }))
        ]
      }),

      divider(),

      // MUAMMOLAR
      h1("KO'P UCHRAYDIGAN MUAMMOLAR"),

      h2("Bot javob bermayapti"),
      bullet("Render da bot 'Running' (yashil) holatidami? → dashboard.render.com da tekshiring"),
      bullet("BOT_TOKEN to'g'ri kiritilganmi? → Environment bo'limida ko'ring"),
      bullet("Bepul Render da 15 daqiqa faoliyat bo'lmasa bot uxlaydi → xabar yuboring, 50 soniyada uyg'onadi"),

      spacing(),
      h2("Menyu kiritishda xato"),
      bullet("Format to'g'ri bo'lsin: Taom nomi - narx  (tire atrofida bo'sh joy bo'lsin)"),
      bullet("Narxda faqat raqam: 25000  (vergul, so'm yoki boshqa belgi bo'lmasin)"),
      bullet("Har yangi menyu uchun avval '📋 Menyu kiritish' bosing"),

      spacing(),
      h2("Admin xabarlari kelmayapti"),
      bullet("ADMIN_ID to'g'ri kiritilganmi? @userinfobot dan tekshiring"),
      bullet("Botga /start yuborgan bo'lishingiz kerak"),

      spacing(),
      h2("Lokatsiya ishlamayapti"),
      bullet("Telegram da Location ruxsati berilganmi? → Telegram Sozlamalar → Privacy → Location"),
      bullet("Lokatsiya yubora olmasangiz — matn bilan ham manzil kiritish mumkin"),

      divider(),

      // XULOSA
      h1("MUHIM ESLATMALAR"),

      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [9000],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            width: { size: 9000, type: WidthType.DXA },
            shading: { fill: "FEF9E7", type: ShadingType.CLEAR },
            margins: { top: 180, bottom: 180, left: 280, right: 280 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Xavfsizlik", bold: true, size: 24, font: "Arial", color: "784212" })] }),
              spacing(),
              bullet("BOT_TOKEN ni hech kim bilan ulashmang"),
              bullet("Token oshkor bo'lsa — @BotFather → /revoke → yangi token oling"),
              bullet("ADMIN_ID faqat sizniki bo'lsin"),
            ]
          })]
        })]
      }),

      spacing(),

      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [9000],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            width: { size: 9000, type: WidthType.DXA },
            shading: { fill: "EAFAF1", type: ShadingType.CLEAR },
            margins: { top: 180, bottom: 180, left: 280, right: 280 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Biznesingizni o'stirish uchun maslahat", bold: true, size: 24, font: "Arial", color: "1E8449" })] }),
              spacing(),
              bullet("Bot havolasini ofislarga yuboring: t.me/gulchataom_bot"),
              bullet("Har kuni menyu rasmini Telegram kanalga ham joylashtiring"),
              bullet("VIP mijozlar (5+ buyurtma) uchun maxsus chegirma e'lon qiling"),
              bullet("Haftalik hisobotni ko'rib, eng ko'p sotilgan taomlarni ko'paytiring"),
            ]
          })]
        })]
      }),

      spacing(3),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Omad! Gulcha Taom biznesingiz rivojlansin! 🚀", bold: true, size: 28, font: "Arial", color: "1A5276" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100 },
        children: [new TextRun({ text: "t.me/gulchataom_bot", size: 24, font: "Arial", color: "2E86C1", italics: true })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Gulcha_Taom_Deploy_Qollanma.docx", buffer);
  console.log("Tayyor!");
}).catch(e => { console.error(e); process.exit(1); });
