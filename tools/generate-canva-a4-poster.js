const pptxgen = require("pptxgenjs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const out = path.join(root, "SpotFinder-Poster-A4-Canva-Editable.pptx");

const pptx = new pptxgen();
pptx.author = "Zahran Al Syafit";
pptx.subject = "Poster A4 SpotFinder";
pptx.title = "SpotFinder - Poster A4";
pptx.company = "SpotFinder";
pptx.lang = "id-ID";
pptx.defineLayout({ name: "A4_PORTAIT", width: 8.27, height: 11.69 });
pptx.layout = "A4_PORTAIT";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "id-ID",
};

const C = {
  bg: "071016",
  panel: "10222D",
  panel2: "0B1922",
  line: "335263",
  text: "EEF7F4",
  muted: "B8C8D0",
  cyan: "35D0FF",
  green: "77E08F",
  yellow: "F7C948",
  orange: "FFA640",
  red: "FF7A7A",
  blue: "1798C8",
  ink: "071016",
};

const img = {
  logo: path.join(root, "assets", "img", "favicon.png"),
  hero: path.join(root, "assets", "img", "hero.png"),
  screen1: path.join(root, "assets", "img", "screen-1.png"),
  screen2: path.join(root, "assets", "img", "screen-2.png"),
  screen3: path.join(root, "assets", "img", "screen-3.png"),
};

function text(slide, value, x, y, w, h, size, color = C.text, opts = {}) {
  slide.addText(value, {
    x,
    y,
    w,
    h,
    fontSize: size,
    color,
    margin: 0.02,
    breakLine: false,
    fit: "shrink",
    ...opts,
  });
}

function card(slide, x, y, w, h, title, body, accent = C.cyan, titleSize = 12, bodySize = 8.7) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.05,
    fill: { color: C.panel, transparency: 2 },
    line: { color: C.line, transparency: 5, width: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 0.055,
    h,
    fill: { color: accent },
    line: { color: accent },
  });
  text(slide, title, x + 0.16, y + 0.13, w - 0.28, 0.25, titleSize, C.text, { bold: true });
  text(slide, body, x + 0.16, y + 0.46, w - 0.28, h - 0.55, bodySize, C.muted);
}

function pill(slide, value, x, y, w, accent = C.green) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.3,
    rectRadius: 0.12,
    fill: { color: accent },
    line: { color: accent },
  });
  text(slide, value, x, y + 0.075, w, 0.12, 8.3, C.ink, { bold: true, align: "center" });
}

function step(slide, n, title, body, x, y, w, accent) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y: y + 0.05,
    w: 0.34,
    h: 0.34,
    fill: { color: accent },
    line: { color: accent },
  });
  text(slide, String(n), x, y + 0.145, 0.34, 0.08, 8.5, C.ink, { bold: true, align: "center" });
  text(slide, title, x + 0.45, y, w - 0.45, 0.18, 9.4, C.text, { bold: true });
  text(slide, body, x + 0.45, y + 0.23, w - 0.45, 0.33, 7.8, C.muted);
}

const slide = pptx.addSlide();
slide.background = { color: C.bg };
slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 8.27, h: 11.69, fill: { color: C.bg }, line: { color: C.bg } });
slide.addShape(pptx.ShapeType.arc, { x: -1.35, y: -1.2, w: 3.9, h: 3.9, fill: { color: C.cyan, transparency: 74 }, line: { color: C.cyan, transparency: 100 } });
slide.addShape(pptx.ShapeType.arc, { x: 6.4, y: 8.6, w: 3.7, h: 3.7, fill: { color: C.green, transparency: 82 }, line: { color: C.green, transparency: 100 } });

// Header
slide.addImage({ path: img.logo, x: 0.45, y: 0.38, w: 0.48, h: 0.48 });
text(slide, "SPOTFINDER", 1.02, 0.47, 2.4, 0.18, 14, C.text, { bold: true, charSpace: 1.1 });
pill(slide, "PROTOTYPE SMART PARKING", 5.35, 0.43, 2.25, C.green);

text(slide, "Sistem Parkir Cerdas Berbasis IoT, Web, dan AI", 0.45, 1.08, 7.3, 0.78, 27, C.text, {
  bold: true,
  fontFace: "Aptos Display",
});
text(slide, "Membantu pengguna menemukan slot parkir lebih cepat dan membantu pengelola memantau parkir secara real-time.", 0.48, 1.9, 6.85, 0.35, 10.6, "D8E9ED");

// Hero visual
slide.addShape(pptx.ShapeType.roundRect, {
  x: 0.48,
  y: 2.35,
  w: 7.3,
  h: 2.05,
  rectRadius: 0.07,
  fill: { color: C.panel2 },
  line: { color: C.line },
});
slide.addImage({ path: img.hero, x: 0.62, y: 2.48, w: 2.15, h: 1.75 });
text(slide, "Apa itu SpotFinder?", 3.0, 2.57, 4.3, 0.27, 15, C.cyan, { bold: true });
text(slide, "SpotFinder adalah miniatur sistem parkir modern. Kendaraan masuk memakai RFID, slot dibaca sensor ultrasonik, palang bergerak otomatis, dan status parkir tampil di dashboard web.", 3.0, 2.95, 4.35, 0.82, 10.1, C.muted);
pill(slide, "RFID", 3.0, 3.92, 0.85, C.cyan);
pill(slide, "ESP32", 4.0, 3.92, 0.95, C.green);
pill(slide, "Dashboard", 5.1, 3.92, 1.25, C.yellow);
pill(slide, "AI Voice", 6.5, 3.92, 1.0, C.orange);

// Problem/Solution
card(slide, 0.48, 4.72, 2.25, 1.42, "Masalah", "Parkir manual membuat pengguna lama mencari slot, data kendaraan kurang rapi, dan pengelola sulit memantau kapasitas.", C.red, 12, 8.1);
card(slide, 3.02, 4.72, 2.25, 1.42, "Solusi", "SpotFinder menampilkan slot real-time, mengatur akses RFID, mencatat transaksi, dan memberi kontrol admin.", C.green, 12, 8.1);
card(slide, 5.56, 4.72, 2.22, 1.42, "Status Project", "Prototype miniatur untuk membuktikan konsep. Sistem dapat dikembangkan ke implementasi nyata.", C.cyan, 12, 8.1);

// How it works
text(slide, "Cara Kerja Singkat", 0.5, 6.45, 3.2, 0.24, 14.5, C.text, { bold: true });
slide.addShape(pptx.ShapeType.roundRect, { x: 0.48, y: 6.83, w: 7.3, h: 1.58, rectRadius: 0.05, fill: { color: C.panel2 }, line: { color: C.line } });
step(slide, 1, "Login / Daftar", "Pengguna masuk ke portal client.", 0.75, 7.08, 1.95, C.cyan);
step(slide, 2, "Scan RFID", "Server mengecek akses kendaraan.", 2.75, 7.08, 1.9, C.green);
step(slide, 3, "Sensor Slot", "Ultrasonik membaca slot terisi/kosong.", 4.68, 7.08, 2.2, C.yellow);
step(slide, 4, "Bayar & Keluar", "Biaya dihitung dan riwayat tersimpan.", 0.75, 7.76, 2.25, C.orange);
step(slide, 5, "Dashboard", "Admin melihat data real-time.", 3.35, 7.76, 2.1, C.blue);

// Features
text(slide, "Fitur Utama", 0.5, 8.7, 2.3, 0.22, 14.5, C.text, { bold: true });
card(slide, 0.48, 9.05, 1.72, 0.92, "Client Portal", "Saldo, booking, riwayat, dan chat.", C.cyan, 10, 7.2);
card(slide, 2.32, 9.05, 1.72, 0.92, "Admin Panel", "User, top-up, gate, analytics.", C.green, 10, 7.2);
card(slide, 4.16, 9.05, 1.72, 0.92, "AI Voice", "Notifikasi suara sistem.", C.yellow, 10, 7.2);
card(slide, 6.0, 9.05, 1.78, 0.92, "Report", "Struk dan riwayat transaksi.", C.orange, 10, 7.2);

// Note and future development
slide.addShape(pptx.ShapeType.roundRect, { x: 0.48, y: 10.22, w: 7.3, h: 0.72, rectRadius: 0.05, fill: { color: "0A202B" }, line: { color: C.line } });
text(slide, "Pengembangan Selanjutnya:", 0.68, 10.35, 1.85, 0.12, 8.8, C.green, { bold: true });
text(slide, "QRIS asli, kamera plat nomor, aplikasi mobile, notifikasi WhatsApp, deployment cloud, dan AI chatbot yang lebih pintar.", 2.45, 10.34, 5.1, 0.2, 8.1, C.muted);

// Footer
slide.addShape(pptx.ShapeType.rect, { x: 0, y: 11.17, w: 8.27, h: 0.52, fill: { color: "050B0F" }, line: { color: "050B0F" } });
text(slide, "Dibuat oleh: Zahran Al Syafit", 0.48, 11.36, 2.6, 0.12, 8.8, C.text, { bold: true });
text(slide, "Estimasi project: ± 6 bulan | Modal prototype: ± Rp 500.000", 3.0, 11.36, 4.8, 0.12, 8.1, C.muted, { align: "right" });

pptx.writeFile({ fileName: out }).then(() => console.log(out));
