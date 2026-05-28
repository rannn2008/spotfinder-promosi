const pptxgen = require("pptxgenjs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const out = path.join(root, "SpotFinder-Presentasi-Canva-Editable-v2.pptx");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Zahran Al Syafit";
pptx.company = "SpotFinder";
pptx.subject = "Presentasi SpotFinder 10 Menit";
pptx.title = "SpotFinder - Sistem Parkir Cerdas";
pptx.lang = "id-ID";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "id-ID",
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";
pptx.margin = 0;

const C = {
  bg: "071016",
  panel: "13242E",
  panel2: "0E1B24",
  line: "34515E",
  text: "EEF7F4",
  muted: "A8B9C2",
  cyan: "35D0FF",
  green: "77E08F",
  yellow: "F7C948",
  red: "FF7A7A",
  ink: "071016",
};

const img = {
  hero: path.join(root, "assets", "img", "hero.png"),
  logo: path.join(root, "assets", "img", "favicon.png"),
  s1: path.join(root, "assets", "img", "screen-1.png"),
  s2: path.join(root, "assets", "img", "screen-2.png"),
  s3: path.join(root, "assets", "img", "screen-3.png"),
};

function addBg(slide) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: C.bg },
    line: { color: C.bg },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: -1.15,
    y: -0.9,
    w: 4.2,
    h: 4.2,
    adjustPoint: 0.25,
    fill: { color: C.cyan, transparency: 76 },
    line: { color: C.cyan, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.arc, {
    x: 10.3,
    y: 4.6,
    w: 4.2,
    h: 4.2,
    fill: { color: C.green, transparency: 82 },
    line: { color: C.green, transparency: 100 },
  });
}

function title(slide, text, y = 0.48) {
  slide.addText(text, {
    x: 0.62,
    y,
    w: 11.8,
    h: 0.55,
    fontFace: "Aptos Display",
    fontSize: 31,
    bold: true,
    color: C.text,
    margin: 0,
    breakLine: false,
    fit: "shrink",
  });
}

function kicker(slide, text, x = 0.62, y = 0.22, w = 5) {
  slide.addText(text.toUpperCase(), {
    x,
    y,
    w,
    h: 0.24,
    fontSize: 10,
    bold: true,
    color: C.green,
    charSpace: 1.2,
    margin: 0,
    fit: "shrink",
  });
}

function paragraph(slide, text, x, y, w, h, size = 17, color = C.muted) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: size,
    color,
    breakLine: false,
    valign: "mid",
    margin: 0.02,
    fit: "shrink",
  });
}

function footer(slide, label, time) {
  slide.addText(label, {
    x: 0.62,
    y: 7.08,
    w: 5,
    h: 0.2,
    fontSize: 8.5,
    color: "B7C8D0",
    transparency: 20,
    margin: 0,
  });
  slide.addText(time, {
    x: 11.5,
    y: 7.08,
    w: 1.2,
    h: 0.2,
    fontSize: 8.5,
    color: "B7C8D0",
    transparency: 20,
    align: "right",
    margin: 0,
  });
}

function card(slide, x, y, w, h, head, body, accent = C.cyan) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: C.panel, transparency: 2 },
    line: { color: C.line, transparency: 10, width: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 0.06,
    h,
    fill: { color: accent },
    line: { color: accent },
  });
  slide.addText(head, {
    x: x + 0.18,
    y: y + 0.16,
    w: w - 0.34,
    h: 0.32,
    fontSize: 16,
    bold: true,
    color: C.text,
    margin: 0,
    fit: "shrink",
  });
  slide.addText(body, {
    x: x + 0.18,
    y: y + 0.58,
    w: w - 0.34,
    h: h - 0.72,
    fontSize: 11.5,
    color: C.muted,
    margin: 0,
    fit: "shrink",
    breakLine: false,
  });
}

function bullets(slide, items, x, y, w, h, size = 15) {
  slide.addText(
    items.map((t) => ({ text: t, options: { bullet: { type: "bullet" } } })),
    {
      x,
      y,
      w,
      h,
      fontSize: size,
      color: C.muted,
      breakLine: true,
      fit: "shrink",
      paraSpaceAfterPt: 8,
      margin: 0,
    }
  );
}

function addCover() {
  const s = pptx.addSlide();
  addBg(s);
  s.addImage({ path: img.logo, x: 0.66, y: 0.62, w: 0.45, h: 0.45 });
  s.addText("SpotFinder", { x: 1.2, y: 0.7, w: 2.6, h: 0.25, fontSize: 16, bold: true, color: C.text, margin: 0 });
  kicker(s, "Project Presentation", 0.66, 1.45, 4.5);
  s.addText("Sistem Parkir Cerdas Berbasis IoT dan Web", {
    x: 0.66,
    y: 1.78,
    w: 7.2,
    h: 1.75,
    fontFace: "Aptos Display",
    fontSize: 43,
    bold: true,
    color: C.text,
    margin: 0,
    fit: "shrink",
    breakLine: false,
  });
  paragraph(
    s,
    "Solusi untuk membantu pengguna mencari slot parkir, masuk dengan RFID, memantau status secara real-time, dan mengelola pembayaran secara digital.",
    0.7,
    3.72,
    6.7,
    0.85,
    17,
    "D6E6EA"
  );
  s.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 5.0, w: 2.9, h: 0.42, fill: { color: C.panel2 }, line: { color: C.line } });
  s.addText("By: Zahran Al Syafit", { x: 0.88, y: 5.12, w: 2.5, h: 0.16, fontSize: 11, bold: true, color: C.text, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 3.85, y: 5.0, w: 2.25, h: 0.42, fill: { color: C.panel2 }, line: { color: C.line } });
  s.addText("Durasi: 10 menit", { x: 4.02, y: 5.12, w: 1.9, h: 0.16, fontSize: 11, bold: true, color: C.text, margin: 0 });
  s.addImage({ path: img.hero, x: 8.05, y: 0.95, w: 4.15, h: 4.8 });
  footer(s, "01 - Cover", "0:45");
}

function addProblem() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Latar Belakang"); title(s, "Masalah parkir sering terlihat kecil, tetapi dampaknya besar.");
  paragraph(s, "Di area ramai, pengguna sering menghabiskan waktu untuk mencari slot kosong. Pengelola juga sulit mengetahui kapasitas, riwayat kendaraan, dan pemasukan secara cepat.", 0.7, 1.42, 5.6, 1.0, 16);
  bullets(s, ["Pengguna berputar mencari tempat kosong.", "Riwayat parkir dan transaksi tidak terpusat.", "Palang masuk masih banyak yang manual.", "Pengelola butuh rekap yang cepat dan rapi."], 0.72, 2.65, 5.3, 2.7);
  card(s, 6.8, 1.5, 2.0, 1.36, "Waktu", "Pencarian parkir lebih lama.", C.cyan);
  card(s, 9.1, 1.5, 2.0, 1.36, "Data", "Status dan riwayat belum rapi.", C.green);
  card(s, 6.8, 3.25, 2.0, 1.36, "Akses", "Kontrol masuk-keluar belum otomatis.", C.yellow);
  card(s, 9.1, 3.25, 2.0, 1.36, "Laporan", "Rekap operasional masih lambat.", C.red);
  footer(s, "02 - Masalah", "1:00");
}

function addGoals() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Tujuan Proyek"); title(s, "Membuat parkir lebih cepat, tertata, dan mudah dipantau.");
  paragraph(s, "SpotFinder dibuat sebagai prototype smart parking yang menyatukan deteksi slot, autentikasi pengguna, dashboard, wallet, booking, dan laporan.", 0.7, 1.35, 10.5, 0.65, 16);
  card(s, 0.7, 2.35, 2.75, 1.55, "Real-time Slot", "Status slot kosong, terisi, booking, atau pelanggaran terlihat langsung.", C.cyan);
  card(s, 3.85, 2.35, 2.75, 1.55, "RFID Access", "Kendaraan masuk melalui verifikasi kartu RFID dan data member.", C.green);
  card(s, 7.0, 2.35, 2.75, 1.55, "Cashless Wallet", "Saldo, top-up, biaya parkir, dan booking fee dikelola dari aplikasi.", C.yellow);
  card(s, 10.15, 2.35, 2.55, 1.55, "Dashboard Admin", "Admin melihat pengguna, transaksi, analytics, chat, dan gate control.", C.red);
  footer(s, "03 - Tujuan", "1:00");
}

function addFeatures() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Fitur Utama"); title(s, "Satu sistem untuk pengguna, admin, dan perangkat parkir.");
  const data = [
    ["Client Portal", "Login, registrasi, profil, avatar, saldo, riwayat, dan chat support.", C.cyan],
    ["Booking Slot", "Pengguna bisa memesan slot tersedia dengan biaya booking digital.", C.green],
    ["Admin Panel", "Kelola RFID, member, top-up, tarif, analytics, history, dan laporan.", C.yellow],
    ["IoT Detection", "ESP32 membaca sensor ultrasonik untuk mendeteksi kendaraan per slot.", C.cyan],
    ["Auto Gate", "Servo membuka dan menutup palang setelah akses disetujui.", C.green],
    ["AI Voice", "Asisten suara memberi notifikasi selamat datang, keluar, dan status sistem.", C.yellow],
  ];
  data.forEach((d, i) => card(s, 0.7 + (i % 3) * 3.95, 1.55 + Math.floor(i / 3) * 1.85, 3.45, 1.45, d[0], d[1], d[2]));
  footer(s, "04 - Fitur", "1:00");
}

function addFlow() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Cara Kerja"); title(s, "Alur sistem dari pengguna datang sampai keluar parkir.");
  const data = [
    ["1. Login", "Pengguna masuk ke portal client dan mendaftarkan kendaraan."],
    ["2. Cek Slot", "Dashboard menampilkan slot kosong dan rekomendasi slot terbaik."],
    ["3. Scan RFID", "ESP32 mengirim UID ke server untuk verifikasi akses."],
    ["4. Sensor", "Ultrasonik mendeteksi kendaraan dan memperbarui status slot."],
    ["5. Keluar", "Biaya dihitung otomatis, riwayat tersimpan, dan struk bisa dicetak."],
  ];
  data.forEach((d, i) => card(s, 0.48 + i * 2.48, 1.75, 2.18, 2.6, d[0], d[1], i % 2 ? C.green : C.cyan));
  footer(s, "05 - Alur", "1:00");
}

function addScreens() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Tampilan Aplikasi"); title(s, "Antarmuka dibuat agar status parkir mudah dibaca.");
  const shots = [
    [img.s1, "Smart Dashboard", "Slot, okupansi, aktivitas, statistik."],
    [img.s2, "Client Portal", "Saldo, booking, riwayat, profil."],
    [img.s3, "AI Support Chat", "Bantuan client dan admin."],
  ];
  shots.forEach((d, i) => {
    const x = 0.75 + i * 4.12;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.35, w: 3.05, h: 4.15, fill: { color: C.panel2 }, line: { color: C.line } });
    s.addImage({ path: d[0], x: x + 0.12, y: 1.48, w: 2.81, h: 3.58 });
    s.addText(d[1], { x, y: 5.72, w: 3.05, h: 0.25, fontSize: 15, bold: true, color: C.text, align: "center", margin: 0 });
    s.addText(d[2], { x, y: 6.05, w: 3.05, h: 0.34, fontSize: 10.5, color: C.muted, align: "center", margin: 0, fit: "shrink" });
  });
  footer(s, "06 - UI", "1:00");
}

function addTech() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Teknologi"); title(s, "Komponen web dan hardware saling terhubung melalui API.");
  paragraph(s, "Backend memakai PHP dan MySQL. Frontend memakai HTML, CSS, JavaScript, jQuery, dan Chart.js. Perangkat memakai ESP32, RFID RC522, sensor ultrasonik, OLED, LED, buzzer, dan servo palang.", 0.7, 1.35, 5.55, 1.28, 15);
  bullets(s, ["PHP + MySQL untuk backend dan database.", "ESP32 mengirim status sensor ke server.", "RFID RC522 untuk identifikasi member.", "Servo, LED, buzzer, dan OLED untuk prototype.", "ElevenLabs TTS untuk notifikasi suara."], 0.72, 3.0, 5.2, 2.8, 13.5);
  card(s, 6.75, 1.35, 4.75, 1.15, "API Integration", "Endpoint menghubungkan web app dengan ESP32.", C.cyan);
  card(s, 6.75, 2.85, 4.75, 1.15, "Database Tables", "users, sensor_status, parking_history, settings, topup_requests.", C.green);
  card(s, 6.75, 4.35, 4.75, 1.15, "Scalable Idea", "Bisa dikembangkan ke banyak slot, QRIS asli, dan plat recognition.", C.yellow);
  footer(s, "07 - Teknologi", "1:00");
}

function addDemo() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Cara Penggunaan Saat Demo"); title(s, "Urutan demo yang aman untuk presentasi 10 menit.");
  card(s, 0.72, 1.45, 5.45, 4.0, "Untuk Client", "1. Buka Client Portal.\n2. Login atau registrasi kendaraan.\n3. Lihat saldo dan slot real-time.\n4. Booking slot kosong.\n5. Gunakan chat support jika perlu.", C.cyan);
  card(s, 6.55, 1.45, 5.45, 4.0, "Untuk Admin", "1. Login admin.\n2. Lihat dashboard slot dan aktivitas.\n3. Approve top-up atau tambah RFID.\n4. Cek analytics dan history.\n5. Gunakan gate control untuk palang.", C.green);
  footer(s, "08 - Demo", "1:15");
}

function addBenefits() {
  const s = pptx.addSlide();
  addBg(s); kicker(s, "Kelebihan dan Manfaat"); title(s, "SpotFinder memberi nilai untuk pengguna dan pengelola.");
  card(s, 0.7, 1.55, 2.75, 1.55, "Efisien", "Mengurangi waktu mencari parkir karena status slot terlihat.", C.cyan);
  card(s, 3.85, 1.55, 2.75, 1.55, "Terdokumentasi", "Masuk, keluar, booking, top-up, dan biaya tercatat otomatis.", C.green);
  card(s, 7.0, 1.55, 2.75, 1.55, "Modern", "Mendukung wallet, chat, voice notification, dan analytics.", C.yellow);
  card(s, 10.15, 1.55, 2.55, 1.55, "Bisa Dikembangkan", "Bisa ditambah banyak slot, QRIS asli, dan kamera plat nomor.", C.red);
  paragraph(s, "Pengguna mendapat pengalaman parkir yang lebih cepat dan jelas. Admin mendapat data operasional, transaksi, dan kontrol perangkat dalam satu tempat.", 1.25, 4.35, 10.6, 0.7, 21, C.cyan);
  footer(s, "09 - Manfaat", "1:00");
}

function addClosing() {
  const s = pptx.addSlide();
  addBg(s);
  s.addImage({ path: img.logo, x: 6.12, y: 0.82, w: 0.68, h: 0.68 });
  kicker(s, "Penutup", 5.72, 1.78, 1.8);
  s.addText("Parkir lebih pintar, data lebih rapi, pengalaman lebih cepat.", {
    x: 1.15,
    y: 2.25,
    w: 11.0,
    h: 1.48,
    fontFace: "Aptos Display",
    fontSize: 40,
    bold: true,
    align: "center",
    color: C.text,
    margin: 0,
    fit: "shrink",
  });
  paragraph(s, "SpotFinder membuktikan bahwa masalah parkir bisa diselesaikan dengan kombinasi web app, database, IoT, dan desain penggunaan yang jelas.", 2.0, 4.1, 9.3, 0.75, 18, C.muted);
  s.addShape(pptx.ShapeType.roundRect, { x: 5.25, y: 5.35, w: 2.85, h: 0.55, fill: { color: C.green }, line: { color: C.green } });
  s.addText("Terima kasih", { x: 5.25, y: 5.51, w: 2.85, h: 0.15, fontSize: 16, bold: true, color: C.ink, align: "center", margin: 0 });
  footer(s, "10 - Closing", "1:00");
}

[
  addCover,
  addProblem,
  addGoals,
  addFeatures,
  addFlow,
  addScreens,
  addTech,
  addDemo,
  addBenefits,
  addClosing,
].forEach((fn) => fn());

pptx.writeFile({ fileName: out }).then(() => {
  console.log(out);
});
