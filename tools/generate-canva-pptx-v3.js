const pptxgen = require("pptxgenjs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const out = path.join(root, "SpotFinder-Presentasi-Canva-10Menit-v3.pptx");

const pptx = new pptxgen();
pptx.author = "Zahran Al Syafit";
pptx.subject = "Presentasi Project SpotFinder";
pptx.title = "SpotFinder - Presentasi 10 Menit";
pptx.company = "SpotFinder";
pptx.lang = "id-ID";
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "id-ID",
};

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
  orange: "FFA640",
  blue: "1798C8",
  ink: "071016",
};

const img = {
  hero: path.join(root, "assets", "img", "hero.png"),
  logo: path.join(root, "assets", "img", "favicon.png"),
  s1: path.join(root, "assets", "img", "screen-1.png"),
  s2: path.join(root, "assets", "img", "screen-2.png"),
  s3: path.join(root, "assets", "img", "screen-3.png"),
};

function bg(slide) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg }, line: { color: C.bg } });
  slide.addShape(pptx.ShapeType.arc, { x: -1.25, y: -1.2, w: 4.4, h: 4.4, fill: { color: C.cyan, transparency: 78 }, line: { color: C.cyan, transparency: 100 } });
  slide.addShape(pptx.ShapeType.arc, { x: 10.7, y: 4.5, w: 4.3, h: 4.3, fill: { color: C.green, transparency: 84 }, line: { color: C.green, transparency: 100 } });
}

function kicker(slide, txt, x = 0.68, y = 0.25, w = 5.5) {
  slide.addText(txt.toUpperCase(), { x, y, w, h: 0.22, fontSize: 10, bold: true, color: C.green, charSpace: 1.2, margin: 0 });
}

function title(slide, txt, y = 0.58, size = 32) {
  slide.addText(txt, { x: 0.68, y, w: 11.8, h: 0.75, fontFace: "Aptos Display", fontSize: size, bold: true, color: C.text, margin: 0, fit: "shrink" });
}

function text(slide, txt, x, y, w, h, size = 15, color = C.muted, opts = {}) {
  slide.addText(txt, { x, y, w, h, fontSize: size, color, margin: 0.02, fit: "shrink", breakLine: false, ...opts });
}

function footer(slide, label, time) {
  slide.addText(label, { x: 0.68, y: 7.05, w: 5, h: 0.22, fontSize: 8.5, color: "B7C8D0", transparency: 18, margin: 0 });
  slide.addText(time, { x: 11.55, y: 7.05, w: 1.1, h: 0.22, fontSize: 8.5, color: "B7C8D0", transparency: 18, align: "right", margin: 0 });
}

function card(slide, x, y, w, h, head, body, accent = C.cyan, size = 11.5) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.line, transparency: 8, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.065, h, fill: { color: accent }, line: { color: accent } });
  slide.addText(head, { x: x + 0.18, y: y + 0.16, w: w - 0.34, h: 0.32, fontSize: 16, bold: true, color: C.text, margin: 0, fit: "shrink" });
  slide.addText(body, { x: x + 0.18, y: y + 0.58, w: w - 0.34, h: h - 0.72, fontSize: size, color: C.muted, margin: 0, fit: "shrink", breakLine: false });
}

function bullets(slide, items, x, y, w, h, size = 13.8) {
  slide.addText(items.map((item) => ({ text: item, options: { bullet: { type: "bullet" } } })), {
    x, y, w, h, fontSize: size, color: C.muted, breakLine: true, fit: "shrink", paraSpaceAfterPt: 7, margin: 0,
  });
}

function notes(slide, speak, animate) {
  slide.addNotes(`NARASI: ${speak}\n\nANIMASI CANVA: ${animate}`);
}

function slide1() {
  const s = pptx.addSlide(); bg(s);
  s.addImage({ path: img.logo, x: 0.72, y: 0.62, w: 0.45, h: 0.45 });
  text(s, "SpotFinder", 1.25, 0.72, 2.4, 0.24, 16, C.text, { bold: true });
  kicker(s, "Judul dan Tema Project", 0.72, 1.42);
  text(s, "SpotFinder", 0.72, 1.8, 6.6, 0.7, 45, C.text, { bold: true, fontFace: "Aptos Display" });
  text(s, "Sistem Parkir Cerdas Berbasis IoT, Web, Database, dan AI Voice Assistant", 0.74, 2.55, 6.75, 0.85, 24, C.cyan, { bold: true });
  text(s, "Project ini dibuat untuk membantu pengguna menemukan slot parkir lebih cepat dan membantu pengelola memantau data parkir secara real-time.", 0.78, 3.75, 6.2, 0.8, 15.5, "D6E6EA");
  card(s, 0.78, 5.0, 2.8, 0.64, "By", "Zahran Al Syafit", C.green, 13);
  card(s, 3.85, 5.0, 2.4, 0.64, "Durasi", "10 Menit", C.yellow, 13);
  s.addImage({ path: img.hero, x: 8.0, y: 0.88, w: 4.15, h: 4.9 });
  notes(s, "Assalamualaikum, perkenalkan saya Zahran Al Syafit. Saya akan mempresentasikan SpotFinder, sistem parkir cerdas berbasis IoT dan web.", "Logo: Breathe. Judul: Rise. Gambar hero: Pan/Photo Zoom. Byline: Fade.");
  footer(s, "01 - Judul dan Tema", "0:45");
}

function slide2() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Latar Belakang / Masalah"); title(s, "Masalah parkir manual masih sering terjadi di area ramai.");
  text(s, "Pada sistem parkir biasa, pengguna sering tidak tahu slot mana yang kosong. Di sisi pengelola, data kendaraan, transaksi, dan kapasitas parkir juga tidak selalu tercatat secara cepat.", 0.72, 1.45, 6.0, 1.1, 15.5);
  bullets(s, ["Pengguna membuang waktu mencari tempat kosong.", "Status slot tidak terlihat secara real-time.", "Proses masuk dan keluar masih banyak yang manual.", "Riwayat parkir dan transaksi belum tersusun rapi.", "Pengelola sulit melihat laporan operasional dengan cepat."], 0.75, 3.0, 5.6, 2.5);
  card(s, 7.05, 1.55, 2.0, 1.35, "Waktu", "Pencarian parkir menjadi lebih lama.", C.cyan);
  card(s, 9.45, 1.55, 2.0, 1.35, "Data", "Riwayat dan transaksi tidak terpusat.", C.green);
  card(s, 7.05, 3.35, 2.0, 1.35, "Akses", "Palang masuk-keluar belum otomatis.", C.yellow);
  card(s, 9.45, 3.35, 2.0, 1.35, "Laporan", "Rekap admin membutuhkan waktu.", C.red);
  notes(s, "Latar belakang project ini adalah masalah parkir manual: sulit mencari slot kosong, data tidak real-time, dan pengelolaan masih kurang efisien.", "Judul: Rise. Bullet masuk satu per satu. Cards: Pop/Stomp berurutan.");
  footer(s, "02 - Latar Belakang", "1:00");
}

function slide3() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Tujuan / Solusi"); title(s, "SpotFinder hadir sebagai solusi parkir yang lebih cepat dan terdata.");
  text(s, "Tujuan utama project ini adalah membuat prototype smart parking yang menggabungkan aplikasi web, database, perangkat IoT, dan notifikasi AI agar proses parkir lebih mudah dipantau.", 0.72, 1.42, 11.0, 0.7, 15.5);
  card(s, 0.8, 2.35, 2.75, 1.55, "Real-time Slot", "Status kosong, terisi, booking, dan pelanggaran terlihat langsung.", C.cyan);
  card(s, 3.85, 2.35, 2.75, 1.55, "RFID Access", "Kartu RFID dipakai untuk validasi kendaraan dan member.", C.green);
  card(s, 6.9, 2.35, 2.75, 1.55, "Digital Wallet", "Saldo, top-up, booking, dan biaya parkir dapat tercatat.", C.yellow);
  card(s, 9.95, 2.35, 2.55, 1.55, "Admin Control", "Admin memantau transaksi, pengguna, gate, chat, dan analytics.", C.red);
  text(s, "Solusi utama: pengguna lebih cepat menemukan slot, sedangkan admin mendapatkan data parkir yang lebih rapi.", 1.25, 5.1, 10.8, 0.45, 18, C.cyan, { bold: true, align: "center" });
  notes(s, "Solusinya adalah SpotFinder: sistem yang memperlihatkan slot real-time, validasi RFID, pembayaran digital, dan dashboard admin.", "Cards: Fade from bottom satu per satu. Kalimat solusi: Typewriter atau Fade.");
  footer(s, "03 - Tujuan dan Solusi", "1:00");
}

function slide4() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Cara Kerja"); title(s, "Alur sistem dari kendaraan masuk sampai keluar parkir.");
  const steps = [
    ["1. Login", "Pengguna masuk ke portal client atau daftar kendaraan."],
    ["2. Cek Slot", "Dashboard menampilkan slot kosong dan status parkir."],
    ["3. Scan RFID", "UID kartu dikirim ESP32 ke server untuk dicek."],
    ["4. Sensor", "Ultrasonik membaca kendaraan dan update status slot."],
    ["5. Keluar", "Biaya dihitung, riwayat tersimpan, dan struk bisa dicetak."],
  ];
  steps.forEach((d, i) => card(s, 0.48 + i * 2.48, 1.75, 2.18, 2.6, d[0], d[1], i % 2 ? C.green : C.cyan));
  text(s, "Data mengalir dari hardware ke server, lalu ditampilkan di dashboard client dan admin.", 1.35, 5.05, 10.65, 0.45, 18, C.yellow, { bold: true, align: "center" });
  notes(s, "Cara kerjanya dimulai dari login, cek slot, scan RFID, sensor membaca kendaraan, lalu sistem menghitung biaya saat keluar.", "Flow cards: Wipe kiri ke kanan. Garis/teks bawah: Fade.");
  footer(s, "04 - Cara Kerja", "1:05");
}

function slide5() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Fitur Utama"); title(s, "Fitur yang sudah dibuat dalam prototype SpotFinder.");
  const items = [
    ["Client Portal", "Login, registrasi, profil, saldo, riwayat, booking, dan chat support.", C.cyan],
    ["Admin Dashboard", "Pantau slot, okupansi, aktivitas, revenue, analytics, dan gate control.", C.green],
    ["RFID & Wallet", "Member dapat masuk dengan RFID dan saldo parkir tercatat otomatis.", C.yellow],
    ["Booking Slot", "Pengguna bisa memesan slot kosong dengan biaya booking.", C.cyan],
    ["AI Voice", "Notifikasi suara untuk welcome, goodbye, dan status sistem.", C.green],
    ["Report & Receipt", "Riwayat transaksi, cetak struk, dan laporan operasional.", C.yellow],
  ];
  items.forEach((d, i) => card(s, 0.72 + (i % 3) * 3.95, 1.55 + Math.floor(i / 3) * 1.85, 3.45, 1.45, d[0], d[1], d[2]));
  notes(s, "Fitur utama yang saya buat meliputi portal client, dashboard admin, RFID, wallet, booking, AI voice, dan laporan.", "Cards: Pop satu per satu. Gunakan animation delay 0.1 detik antar card.");
  footer(s, "05 - Fitur Utama", "1:00");
}

function slide6() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Teknologi dan Alasan"); title(s, "Teknologi dipilih karena sesuai kebutuhan web, IoT, dan AI.");
  const items = [
    ["PHP + MySQL", "Dipilih karena ringan untuk web lokal, mudah dihubungkan ke database, dan cocok dengan XAMPP.", C.cyan],
    ["HTML, CSS, JS", "Digunakan untuk membuat tampilan dashboard yang interaktif dan mudah dibuka di browser.", C.green],
    ["ESP32 WiFi", "Dipilih karena sudah memiliki koneksi WiFi sehingga bisa berkomunikasi dengan server.", C.yellow],
    ["RFID RC522", "Digunakan untuk identifikasi member agar akses masuk lebih cepat dan otomatis.", C.orange],
    ["Ultrasonik + Servo", "Sensor membaca kendaraan di slot, sedangkan servo menggerakkan palang parkir.", C.blue],
    ["AI Voice / TTS", "Digunakan untuk notifikasi suara agar sistem terasa lebih informatif dan modern.", C.red],
  ];
  items.forEach((d, i) => card(s, 0.72 + (i % 2) * 5.85, 1.42 + Math.floor(i / 2) * 1.55, 5.15, 1.16, d[0], d[1], d[2], 10.7));
  notes(s, "Teknologi yang digunakan tidak hanya disebutkan, tetapi juga ada alasannya. PHP dan MySQL untuk web dan database, ESP32 untuk IoT, RFID untuk akses, sensor untuk deteksi, dan AI voice untuk notifikasi.", "Judul: Rise. Setiap card: Fade. AI card terakhir diberi Pop agar menonjol.");
  footer(s, "06 - Teknologi dan Alasan", "1:20");
}

function slide7() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Modal dan Waktu"); title(s, "Estimasi pengerjaan prototype SpotFinder.");
  card(s, 0.8, 1.55, 4.25, 1.55, "Modal Keseluruhan", "± Rp 500.000\nMencakup komponen, alat dan bahan, kebutuhan testing, makanan, minuman, dan pendukung pengerjaan.", C.green, 12);
  card(s, 8.15, 1.55, 4.25, 1.55, "Waktu Pengerjaan", "± 6 Bulan\nMencakup riset, coding, perakitan hardware, debugging, dan testing.", C.cyan, 12);
  text(s, "Tahapan Pengerjaan", 4.5, 3.55, 4.4, 0.35, 22, C.text, { bold: true, align: "center" });
  const timeline = [["Ide", C.red], ["Web", C.cyan], ["Database", C.green], ["IoT", C.orange], ["Testing", C.blue]];
  timeline.forEach((d, i) => {
    const x = 2.45 + i * 1.72;
    s.addShape(pptx.ShapeType.chevron, { x, y: 4.35, w: 1.22, h: 1.15, rotate: 90, fill: { color: d[1] }, line: { color: "FFFFFF", width: 1.3 } });
    text(s, d[0], x + 0.05, 4.77, 1.1, 0.22, d[0] === "Database" ? 12 : 13.5, "FFFFFF", { bold: true, align: "center" });
  });
  notes(s, "Project ini saya kerjakan sekitar enam bulan. Modalnya sekitar lima ratus ribu rupiah, termasuk komponen, bahan, testing, dan kebutuhan pendukung selama pengerjaan.", "Dua card utama: Pop. Timeline: Wipe kiri ke kanan atau Rise berurutan.");
  footer(s, "07 - Modal dan Waktu", "0:55");
}

function slide8() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Permasalahan Saat Pengerjaan"); title(s, "Beberapa kendala yang ditemui selama project.");
  card(s, 0.72, 1.48, 3.45, 1.38, "Biaya", "Komponen, alat, bahan, dan kebutuhan pendukung harus disesuaikan dengan modal yang terbatas.", C.yellow);
  card(s, 4.95, 1.48, 3.45, 1.38, "Hardware", "Sensor perlu kalibrasi karena pembacaan jarak kadang tidak stabil.", C.cyan);
  card(s, 9.18, 1.48, 3.45, 1.38, "Integrasi", "Menghubungkan ESP32, API PHP, dan database membutuhkan banyak testing.", C.green);
  card(s, 0.72, 3.45, 3.45, 1.38, "Debugging", "Error pada koneksi, status slot, booking, dan gate perlu diperbaiki bertahap.", C.red);
  card(s, 4.95, 3.45, 3.45, 1.38, "Waktu", "Project memakan waktu cukup panjang karena menggabungkan web dan IoT.", C.orange);
  card(s, 9.18, 3.45, 3.45, 1.38, "Testing", "Perlu uji coba berulang agar alur masuk, parkir, dan keluar berjalan benar.", C.blue);
  notes(s, "Selama pengerjaan ada beberapa kendala, seperti biaya, sensor yang perlu kalibrasi, integrasi ESP32 dengan website, debugging, dan testing berulang.", "Cards: Fade satu per satu. Card biaya diberi Pop karena menjadi salah satu kendala penting.");
  footer(s, "08 - Kendala Project", "1:05");
}

function slide9() {
  const s = pptx.addSlide(); bg(s); kicker(s, "Pengembangan Selanjutnya"); title(s, "Fitur yang bisa dikembangkan agar lebih maksimal.");
  const items = [
    ["QRIS Asli", "Integrasi payment gateway agar top-up dan pembayaran benar-benar cashless.", C.green],
    ["Kamera Plat Nomor", "Deteksi plat otomatis untuk meningkatkan validasi kendaraan.", C.cyan],
    ["Aplikasi Mobile", "Akses lebih mudah lewat Android/iOS untuk pengguna.", C.yellow],
    ["Cloud Deployment", "Sistem bisa diakses online, tidak hanya lokal XAMPP.", C.orange],
    ["Notifikasi WhatsApp", "Pemberitahuan booking, saldo, dan durasi parkir secara otomatis.", C.blue],
    ["AI Chatbot Lebih Pintar", "Bantuan pengguna yang lebih kontekstual dan responsif.", C.red],
  ];
  items.forEach((d, i) => card(s, 0.72 + (i % 3) * 3.95, 1.55 + Math.floor(i / 3) * 1.85, 3.45, 1.45, d[0], d[1], d[2]));
  notes(s, "Ke depannya, SpotFinder bisa dikembangkan dengan QRIS asli, kamera plat nomor, aplikasi mobile, cloud deployment, notifikasi WhatsApp, dan AI chatbot yang lebih pintar.", "Cards: Rise berurutan. Untuk QRIS dan AI Chatbot gunakan Pop agar terlihat sebagai highlight.");
  footer(s, "09 - Pengembangan", "1:00");
}

function slide10() {
  const s = pptx.addSlide(); bg(s);
  s.addImage({ path: img.logo, x: 6.12, y: 0.72, w: 0.7, h: 0.7 });
  kicker(s, "Penutup", 5.72, 1.72, 1.8);
  text(s, "SpotFinder membuat proses parkir lebih cepat, lebih tertata, dan lebih mudah dipantau.", 1.25, 2.18, 10.85, 1.2, 35, C.text, { bold: true, align: "center", fontFace: "Aptos Display" });
  text(s, "Dengan menggabungkan web, database, IoT, dan AI voice assistant, project ini menjadi prototype smart parking yang dapat terus dikembangkan.", 2.0, 4.05, 9.35, 0.72, 17, C.muted, { align: "center" });
  s.addShape(pptx.ShapeType.roundRect, { x: 5.25, y: 5.35, w: 2.85, h: 0.55, fill: { color: C.green }, line: { color: C.green } });
  text(s, "Terima kasih", 5.25, 5.51, 2.85, 0.15, 16, C.ink, { bold: true, align: "center" });
  notes(s, "Kesimpulannya, SpotFinder adalah solusi parkir cerdas yang menggabungkan web, database, IoT, dan AI agar parkir lebih cepat dan mudah dipantau. Sekian presentasi dari saya, terima kasih.", "Logo: Breathe. Judul: Fade. Tombol terima kasih: Pop.");
  footer(s, "10 - Penutup", "0:50");
}

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10].forEach((fn) => fn());

pptx.writeFile({ fileName: out }).then(() => console.log(out));
