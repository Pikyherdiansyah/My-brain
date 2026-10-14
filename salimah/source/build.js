const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/d300ef73-9ed2-4c76-ad4c-42923c7088f7_0dd05927-f8c7-4613-926f-ad4d4926478e/pptx/scripts/apply_theme.js");

const IMG = (n) => path.join(__dirname, "img", n);
const OUT = process.argv[2] || path.join(__dirname, "Company_Profile_Salimah_PW_DKI_Jakarta.pptx");

// Palette taken from the Salimah logo: orchid purple dominates, logo gold is the one accent.
const HEX = {
  plum: "3A0F52",     // dark slides
  purple: "5A1E78",   // brand purple
  lilac: "F3ECF7",    // light tint for cards
  ink: "2A1838",      // body text
  muted: "6B5A78",    // captions
  gold: "E9A224",     // accent, only on purple/plum or as a fill
  mist: "D9C8E6",     // soft text on dark
};

const THEME = {
  name: "Salimah Anggrek",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.ink, lt1: "FFFFFF", dk2: HEX.purple, lt2: HEX.lilac,
    accent1: HEX.purple, accent2: HEX.gold, accent3: HEX.plum, accent4: "8E5BA8",
    accent5: HEX.mist, accent6: HEX.muted, hlink: HEX.purple, folHlink: HEX.plum,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Company Profile Salimah PW DKI Jakarta";
pres.author = "Pimpinan Wilayah Salimah DKI Jakarta";
pres.company = "Persaudaraan Muslimah (Salimah)";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
// accent1..6 map to C.accent1..; lilac = background2, purple = text2
const P = { purple: C.accent1, gold: C.accent2, plum: C.accent3, orchid: C.accent4, mist: C.accent5, muted: C.accent6, ink: C.text1, lilac: C.background2, white: C.background1 };

const W = 13.333, H = 7.5, M = 0.6;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: P.white },
  objects: [
    { image: { path: IMG("logo.png"), x: W - M - 0.32, y: 0.42, w: 0.32, h: 0.53 } },
    { text: { text: "Company Profile Salimah PW DKI Jakarta", options: { x: M, y: H - 0.42, w: 6, h: 0.3, fontSize: 10, color: P.muted, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.4, w: W - 2 * M - 0.6, h: 0.75, fontFace: THEME.headFontFace, fontSize: 34, bold: true, color: P.purple, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.3, fontSize: 10, color: P.muted, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: P.plum },
  objects: [],
});

let n = 0;
const nm = (s) => `${s}-${++n}`;
const shadow = () => ({ type: "outer", color: "3A0F52", opacity: 0.18, blur: 8, offset: 2, angle: 90 });

function content(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: THEME.bodyFontFace, color: P.ink, fontSize: 14, valign: "top" }, o));
}

// ---------- 1. Cover ----------
pres.addSection({ title: "Pembuka" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Pembuka" });
  // Oval photo frame echoes the oval of the Salimah logo (identity motif).
  s.addShape(pres.shapes.OVAL, { x: 6.55, y: 0.55, w: 6.5, h: 6.4, fill: { color: P.gold }, objectName: nm("cover-ring") });
  s.addImage({ path: IMG("hero.jpg"), x: 6.65, y: 0.65, w: 6.3, h: 6.2, sizing: { type: "cover", w: 6.3, h: 6.2 }, rounding: true, altText: "Pengurus PW Salimah DKI Jakarta berfoto bersama", objectName: nm("cover-photo") });
  s.addImage({ path: IMG("logo.png"), x: M, y: 0.6, w: 0.62, h: 1.02, objectName: nm("cover-logo") });
  txt(s, "Company Profile", { x: M, y: 2.0, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: P.gold });
  txt(s, "Persaudaraan Muslimah (Salimah)", { x: M, y: 2.5, w: 5.8, h: 1.5, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: P.white, valign: "top" });
  txt(s, "Pimpinan Wilayah DKI Jakarta", { x: M, y: 4.05, w: 5.6, h: 0.5, fontSize: 22, color: P.white });
  txt(s, "Peduli Perempuan, Anak dan Keluarga Indonesia", { x: M, y: 4.75, w: 5.6, h: 0.45, fontFace: THEME.headFontFace, italic: true, fontSize: 18, color: P.mist });
  txt(s, "Periode kepengurusan 2021-2025", { x: M, y: 6.55, w: 5.6, h: 0.35, fontSize: 12, color: P.mist });
  s.addNotes("Sumber: Proposal Kegiatan Salimah PW DKI Jakarta 2025 dan Laporan Pertanggungjawaban Muswil, 5 Oktober 2025. Foto: pengurus PW Salimah DKI Jakarta (LPJ Muswil 2025).");
}

// ---------- 2. Tentang Salimah ----------
pres.addSection({ title: "Profil" });
{
  const s = content("Tentang Salimah", "Profil");
  txt(s, [
    { text: "Persaudaraan Muslimah (Salimah) adalah organisasi kemasyarakatan perempuan muslimah yang didirikan di Jakarta pada 8 Maret 2000.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Salimah PW DKI Jakarta berjalan sejak tahun 2000 bersama Pimpinan Pusat, sempat vakum lima tahun, lalu aktif kembali pada 2016. Program kerjanya berbasis rumah: dimulai dari kerabat dan tetangga, menjangkau anak, remaja, usia produktif hingga lansia.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Badan hukum: ", options: { bold: true } },
    { text: "Keputusan Menkumham No. AHU-0000395.AH.01.08 Tahun 2020 tentang Persetujuan Badan Hukum Perkumpulan Persaudaraan Muslimah." },
  ], { x: M, y: 1.55, w: 6.3, h: 4.6, fontSize: 16, paraSpaceAfter: 4 });

  const stats = [
    ["8 Maret 2000", "didirikan di Jakarta"],
    ["2.466", "titik struktur Salimah di seluruh Indonesia"],
    ["2016", "PW DKI Jakarta aktif kembali"],
  ];
  stats.forEach(([big, lab], i) => {
    const y = 1.55 + i * 1.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y, w: 5.13, h: 1.35, rectRadius: 0.12, fill: { color: i === 1 ? P.purple : P.lilac }, objectName: nm("about-card") });
    txt(s, big, { x: 7.9, y: y + 0.17, w: 4.6, h: 0.65, fontFace: THEME.headFontFace, fontSize: 32, bold: true, color: i === 1 ? P.gold : P.purple, valign: "middle" });
    txt(s, lab, { x: 7.9, y: y + 0.82, w: 4.6, h: 0.35, fontSize: 14, color: i === 1 ? P.white : P.ink });
  });
  s.addNotes("Sumber: Proposal 2025 (Kata Pengantar, Selayang Pandang) dan LPJ Muswil 2025 (Profil Salimah).");
}

// ---------- 3. Visi & Misi ----------
{
  const s = content("Visi & Misi", "Profil");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 1.55, w: 4.6, h: 5.15, rectRadius: 0.15, fill: { color: P.purple }, objectName: nm("visi-panel") });
  txt(s, "Visi", { x: M + 0.4, y: 1.9, w: 3.8, h: 0.45, fontSize: 20, bold: true, color: P.gold });
  txt(s, "Menjadi ormas perempuan pelopor dalam meningkatkan kualitas hidup perempuan, anak dan keluarga Indonesia", { x: M + 0.4, y: 2.45, w: 3.8, h: 3.6, fontFace: THEME.headFontFace, fontSize: 24, italic: true, color: P.white });

  const misi = [
    "Memperluas dan memperkokoh soliditas struktur di semua tingkatan",
    "Meningkatkan kualitas pengurus agar mampu merealisasikan visi dan misi Salimah",
    "Meningkatkan kualitas hidup anggota melalui program Salimah",
    "Meningkatkan kuantitas anggota sebagai basis massa Salimah",
    "Meningkatkan peran Salimah dalam peningkatan kualitas perempuan, anak dan keluarga di bidang dakwah, pendidikan, ekonomi, sosial dan kesehatan",
    "Memperluas dan memperkokoh kemitraan dengan pemerintah dan lembaga lain terkait isu perempuan, anak dan keluarga",
    "Mengokohkan peran Salimah sebagai ormas pelopor dalam pergerakan perempuan Indonesia",
  ];
  txt(s, "Misi", { x: 5.65, y: 1.55, w: 3, h: 0.45, fontSize: 20, bold: true, color: P.purple });
  const y0 = 2.1, step = 0.66;
  misi.forEach((m, i) => {
    const y = y0 + i * step;
    s.addShape(pres.shapes.OVAL, { x: 5.65, y: y + 0.02, w: 0.42, h: 0.42, fill: { color: P.lilac }, objectName: nm("misi-dot") });
    txt(s, String(i + 1), { x: 5.65, y: y + 0.02, w: 0.42, h: 0.42, align: "center", valign: "middle", bold: true, fontSize: 14, color: P.purple });
    txt(s, m, { x: 6.25, y, w: 6.5, h: 0.6, fontSize: 14, valign: "middle" });
  });
  s.addNotes("Sumber: Proposal 2025 (Visi & Misi) dan LPJ Muswil 2025.");
}

// ---------- 4. Program Unggulan ----------
pres.addSection({ title: "Program" });
{
  const s = content("Program Unggulan", "Program");
  txt(s, "Dijalankan melalui tiga departemen inti (Dakwah, Pendidikan & Pelatihan, Ekonomi) dan didukung Lembaga Kelengkapan Salimah.", { x: M, y: 1.25, w: 12, h: 0.4, fontSize: 14, color: P.muted });
  const progs = [
    ["KOSSUMA", "Koperasi Serba Usaha Salimah: simpan pinjam untuk rumah tangga dan usaha anggota"],
    ["FORSIL", "Forum Silaturahim Majelis Taklim Salimah"],
    ["PKPS", "Pelatihan Kepemimpinan Pengurus Salimah, jenjang 1 sampai 3"],
    ["MSI", "Mubalighah Salimah Indonesia, wadah kader dai perempuan"],
    ["P2MB", "Pos Penanggulangan Musibah dan Bencana"],
    ["Komunitas Perempuan", "KOB, Komunitas Entrepreneur Perempuan, SABILA, KhatamQu Bersama Salimah"],
  ];
  const cw = 3.55, ch = 1.55, gx = 0.3, gy = 0.3, x0 = M, yy = 1.85;
  progs.forEach(([t, d], i) => {
    const x = x0 + (i % 2) * (cw + gx), y = yy + Math.floor(i / 2) * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.1, fill: { color: P.lilac }, objectName: nm("prog-card") });
    txt(s, t, { x: x + 0.25, y: y + 0.18, w: cw - 0.5, h: 0.4, fontSize: 17, bold: true, color: P.purple });
    txt(s, d, { x: x + 0.25, y: y + 0.62, w: cw - 0.5, h: 0.85, fontSize: 13 });
  });
  // Rumah Pendidikan Salimah gets the large card: it holds six programmes.
  const bx = x0 + 2 * (cw + gx), bw = W - M - bx, bh = 3 * ch + 2 * gy;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: yy, w: bw, h: bh, rectRadius: 0.12, fill: { color: P.purple }, objectName: nm("rps-card") });
  txt(s, "Rumah Pendidikan Salimah", { x: bx + 0.35, y: yy + 0.25, w: bw - 0.7, h: 0.85, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: P.white });
  const rps = [
    ["P2AYD", "Pembinaan Anak Yatim dan Dhuafa"],
    ["BQS", "Baitul Qur'an Salimah"],
    ["SERASI", "Sekolah Pranikah Salimah Indonesia"],
    ["SISTER", "Sekolah Ibu Salimah Terpadu"],
    ["SALSA", "Sekolah Lansia Salimah"],
    ["PAUD", "Pendidikan anak usia dini"],
  ];
  rps.forEach(([a, b], i) => {
    const y = yy + 1.3 + i * 0.64;
    txt(s, a, { x: bx + 0.35, y, w: 1.25, h: 0.5, fontSize: 15, bold: true, color: P.gold, valign: "middle" });
    txt(s, b, { x: bx + 1.65, y, w: bw - 2.0, h: 0.5, fontSize: 14, color: P.white, valign: "middle" });
  });
  s.addNotes("Sumber: Proposal 2025 (Program Unggulan) dan LPJ Muswil 2025 (laporan departemen).");
}

// ---------- 5. Struktur Organisasi ----------
pres.addSection({ title: "Organisasi" });
{
  const s = content("Struktur Organisasi PW DKI Jakarta", "Organisasi");
  const box = (label, name, x, y, w, dark) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.95, rectRadius: 0.08, fill: { color: dark ? P.purple : P.lilac }, objectName: nm("org-box") });
    txt(s, [
      { text: label, options: { bold: true, fontSize: 14, color: dark ? P.white : P.purple, breakLine: true } },
      { text: name, options: { fontSize: 12, color: dark ? P.mist : P.ink } },
    ], { x: x + 0.12, y: y + 0.08, w: w - 0.24, h: 0.8, align: "center", valign: "middle" });
  };
  const line = (x1, y1, x2, y2) => s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: P.orchid, width: 1.25 }, objectName: nm("org-line") });

  const cx = W / 2;
  box("Ketua", "dr. Yulia Andani Murti, M.K.K.K", cx - 1.9, 1.45, 3.8, true);
  // Sekretaris & Bendahara
  line(cx, 2.4, cx, 2.75);
  line(cx - 3.0, 2.75, cx + 3.0, 2.75);
  line(cx - 3.0, 2.75, cx - 3.0, 2.95); line(cx + 3.0, 2.75, cx + 3.0, 2.95);
  box("Sekretaris", "Desi Lisnawati, S.Pd., MM", cx - 4.9, 2.95, 3.8, false);
  box("Bendahara", "Suprih Murdjiyati", cx + 1.1, 2.95, 3.8, false);
  // Departments
  line(cx, 2.75, cx, 4.3);
  const depts = [
    ["Dakwah", "Iin Setiyani"],
    ["Pendidikan & Latihan", "Eka Dewi S., S.Psi., M.Sc"],
    ["Ekonomi", "Ninik Riwahayuanti, S.Pd"],
    ["Bangda & SDM", "RR. Tri Rizalina RS"],
    ["Humas & Kajian", "Emy Dharmawaty"],
  ];
  const dw = 2.3, dg = 0.18, dx0 = (W - (5 * dw + 4 * dg)) / 2;
  line(dx0 + dw / 2, 4.3, dx0 + 4 * (dw + dg) + dw / 2, 4.3);
  depts.forEach(([d, nmx], i) => {
    const x = dx0 + i * (dw + dg);
    line(x + dw / 2, 4.3, x + dw / 2, 4.5);
    box("Dept. " + d, nmx, x, 4.5, dw, false);
  });
  // Dewan Pertimbangan & LKS
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.85, w: 5.9, h: 0.9, rectRadius: 0.08, line: { color: P.orchid, width: 1 }, fill: { color: P.white }, objectName: nm("dps") });
  txt(s, [
    { text: "Dewan Pertimbangan Salimah Wilayah", options: { bold: true, color: P.purple, breakLine: true } },
    { text: "Ketua: Siti Nurlailah, S.Pd" },
  ], { x: M + 0.2, y: 5.92, w: 5.5, h: 0.76, fontSize: 13, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: W - M - 5.9, y: 5.85, w: 5.9, h: 0.9, rectRadius: 0.08, line: { color: P.orchid, width: 1 }, fill: { color: P.white }, objectName: nm("lks") });
  txt(s, [
    { text: "Lembaga Kelengkapan Salimah (LKS)", options: { bold: true, color: P.purple, breakLine: true } },
    { text: "Lembaga Wakaf Salimah (LWS) dan Koperasi Salimah (Kossuma)" },
  ], { x: W - M - 5.7, y: 5.92, w: 5.5, h: 0.76, fontSize: 13, valign: "middle" });
  s.addNotes("Sumber: LPJ Muswil 2025 (Susunan Pengurus, Struktur Organisasi, Dewan Pertimbangan, LKS).");
}

// ---------- 6. Jumlah Pengurus ----------
{
  const s = content("Jumlah Pengurus", "Organisasi");
  s.addShape(pres.shapes.OVAL, { x: M, y: 1.6, w: 4.4, h: 4.4, fill: { color: P.purple }, objectName: nm("pengurus-oval") });
  txt(s, "282", { x: M, y: 2.55, w: 4.4, h: 1.35, align: "center", valign: "middle", fontFace: THEME.headFontFace, fontSize: 80, bold: true, color: P.gold });
  txt(s, "pengurus Salimah tercatat\ndi DKI Jakarta", { x: M + 0.5, y: 3.95, w: 3.4, h: 0.85, align: "center", fontSize: 16, color: P.white });
  txt(s, "Data sampai September 2025, dari PW hingga Pimpinan Ranting.", { x: M, y: 6.2, w: 4.6, h: 0.5, fontSize: 12, color: P.muted });

  const units = ["Dakwah", "Diklat", "Ekonomi", "Bangda & SDM", "Humas & Kajian", "Bendahara", "Sekretariat", "Ketua"];
  const counts = [5, 5, 7, 5, 4, 4, 3, 1];
  s.addChart(pres.charts.BAR, [{ name: "Pengurus PW", labels: units, values: counts }], {
    x: 5.6, y: 1.5, w: 7.1, h: 5.3, barDir: "bar",
    showTitle: true, title: "Komposisi pengurus Pimpinan Wilayah (34 orang)", titleFontSize: 15, titleColor: HEX.purple, titleFontFace: "+mj-lt",
    chartColors: [HEX.purple], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60,
    objectName: nm("pengurus-chart"),
  });
  s.addNotes("Sumber: LPJ Muswil 2025, Komposisi Pengurus (282 orang, data sampai September 2025). Komposisi PW dihitung dari Susunan Pengurus pada LPJ: ketua departemen ditambah timnya. Dewan Pertimbangan (4 orang) tidak termasuk.");
}

// ---------- 7. Jumlah Anggota ----------
{
  const s = content("Jumlah Anggota", "Organisasi");
  txt(s, "32.118", { x: M, y: 1.6, w: 4.5, h: 1.2, fontFace: THEME.headFontFace, fontSize: 66, bold: true, color: P.purple });
  txt(s, "anggota Salimah tercatat di DKI Jakarta sampai tahun 2025", { x: M, y: 2.85, w: 4.3, h: 0.8, fontSize: 16 });
  const rows = [["PD Jakarta Selatan", "10.229"], ["PD Jakarta Timur", "8.875"], ["PD Jakarta Barat", "5.464"], ["Pimpinan Wilayah", "1.901"], ["PD Jakarta Pusat", "1.003"], ["PD Jakarta Utara", "420"]];
  rows.forEach(([a, b], i) => {
    const y = 3.85 + i * 0.45;
    txt(s, a, { x: M, y, w: 3, h: 0.4, fontSize: 14, valign: "middle" });
    txt(s, b, { x: M + 3.0, y, w: 1.4, h: 0.4, fontSize: 14, bold: true, align: "right", color: P.purple, valign: "middle" });
  });
  s.addChart(pres.charts.BAR, [{ name: "Anggota", labels: ["Jakarta Selatan", "Jakarta Timur", "Jakarta Barat", "Pimpinan Wilayah", "Jakarta Pusat", "Jakarta Utara"], values: [10229, 8875, 5464, 1901, 1003, 420] }], {
    x: 5.6, y: 1.5, w: 7.1, h: 5.3, barDir: "col",
    showTitle: true, title: "Sebaran anggota per pimpinan", titleFontSize: 15, titleColor: HEX.purple, titleFontFace: "+mj-lt",
    chartColors: [HEX.purple], showValue: false,
    catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 50,
    objectName: nm("anggota-chart"),
  });
  s.addNotes("Sumber: LPJ Muswil 2025, Keanggotaan Salimah tercatat sampai dengan tahun 2025.");
}

// ---------- 8. Jaringan Struktur ----------
{
  const s = content("Jaringan Struktur di DKI Jakarta", "Organisasi");
  const big = [["6", "Pimpinan Daerah", "setingkat kota/kabupaten"], ["21", "Pimpinan Cabang", "setingkat kecamatan"], ["17", "Pimpinan Ranting", "setingkat kelurahan"]];
  big.forEach(([v, l, sub], i) => {
    const x = M + i * 4.1;
    s.addShape(pres.shapes.OVAL, { x, y: 1.5, w: 1.3, h: 1.3, fill: { color: i === 0 ? P.purple : P.lilac }, objectName: nm("net-oval") });
    txt(s, v, { x, y: 1.5, w: 1.3, h: 1.3, align: "center", valign: "middle", fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: i === 0 ? P.gold : P.purple });
    txt(s, [{ text: l, options: { bold: true, fontSize: 17, breakLine: true } }, { text: sub, options: { fontSize: 13, color: P.muted } }], { x: x + 1.5, y: 1.7, w: 2.4, h: 0.9, valign: "middle" });
  });
  const hdr = (t) => ({ text: t, options: { bold: true, color: P.white, fill: { color: P.purple }, fontSize: 13 } });
  const rows = [
    [hdr("Pimpinan Daerah"), hdr("Pimpinan Cabang"), hdr("Pimpinan Ranting")],
    ["Jakarta Timur", "7", "9"],
    ["Jakarta Barat", "7", "1"],
    ["Jakarta Selatan", "5", "6"],
    ["Jakarta Utara", "1", "1"],
    ["Jakarta Pusat", "1", "0"],
    ["Kepulauan Seribu", "0", "0"],
  ].map((r, i) => i === 0 ? r : r.map((c, j) => ({ text: c, options: { fontSize: 13, color: P.ink, bold: j === 0, align: j === 0 ? "left" : "center", fill: { color: i % 2 ? P.white : P.lilac } } })));
  rows[0][1].options.align = "center"; rows[0][2].options.align = "center";
  s.addTable(rows, { x: M, y: 3.25, w: 7.4, colW: [3.2, 2.1, 2.1], rowH: 0.42, border: { type: "solid", pt: 0.5, color: "D9C8E6" }, fontFace: THEME.bodyFontFace, margin: [0.04, 0.12, 0.04, 0.12], objectName: nm("net-table") });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.4, y: 3.25, w: W - M - 8.4, h: 2.94, rectRadius: 0.1, fill: { color: P.lilac }, objectName: nm("net-note") });
  txt(s, [
    { text: "Ruang tumbuh", options: { bold: true, color: P.purple, fontSize: 16, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "PD Kepulauan Seribu dilantik di Pulau Pramuka pada 9 Desember 2023 dan belum memiliki cabang.", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Penguatan struktur di semua tingkatan adalah misi pertama Salimah." },
  ], { x: 8.65, y: 3.45, w: W - M - 8.9, h: 2.6, fontSize: 14 });
  s.addNotes("Sumber: LPJ Muswil 2025, Sebaran Salimah Wilayah DKI Jakarta. Catatan: proposal Januari 2025 mencatat 6 PD, 19 PC, 24 PRa (49 titik); halaman Perbandingan pada LPJ menulis 8 PC untuk Jakarta Timur. Angka di slide memakai tabel Sebaran LPJ.");
}

// ---------- 9. Capaian ----------
pres.addSection({ title: "Capaian" });
{
  const s = content("Capaian Periode 2021-2025", "Capaian");
  const items = [
    ["37", "Baitul Qur'an Salimah", "dengan 521 santri terdata"],
    ["15.000+", "paket Jumat Berkah", "tersalurkan selama empat tahun"],
    ["400+", "peserta Kampus Salimah", "kajian muslimah perkantoran daring"],
    ["205+", "peserta PKPS 1-3", "pelatihan kepemimpinan pengurus"],
    ["213", "rapat Dewan Pimpinan Harian", "ditambah 25 rapat pleno"],
    ["Rp 9,5 jt", "dana wakaf terhimpun", "melalui Lembaga Wakaf Salimah"],
  ];
  const cw = 3.85, ch = 2.35, g = 0.3;
  items.forEach(([v, l, sub], i) => {
    const x = M + (i % 3) * (cw + g), y = 1.5 + Math.floor(i / 3) * (ch + g);
    const dark = i === 1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.12, fill: { color: dark ? P.purple : P.lilac }, objectName: nm("cap-card") });
    txt(s, v, { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.9, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: dark ? P.gold : P.purple, valign: "middle" });
    txt(s, l, { x: x + 0.3, y: y + 1.2, w: cw - 0.6, h: 0.45, fontSize: 16, bold: true, color: dark ? P.white : P.ink });
    txt(s, sub, { x: x + 0.3, y: y + 1.65, w: cw - 0.6, h: 0.45, fontSize: 13, color: dark ? P.mist : P.muted });
  });
  s.addNotes("Sumber: LPJ Muswil 2025. BQS 37 titik / 521 santri; Jumat Berkah sekitar 15.000 paket; Kampus Salimah lebih dari 400 peserta; PKPS 1 (40 pada 2021 dan 40 pada 2024), PKPS 2 (77), PKPS 3 (48) = 205 peserta pada sesi yang tercatat jumlahnya; sesi PKPS 1 Jakarta Utara 2022 tidak mencantumkan jumlah; 213 rapat DPH dan 25 rapat pleno; wakaf LWS Rp 9.503.000 (2021-2025).");
}

// ---------- 10. Dokumentasi Kegiatan ----------
{
  const s = content("Dokumentasi Kegiatan", "Capaian");
  const ph = [
    ["pkps.jpg", "Pelatihan Kepemimpinan Pengurus (PKPS)"],
    ["jumat.jpg", "Jumat Berkah"],
    ["bqs.jpg", "Baitul Qur'an Salimah (BQS)"],
    ["lansia.jpg", "Sekolah Lansia Salimah (SALSA)"],
    ["takjil.jpg", "Berbagi takjil Ramadhan"],
    ["maulid.jpg", "Peringatan Hari Besar Islam"],
  ];
  const cw = 3.85, ih = 2.05, g = 0.3;
  ph.forEach(([f, cap], i) => {
    const x = M + (i % 3) * (cw + g), y = 1.45 + Math.floor(i / 3) * (ih + 0.55);
    s.addImage({ path: IMG(f), x, y, w: cw, h: ih, sizing: { type: "cover", w: cw, h: ih }, altText: cap, objectName: nm("doc-photo") });
    txt(s, cap, { x, y: y + ih + 0.08, w: cw, h: 0.35, fontSize: 12, color: P.muted });
  });
  s.addNotes("Foto kegiatan dari Proposal 2025 (Kegiatan Salimah Tahun 2024).");
}

// ---------- 11. Mitra ----------
pres.addSection({ title: "Kemitraan" });
{
  const s = content("Mitra & Jejaring", "Kemitraan");
  const groups = [
    ["Pemerintah & lembaga negara", ["Kementerian Agama", "Kementerian Koperasi dan UKM", "BPOM", "DPR/MPR", "DPRD DKI Jakarta", "Kesbangpol DKI Jakarta", "MUI"]],
    ["Zakat & kemanusiaan", ["BAZNAS", "BAZIS DKI Jakarta", "Rumah Zakat", "YAKESMA", "Dompet Dhuafa", "Inisiatif Zakat Indonesia", "Adara Relief International"]],
    ["Ormas, pendidikan & media", ["KOWANI", "BMIWI", "KPIPA", "HIMPAUDI", "Pegadaian Syariah", "Republika", "Channel Muslim", "Persis Media"]],
  ];
  const cw = 3.85, g = 0.3;
  groups.forEach(([t, list], i) => {
    const x = M + i * (cw + g);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.5, w: cw, h: 5.2, rectRadius: 0.12, fill: { color: i === 0 ? P.purple : P.lilac }, objectName: nm("mitra-card") });
    txt(s, t, { x: x + 0.3, y: 1.75, w: cw - 0.6, h: 0.5, fontSize: 17, bold: true, color: i === 0 ? P.gold : P.purple });
    txt(s, list.map((m, k) => ({ text: m, options: { bullet: true, breakLine: k < list.length - 1 } })), { x: x + 0.3, y: 2.4, w: cw - 0.6, h: 4.1, fontSize: 16, color: i === 0 ? P.white : P.ink, paraSpaceAfter: 14 });
  });
  s.addNotes("Sumber: Proposal 2025 (Mitra Salimah, kemitraan tingkat Pimpinan Pusat) dan LPJ Muswil 2025 (Perkembangan Jaringan, kemitraan departemen PW DKI).");
}

// ---------- 12. Arah ke depan ----------
{
  const s = content("Arah Pengembangan", "Kemitraan");
  s.addImage({ path: IMG("ekonomi.jpg"), x: M, y: 1.5, w: 4.4, h: 4.4, sizing: { type: "cover", w: 4.4, h: 4.4 }, rounding: true, altText: "Pelatihan Wanita Berdaya Keluarga Sejahtera", objectName: nm("arah-photo") });
  txt(s, "Pelatihan \"Wanita Berdaya Keluarga Sejahtera\", 28 Juli 2024", { x: M, y: 6.05, w: 4.4, h: 0.5, fontSize: 12, color: P.muted, align: "center" });
  const pts = [
    ["Sekretariat tetap", "Kantor sekretariat yang representatif sebagai pusat kegiatan Salimah di Jakarta."],
    ["Kapasitas pengurus", "PKPS, workshop, coaching dan literasi digital; kaderisasi yang terencana."],
    ["Perluasan struktur", "Pendampingan PD Kepulauan Seribu dan cabang serta ranting baru."],
    ["Kemandirian dana", "Kewirausahaan sosial, bazar, pelatihan berbayar dan donatur tetap."],
    ["Kemitraan", "Komunikasi rutin dengan Pemprov DKI dan kolaborasi dengan organisasi sevisi."],
  ];
  pts.forEach(([t, d], i) => {
    const y = 1.5 + i * 0.98;
    s.addShape(pres.shapes.OVAL, { x: 5.6, y: y + 0.08, w: 0.32, h: 0.32, fill: { color: P.gold }, objectName: nm("arah-dot") });
    txt(s, [{ text: t, options: { bold: true, fontSize: 17, color: P.purple, breakLine: true } }, { text: d, options: { fontSize: 14 } }], { x: 6.15, y, w: W - M - 6.15, h: 0.9 });
  });
  s.addNotes("Sumber: LPJ Muswil 2025 (Solusi dan Rekomendasi) dan Proposal 2025 (Rencana Kegiatan Tahun 2025).");
}

// ---------- 13. Kontak ----------
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Kemitraan" });
  s.addShape(pres.shapes.OVAL, { x: 7.35, y: 0.95, w: 5.6, h: 5.6, fill: { color: P.gold }, objectName: nm("close-ring") });
  s.addImage({ path: IMG("forsil.jpg"), x: 7.45, y: 1.05, w: 5.4, h: 5.4, sizing: { type: "cover", w: 5.4, h: 5.4 }, rounding: true, altText: "Kegiatan bersama masyarakat", objectName: nm("close-photo") });
  s.addImage({ path: IMG("logo.png"), x: M, y: 0.6, w: 0.62, h: 1.02, objectName: nm("close-logo") });
  txt(s, "Mari berkolaborasi", { x: M, y: 2.0, w: 6.4, h: 0.9, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: P.white });
  txt(s, "untuk perempuan, anak dan keluarga Jakarta.", { x: M, y: 2.9, w: 6.4, h: 0.5, fontSize: 20, color: P.mist });
  txt(s, [
    { text: "Pimpinan Wilayah Salimah DKI Jakarta", options: { bold: true, color: P.gold, breakLine: true } },
    { text: "TU Center, Jl. H. Nawi Raya RT 3/RW 11, Gandaria Utara,", options: { breakLine: true } },
    { text: "Kebayoran Baru, Jakarta Selatan", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Salimah Call Center: 0822 1363 0747" },
  ], { x: M, y: 4.2, w: 6.4, h: 1.8, fontSize: 16, color: P.white });
  s.addNotes("Sumber alamat dan telepon: LPJ Muswil 2025 (Profil Salimah, Contact Us). Proposal Januari 2025 mencantumkan alamat Jl. Mampang Prapatan VIII no R1; slide memakai alamat terbaru dari LPJ.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
