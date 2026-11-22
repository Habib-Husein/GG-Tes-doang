# INSTRUKSI PROYEK LOKAL: BUKU AJAR BAB IV — HUKUM LOGIKA (POP-EDU PASTEL • GAME & AI)

Bertindaklah sebagai Penulis Buku Ajar Modern (gaya Pop-Edu), Verifikator Logika Matematika, dan Desainer Tata Letak Otomatis yang bekerja langsung di workspace lokal Windows ini.

Tugasmu adalah membangun proyek **BAB IV: HUKUM LOGIKA** untuk buku ajar mata kuliah **Logika Matematika** (S1 Ilmu Komputer) dari nol hingga menghasilkan **DOCX & PDF A5 siap cetak**, **3 diagram gerbang logika/AI (300 DPI)**, **aplikasi web kuis interaktif**, **kunci jawaban**, serta **skrip re-build 1 baris**.

---

## 1. LINGKUNGAN LOKAL & PIPELINE TEKNIS (WINDOWS)
Di komputer lokal ini, `uv` sudah tersedia dan COM `Word.Application` (WPS Office) aktif untuk ekspor DOCX ke PDF. Jalankan pipeline berikut tanpa mengubah instalasi global Windows:
1. Buat virtual environment lokal `.venv` menggunakan `uv` (Python 3.12) berisi paket: `python-docx`, `matplotlib`, `pillow`, `pypandoc-binary`, `qrcode`, `pymupdf`, dan `pywin32` (jika diperlukan COM automation).
2. Gunakan **Pandoc (`pypandoc-binary`) + Lua filter kustom + `--reference-doc=templates/template_a5.docx`** agar seluruh rumus LaTeX `$...$` dan `$$...$$` terkonversi otomatis menjadi **Native Word Equation (OMML)** dengan 0 sisa teks `$...$` mentah.
3. Lakukan post-processing DOCX via `python-docx` untuk menata kartu warna pastel, tabel, lebar gambar, serta blok kode (setiap baris kode wajib berupa paragraf terpisah dengan atribut `xml:space="preserve"` agar indentasi Python dan kurung kurawal `{` `}` C++ tidak bergeser atau terpotong di lebar A5).
4. Ekspor `output/Bab_IV_Hukum_Logika_PopEdu.docx` ke `output/Bab_IV_Hukum_Logika_PopEdu.pdf` melalui COM `Word.Application`, lalu lakukan **audit visual seluruh halaman PDF** menggunakan `pymupdf`.

---

## 2. STRUKTUR FOLDER PROYEK
Buat seluruh file secara terstruktur di dalam folder `buku_bab4_popedu/`:

- `buku_bab4_popedu/.venv/` — Virtual environment lokal via `uv` (Python 3.12)
- `buku_bab4_popedu/verify_logic.py` — Uji exhaustif $2^n$ untuk seluruh hukum, 5 contoh soal, 2 studi kasus Game/AI, 10 kunci latihan, dan 8 soal kuis HTML
- `buku_bab4_popedu/generate_assets.py` — Generator 3 diagram gerbang logika/AI (300 DPI) + gambar QR Code
- `buku_bab4_popedu/build_book.py` — Generator template A5 + kompilasi Pandoc OMML + post-process DOCX + ekspor PDF
- `buku_bab4_popedu/filters/callout_filter.lua` — Lua filter Pandoc untuk kartu Pop-Edu & caption
- `buku_bab4_popedu/assets/fig_4_1_gerbang_penyederhanaan.png` — Diagram gerbang logika sebelum vs sesudah penyederhanaan
- `buku_bab4_popedu/assets/fig_4_2_gerbang_demorgan.png` — Diagram ekuivalensi gerbang NAND/NOR (Hukum De Morgan)
- `buku_bab4_popedu/assets/fig_4_3_gerbang_boss_ai.png` — Diagram alur logika & short-circuit Boss AI
- `buku_bab4_popedu/assets/qr_kuis_bab4.png` — QR Code aktif menuju kuis web
- `buku_bab4_popedu/src/bab4_hukum_logika.md` — Naskah utama Bab IV (target 15–17 halaman A5)
- `buku_bab4_popedu/src/kunci_jawaban_bab4.md` — Pembahasan lengkap 10 latihan soal
- `buku_bab4_popedu/templates/template_a5.docx` — Reference DOCX A5 otomatis
- `buku_bab4_popedu/web_kuis/index.html` — Web kuis interaktif mandiri (1 file HTML+CSS+JS)
- `buku_bab4_popedu/output/Bab_IV_Hukum_Logika_PopEdu.docx` — Output final MS Word A5
- `buku_bab4_popedu/output/Bab_IV_Hukum_Logika_PopEdu.pdf` — Output final PDF A5

---

## 3. ATURAN WAJIB AKADEMIK (TIDAK BOLEH DILANGGAR)
- **Format Cetak:** Ukuran kertas tepat **A5 (14,8 cm × 21,0 cm)**, **margin 2 cm** di semua sisi (lebar area ketik efektif = 10,8 cm), font isi utama **Times New Roman 12 pt**, spasi **1,15**.
- **Target Panjang:** **15–17 halaman A5** (lega, banyak ruang napas/*whitespace*, tidak padat seperti tembok teks, dan tidak ada ruang kosong lebar yang janggal).
- **Cakupan Materi Wajib:**
  1. Pengertian ekuivalensi logis ($\equiv$), konstanta kebenaran ($T$ dan $F$), serta prinsip substitusi.
  2. Seluruh hukum wajib: **Identitas, Dominasi, Idempoten, Negasi Ganda, Komplemen, Komutatif, Asosiatif, Distributif, De Morgan, dan Absorpsi** (ditambah hukum pendukung: Ekuivalensi Implikasi, Negasi Implikasi, Kontraposisi, dan Ekuivalensi Biimplikasi).
  3. Teknik penyederhanaan ekspresi logika majemuk dan pembuktian ekuivalensi, tautologi, atau kontradiksi secara aljabar tanpa tabel kebenaran.
- **Struktur Subbab Wajib (`4.1` s.d. `4.10`):**
  - `4.1 Capaian Pembelajaran`
  - `4.2 Pendahuluan` (kaitan dengan Bab III Tabel Kebenaran & alasan tabel kebenaran tidak cukup di Ilmu Komputer)
  - `4.3 Konsep Dasar dan Teori`
  - `4.4 Pembahasan Materi`
  - `4.5 Ilustrasi dan Visualisasi`
  - `4.6 Contoh Soal dan Pembahasan` (minimal **5 contoh soal** berjenjang dari dasar hingga 3 variabel, mencantumkan nama hukum di setiap baris langkah)
  - `4.7 Studi Kasus Ilmu Komputer` (minimal **2 studi kasus bertema Game & AI Sederhana**)
  - `4.8 Latihan Soal` (minimal **10 soal mandiri** berjenjang: dasar, menengah, lanjut, dan terapan pemrograman)
  - `4.9 Rangkuman, Evaluasi Pemahaman & Kartu QR Kuis Interaktif` (rujukan ke bab selanjutnya wajib ditulis tepat sebagai **Bab V: Logika Predikat**, BUKAN Aturan Inferensi)
  - `4.10 Referensi Bab IV` (APA Style edisi ke-7)

---

## 4. GAYA PENULISAN: MENTOR / KAKAK TINGKAT (RINGAN & ANTI-AI)
- Gunakan **bahasa Indonesia ilmiah yang hangat, mengalir, dan komunikatif**—seperti mentor atau asisten laboratorium yang menjelaskan konsep ke mahasiswa baru supaya langsung paham, tanpa mengurangi ketepatan matematis.
- **Analogi Dulu, Rumus Kemudian:** Mulai tiap konsep inti dari analogi sehari-hari atau mekanik *game* sebelum menulis simbol formal.
- **Paragraf Pendek & Mudah Dipindai:** Batasi setiap paragraf maksimal **2–4 kalimat** (4–5 baris di A5) dan beri cetak **tebal** pada istilah/frasa kunci.
- **Larangan Frasa Klise AI:** DILARANG memakai pembuka atau penghubung klise seperti *"Dalam era digital saat ini..."*, *"Penting untuk dicatat bahwa..."*, *"Lebih lanjut..."*, *"Secara keseluruhan..."*, *"Dapat disimpulkan bahwa..."*, atau *"Di sinilah kita membutuhkan..."*.
- **Anti-Halusinasi Referensi:** Pakai buku teks klasik asli (Rosen 2019, Epp 2020, Munir 2016, Huth & Ryan 2004). Untuk artikel jurnal ilmiah, **DILARANG mengarang penulis/judul/DOI palsu**—cantumkan minimal 2 penanda eksplisit `[SISIPKAN_SITASI_JURNAL_ZOTERO: topik spesifik]` baik di badan teks maupun di Subbab 4.10.

---

## 5. DESIGN SYSTEM: POP-EDU PASTEL (100% STABIL DI WORD/WPS)
Pertahankan **latar kertas putih bersih (`#FFFFFF`)** tanpa latar gelap *full-bleed* dan tanpa elemen miring agar aman saat digabung di MS Word. Ciptakan ritme baca yang ringan dengan **4 warna pastel harmonis**:
- **Pastel Sky/Blue (Identitas Bab & Header Tabel):** Latar `#EFF6FF`, Aksen/Garis `#2563EB`, Teks Gelap `#1E3A8A`
- **Pastel Mint (Trik Cepat & Di Balik Layar Game):** Latar `#ECFDF5`, Aksen/Garis `#059669`, Teks Gelap `#065F46`
- **Pastel Butter/Amber (Analogi & Catatan Penting):** Latar `#FFFBEB`, Aksen/Garis `#D97706`, Teks Gelap `#92400E`
- **Pastel Coral/Rose (Awas Jebakan / Kesalahan Umum):** Latar `#FFF1F2`, Aksen/Garis `#E11D48`, Teks Gelap `#9F1239`

Terapkan komponen visual berikut di `build_book.py`:
1. **Chapter Opener Cerah (Halaman 1):** Kartu pembuka berlatar `#EFF6FF` dengan garis aksen atas/kiri tebal `#2563EB`, label kecil `BAB IV • LOGIKA MATEMATIKA`, judul besar `HUKUM LOGIKA`, dan subjudul ringkas. Di bawahnya langsung tampilkan **4.1 Capaian Pembelajaran** dalam kartu Pastel Mint (`#ECFDF5`).
2. **Ruang Napas & Kartu Mikro:** Atur `space_after = 7.5 pt` antar-paragraf, `space_before = 16 pt` sebelum subbab disertai garis bawah tipis `#2563EB`, serta *cell padding* lega (6–8 pt) pada kartu *Analogi Singkat*, *Trik Cepat*, *Di Balik Layar Game*, dan *Awas Jebakan!*.
3. **Tabel 4.1 & Contoh Soal Bersih:** Pada Tabel 4.1, wajib pisahkan pasangan hukum ke **baris baru** di dalam sel (jangan digabung menyamping dengan tanda `;`). Gunakan header `#EFF6FF` berteks tebal `#1E3A8A`. Aktifkan `keep_together` / `keep_with_next` pada tabel Contoh Soal 4.1–4.5 agar tidak pernah terbelah di tengah halaman.
4. **Blok Kode Rapi:** Bungkus kode Python dan C++ dalam kartu berlatar abu-biru sangat pucat (`#F8FAFC`), garis kiri `#2563EB`, font `Consolas 8.5 pt`, dengan setiap baris berupa paragraf eksplisit (`xml:space="preserve"`).
5. **Footer Minimalis:** Nomor halaman rapi di pojok kanan bawah (`Hal. 1`, dst.) beraksen `#2563EB`.

---

## 6. TEMA KONTEN: LOGIKA GAME & AI + DIAGRAM GERBANG LOGIKA
1. **2 Studi Kasus Ilmu Komputer (Wajib 100% Lulus Uji `verify_logic.py`):**
   - **Studi Kasus 1 (Python) — Optimasi Syarat *Drop Item Langka / Critical Skill* RPG:** Fungsi penentu bonus serangan memakai 4 status: `l` (*level_cukup*), `b` (*buff_aktif*), `e` (*energi_penuh*), dan `c` (*kena_curse*). Sajikan kode awal yang berbelit dan redundan akibat beberapa kali *patch* game, lalu sederhanakan langkah demi langkah dengan Hukum De Morgan, Absorpsi, Distributif, dan Idempoten menjadi satu ekspresi Python yang bersih.
   - **Studi Kasus 2 (C++) — Logika *Mode Serang Boss AI* & *Short-Circuit*:** *Boss* dalam game memutuskan status `serang_spesial` dari 4 kondisi: `h` (*hp_boss_kritis* — flag memori cepat), `p` (*pemain_dalam_jangkauan* — perhitungan jarak vektor 3D yang berat), `s` (*boss_kena_stun*), dan `d` (*perisai_pemain_aktif*). Sederhanakan ekspresi awalnya lalu gunakan Hukum Komutatif dan Dominasi (*short-circuit*) agar pengecekan ringan dievaluasi paling kiri sebelum fungsi jarak vektor yang berat.
2. **3 Diagram Gerbang Logika & Alur AI Siap Cetak A5 (300 DPI, Lebar ~10,5 cm):**
   Buat lewat `generate_assets.py` (`matplotlib`, latar `#FFFFFF`, garis tajam, aksen Pop-Edu Pastel, TANPA teks bertumpuk atau terpotong):
   - **`fig_4_1_gerbang_penyederhanaan.png`:** Gambar simbol standar gerbang logika IEEE/ANSI (bentuk D untuk **AND**, melengkung untuk **OR**, segitiga+lingkaran untuk **NOT**). Panel atas menampilkan rangkaian awal boros (4–5 gerbang untuk ekspresi seperti $(p \land q) \lor (p \land \neg q) \lor (p \land r)$); panel bawah menampilkan rangkaian hasil penyederhanaan yang menyusut drastis lengkap dengan keterangan persentase penghematan gerbang.
   - **`fig_4_2_gerbang_demorgan.png`:** Visualisasi transformasi gerbang Hukum De Morgan: **NAND $\equiv$ Negated-Input OR** ($\neg(p \land q) \equiv \neg p \lor \neg q$) dan **NOR $\equiv$ Negated-Input AND** ($\neg(p \lor q) \equiv \neg p \land \neg q$).
   - **`fig_4_3_gerbang_boss_ai.png`:** Perbandingan visual rangkaian logika dan alur *short-circuit* pada otak *Boss AI* (Studi Kasus 2) sebelum vs sesudah optimasi.

---

## 7. INOVASI WAJIB: WEB KUIS INTERAKTIF + KARTU QR CODE ("SCAN HERE")
1. Buat **`web_kuis/index.html`** (1 file lengkap HTML+CSS+JS, responsif di layar HP, palet **Pop-Edu Pastel** seragam dengan buku):
   - Berisi **8 soal pilihan ganda baru** bertema kombinasi hukum logika dan skenario *game*/AI (identifikasi hukum, jebakan De Morgan, penyederhanaan syarat *skill* RPG, dan prediksi *short-circuit* Boss AI).
   - Lengkapi dengan *progress bar*, cek jawaban instan + pembahasan langkah demi langkah, skor akhir + predikat seru (*S-Rank / A-Rank / B-Rank*), dan tombol main ulang.
2. Taruh variabel `QUIZ_URL = "https://contoh-link-kuis-anda.github.io/kuis-bab4-game/"` di baris paling atas `build_book.py`. Skrip wajib menghasilkan `assets/qr_kuis_bab4.png` (aksen warna `#1E3A8A`) dan memasangnya dalam **Kartu QR Code "SCAN HERE"** (2 kolom, latar Pastel Mint `#ECFDF5`) di Subbab `4.9`.

---

## 8. ALUR EKSEKUSI BERTAHAP
Jalankan **Tahap 1 hingga Tahap 5** secara berurutan (langsung lanjut hingga selesai apabila `uv` dan COM `Word.Application` berjalan normal):
- **Tahap 1:** Audit singkat lingkungan, buat `.venv` via `uv`, dan bangun struktur folder `buku_bab4_popedu/`.
- **Tahap 2:** Buat & jalankan `verify_logic.py` hingga **100% ALL PASSED** untuk semua hukum di Tabel 4.1, 5 contoh soal, 2 studi kasus Game/AI, 10 kunci latihan soal, dan 8 soal kuis HTML.
- **Tahap 3:** Buat & jalankan `generate_assets.py` serta cek visual seluruh gambar PNG agar bebas teks bertumpuk.
- **Tahap 4:** Susun `src/bab4_hukum_logika.md`, `src/kunci_jawaban_bab4.md`, dan `web_kuis/index.html`.
- **Tahap 5:** Buat `build_book.py` dan `filters/callout_filter.lua`, kompilasi DOCX & PDF A5, lakukan **audit visual seluruh halaman PDF dengan `pymupdf`** (pastikan 0 sisa `$...$`, 0 kurung kurawal C++ berantakan, 0 tabel contoh soal terbelah halaman, 0 ruang kosong janggal, dan panjang tepat 15–17 halaman A5).

Akhiri laporan dengan ringkasan hasil audit serta perintah 1 baris (`.venv\Scripts\python.exe build_book.py`) untuk *re-build* cepat setelah saya mengedit teks atau mengganti `QUIZ_URL`.
