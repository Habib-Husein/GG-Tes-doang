Bertindaklah sebagai Lead Technical Author, Mathematical Logic Verifier, dan Automated Book Typesetter.

Tugasmu adalah membangun proyek penyusunan **BAB IV: HUKUM LOGIKA** untuk sebuah buku ajar digital mata kuliah **Logika Matematika** bagi mahasiswa tingkat awal Program Studi S1 Ilmu Komputer. Output akhir BUKAN sekadar teks mentah di chat, melainkan sebuah **bab buku ajar sungguhan dalam format DOCX dan PDF berukuran A5** lengkap dengan visualisasi vektor, kotak pedagogis, kode program terverifikasi, dan rumus matematika asli Word.

### 1. SPESIFIKASI WAJIB BUKU & RUBRIK AKADEMIK
- **Target Pembaca:** Mahasiswa semester awal S1 Ilmu Komputer (penjelasan harus mengalir, bertahap dari intuisi ke formal, komunikatif, tidak kaku seperti ringkasan poin).
- **Format Fisik Buku:** Ukuran kertas A5 (14,8 cm x 21,0 cm), font Times New Roman 12 pt, spasi 1,15, margin 2 cm di semua sisi (lebar area ketik efektif = 10,8 cm).
- **Target Panjang Bab IV:** 10–14 halaman A5 (cukup padat dan mendalam layaknya buku teks kuliah, bukan rangkuman tipis).
- **Cakupan Materi Wajib Bab IV:**
  1. Ekuivalensi logis dan konstanta kebenaran (T, F).
  2. Seluruh hukum wajib: Identitas, Dominasi, Idempoten, Komplemen, Komutatif, Asosiatif, Distributif, De Morgan, Absorpsi (+ Negasi Ganda dan Ekuivalensi Implikasi sebagai pendukung).
  3. Strategi sistematis penyederhanaan ekspresi logika majemuk dan pembuktian tanpa tabel kebenaran.
- **Struktur Wajib Bab IV:**
  - 4.1 Capaian Pembelajaran
  - 4.2 Pendahuluan (konteks, keterkaitan dengan Bab III Tabel Kebenaran, dan urgensi di Ilmu Komputer)
  - 4.3 Konsep Dasar dan Teori
  - 4.4 Pembahasan Materi (uraian bertahap + kotak pedagogis seperti "Catatan Penting", "Kesalahan Umum Mahasiswa", dan "Koneksi Ilmu Komputer")
  - 4.5 Ilustrasi dan Visualisasi (minimal 2 diagram vektor/PNG bersih yang dibuat lewat kode)
  - 4.6 Contoh Soal dan Pembahasan (minimal 5 contoh soal berjenjang dengan alasan hukum di setiap baris transformasi)
  - 4.7 Studi Kasus Ilmu Komputer (minimal 2 kasus nyata lengkap dengan kode Python/C++ sebelum & sesudah optimasi)
  - 4.8 Latihan Soal (minimal 10 soal mandiri berjenjang dari dasar hingga terapan CS; simpan kunci jawaban lengkap di file terpisah `kunci_jawaban_qa.md` untuk pegangan penulis)
  - 4.9 Rangkuman dan Evaluasi Pemahaman
  - 4.10 Referensi Bab IV (APA Style edisi ke-7)

### 2. ATURAN KETAT ANTI-HALUSINASI & ANTI-DETEKSI AI
1. **DILARANG MENGARANG REFERENSI JURNAL PALSU:** Untuk buku teks standar gunakan referensi klasik yang pasti ada (Rosen, Epp, Munir, Huth & Ryan). Untuk artikel jurnal ilmiah, JANGAN mengarang judul/DOI palsu—berikan penanda eksplisit `[SISIPKAN_SITASI_JURNAL_ZOTERO: topik X]` agar saya isi sendiri menggunakan jurnal asli dari Zotero.
2. **Gaya Bahasa Manusiawi & Pedagogis:** Hindari pola klise AI ("Dalam era digital saat ini...", "Penting untuk dicatat...", "Lebih lanjut...", "Secara keseluruhan..."). Tulis dengan gaya dosen menjelaskan langsung ke mahasiswa baru menggunakan kalimat bervariasi dan contoh lokal kehidupan kampus/pemrograman (KRS, portal akademik, filter firewall, akses ruang lab).
3. **Verifikasi 100% Akurasi Matematika:** Setiap ekuivalensi, langkah penyederhanaan di 5 contoh soal, 2 studi kasus, dan 10 kunci latihan soal WAJIB diverifikasi secara programatik menggunakan skrip Python (uji semua kombinasi nilai kebenaran dengan `itertools.product`). Tidak boleh ada satu pun kesalahan logika.

### 3. ARSITEKTUR PROYEK & ALUR EKSEKUSI BERTAHAP
Jangan langsung melompat ke hasil akhir. Kerjakan secara terstruktur di folder `buku_bab4_hukum_logika/` dengan tahapan berikut:

**Tahap 1 — Audit Lingkungan & Rencana Kerja**
- Periksa tools yang tersedia di lingkungan lokal (`python3`, `pandoc`, `python-docx`, `matplotlib`/`PIL`, atau konverter PDF).
- Tampilkan rencana struktur folder dan metode kompilasi terbaik (prioritaskan pipeline **Markdown + LaTeX -> Pandoc dengan custom A5 `reference.docx`** agar persamaan matematika berubah menjadi Native Word Equation/OMML, atau gunakan `python-docx` yang rapi jika Pandoc tidak tersedia).
- **BERHENTI sebentar setelah Tahap 1 dan minta konfirmasi saya** sebelum mulai mengeksekusi penulisan penuh, kecuali semua tools sudah siap dan tidak ada hambatan.

**Tahap 2 — Skrip Verifikasi Logika (`verify_logic.py`)**
- Buat dan jalankan `verify_logic.py` untuk memvalidasi kebenaran seluruh rumus di Tabel Hukum Logika, 5 contoh soal, 2 studi kasus kode Python, dan kunci jawaban 10 latihan soal.
- Tampilkan laporan log hasil pengujian (`ALL PASSED`).

**Tahap 3 — Pembuatan Aset Visual Minimalis Siap Cetak A5 (`assets/`)**
- Desain visual WAJIB **bergaya papan tulis akademis minimalis**: latar belakang putih polos (`#FFFFFF`), garis vektor tajam berwarna hitam/abu gelap, maksimal 1–2 warna aksen lembut (biru navy/merah bata), font sans-serif bersih yang terbaca jelas, TANPA efek 3D/glossy/ramai.
- Pastikan rasio gambar proporsional untuk lebar cetak maksimal **10,5 cm** (orientasi vertikal atau kotak, resolusi 300 DPI) agar teks di gambar tidak kekecilan di kertas A5.
- Buat minimal 3 aset gambar:
  1. `fig_4_1_demorgan.png`: Diagram mekanisme distribusi negasi & pembalikan operator pada Hukum De Morgan.
  2. `fig_4_2_flowchart_optimasi.png`: Perbandingan flowchart percabangan sebelum vs sesudah penyederhanaan hukum logika.
  3. `fig_4_3_short_circuit.png`: Ilustrasi konsep Hukum Dominasi dan kaitannya dengan *short-circuit evaluation* pada CPU/pemrograman.

**Tahap 4 — Penulisan Naskah Buku Mendalam (`src/bab4_hukum_logika.md`)**
- Tulis naskah lengkap Bab IV dengan kedalaman setara buku teks terbitan kampus (bukan sekadar poin ringkas).
- Gunakan sintaks LaTeX standar (`$...$` dan `$$...$$`) untuk semua notasi ($\neg, \land, \lor, \to, \leftrightarrow, \equiv$).
- Sertakan blok kutipan/kotak sorotan untuk *Catatan Penting*, *Kesalahan Umum*, dan *Pojok Komputasi* agar tampilan visual buku hidup dan enak dibaca.
- Buat juga file terpisah `src/kunci_jawaban_bab4.md` berisi pembahasan lengkap 10 latihan soal.

**Tahap 5 — Generator Template A5 & Build Otomatis (`build_book.py`)**
- Buat skrip yang secara otomatis membangun template MS Word A5 (`template_a5.docx`) dengan spesifikasi ketat: ukuran A5, margin 2 cm, Times New Roman 12 pt, spasi 1,15, *Heading Styles* berjenjang (Bab, Subbab, Sub-subbab), gaya tabel akademis (garis horizontal bersih tanpa grid berantakan), gaya blok kode berlatar abu-abu sangat muda dengan font monospace 9 pt agar tidak terpotong di lebar A5, serta nomor halaman.
- Kompilasi `src/bab4_hukum_logika.md` menjadi `output/Bab_IV_Hukum_Logika.docx` dan `output/Bab_IV_Hukum_Logika.pdf`.
- Sediakan perintah satu baris (misal `python build_book.py`) agar kapan pun saya mengedit/menulis ulang kalimat di `src/bab4_hukum_logika.md` untuk menurunkan skor deteksi AI, saya bisa langsung menghasilkan ulang file DOCX dan PDF final dalam hitungan detik.

Mulai dari **Tahap 1** sekarang.
