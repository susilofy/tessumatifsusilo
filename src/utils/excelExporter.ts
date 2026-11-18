import * as XLSX from 'xlsx';
import { CONFIG } from '../config';
import {
  ExamResult,
  Question,
  QuestionItemAnalysis,
  StudentItemResponse,
} from '../types';
import { formatCorrectAnswer } from '../data/blueprintData';
import { computeItemAnalysis } from './itemAnalysisHelper';

type CellValue = string | number | boolean | null | undefined;

/**
 * Menghitung lebar kolom otomatis berdasarkan panjang teks header & isi
 */
function buildAutoColWidths(
  rows: CellValue[][],
  minWidth = 8,
  maxWidth = 42,
  compactFromColIdx?: number,
  compactCount = 0
): XLSX.ColInfo[] {
  if (rows.length === 0) return [];
  const colCount = Math.max(...rows.map((r) => r.length));
  const widths: XLSX.ColInfo[] = [];

  for (let c = 0; c < colCount; c++) {
    const isCompactItemCol =
      compactFromColIdx !== undefined &&
      c >= compactFromColIdx &&
      c < compactFromColIdx + compactCount;

    let maxLen = 0;
    for (let r = 0; r < rows.length; r++) {
      const val = rows[r][c];
      const strLen = val !== null && val !== undefined ? String(val).length : 0;
      if (strLen > maxLen) maxLen = strLen;
    }

    const colMin = isCompactItemCol ? 5 : minWidth;
    const wch = Math.min(maxWidth, Math.max(colMin, maxLen + 3));
    widths.push({ wch });
  }

  return widths;
}

/**
 * Membentuk worksheet dari array 2D dengan header standar di baris 1, autofilter, dan lebar kolom otomatis
 */
function createStandardSheet(
  aoa: CellValue[][],
  options?: {
    enableAutoFilter?: boolean;
    filterLastRowIndex?: number;
    minWidth?: number;
    maxWidth?: number;
    compactFromColIdx?: number;
    compactCount?: number;
  }
): XLSX.WorkSheet {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = buildAutoColWidths(
    aoa,
    options?.minWidth ?? 8,
    options?.maxWidth ?? 42,
    options?.compactFromColIdx,
    options?.compactCount ?? 0
  );

  if (options?.enableAutoFilter && aoa.length > 0 && aoa[0].length > 0) {
    const lastRow =
      options.filterLastRowIndex !== undefined
        ? Math.max(0, options.filterLastRowIndex)
        : aoa.length - 1;
    const lastCol = aoa[0].length - 1;
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: lastRow, c: lastCol },
      }),
    };
  }

  return ws;
}

/**
 * Memastikan nama sheet valid untuk Excel (maks 31 karakter, tanpa karakter terlarang)
 */
function sanitizeSheetName(name: string): string {
  return name.replace(/[\\/?*[\]:]/g, '').slice(0, 31);
}

/**
 * Format label bentuk soal
 */
function formatQuestionTypeLabel(type: Question['type']): string {
  if (type === 'pg') return 'Pilihan Ganda';
  if (type === 'pgk') return 'PG Kompleks';
  return 'Benar / Salah';
}

/**
 * Format status pelanggaran & submit
 */
function getIntegrityFields(r: ExamResult) {
  const isAutoSubmit = r.submissionReason === 'violation_limit';
  const statusSubmit = isAutoSubmit
    ? 'Auto-Submit (Batas Pelanggaran 3x)'
    : r.submissionReason === 'time_out'
    ? 'Auto-Submit (Waktu Habis)'
    : 'Submit Normal';

  const catatan =
    r.terminationNote ||
    (r.violationsCount && r.violationsCount > 0
      ? `${r.violationsCount}x peringatan pindah tab`
      : 'Tertib');

  const logHistory =
    r.violationLogs && r.violationLogs.length > 0
      ? r.violationLogs.map((l) => `${l.timestamp} (${l.reason})`).join('; ')
      : '-';

  return { statusSubmit, catatan, logHistory };
}

/**
 * Membangun sheet "Rekap Nilai" (Tabel Standar Header Baris 1)
 */
function buildRekapNilaiSheet(results: ExamResult[]): XLSX.WorkSheet {
  const headers = [
    'No',
    'Timestamp',
    'Nama Siswa',
    'Mata Pelajaran',
    'Kelas',
    'Nomor Absen',
    'Tanggal Lahir',
    'Benar',
    'Salah',
    'Nilai',
    'Status KKTP',
    'Durasi Pengerjaan',
    'Jumlah Pelanggaran',
    'Status Pengerjaan',
    'Catatan Integritas',
  ];

  const rows: CellValue[][] = results.map((r, idx) => {
    const isLulus = r.status === 'Lulus' || (r.nilai ?? 0) >= CONFIG.KKTP;
    const { statusSubmit, catatan } = getIntegrityFields(r);

    return [
      idx + 1,
      r.timestamp || '-',
      r.nama || '-',
      r.mataPelajaran || CONFIG.MATA_PELAJARAN,
      r.kelas || CONFIG.KELAS,
      r.noAbsen || '-',
      r.tglLahir || '-',
      Number(r.benar ?? 0),
      Number(r.salah ?? 0),
      Number(r.nilai ?? 0),
      isLulus ? 'Tuntas' : 'Remedial',
      r.durasiPengerjaan || '-',
      Number(r.violationsCount ?? 0),
      statusSubmit,
      catatan,
    ];
  });

  const aoa = [headers, ...rows];
  return createStandardSheet(aoa, {
    enableAutoFilter: true,
    minWidth: 10,
    maxWidth: 44,
  });
}

/**
 * Membangun sheet "Ringkasan Statistik" (Tabel Standar Header Baris 1)
 */
function buildRingkasanStatistikSheet(results: ExamResult[]): XLSX.WorkSheet {
  const total = results.length;
  const lulusList = results.filter((r) => r.status === 'Lulus' || (r.nilai ?? 0) >= CONFIG.KKTP);
  const lulusCount = lulusList.length;
  const remedialCount = total - lulusCount;
  const persenLulus = total > 0 ? Math.round((lulusCount / total) * 1000) / 10 : 0;
  const persenRemedial = total > 0 ? Math.round((remedialCount / total) * 1000) / 10 : 0;

  const avgNilai =
    total > 0
      ? Math.round((results.reduce((sum, r) => sum + (Number(r.nilai) || 0), 0) / total) * 100) /
        100
      : 0;

  const sortedByScore = [...results].sort((a, b) => (b.nilai ?? 0) - (a.nilai ?? 0));
  const highestStudent = sortedByScore[0];
  const lowestStudent = sortedByScore[sortedByScore.length - 1];

  const totalViolations = results.reduce((sum, r) => sum + (Number(r.violationsCount) || 0), 0);
  const autoSubmitViolationCount = results.filter(
    (r) => r.submissionReason === 'violation_limit'
  ).length;

  const sangatBaikCount = results.filter((r) => (r.nilai ?? 0) >= 90).length;
  const baikTuntasCount = results.filter(
    (r) => (r.nilai ?? 0) >= CONFIG.KKTP && (r.nilai ?? 0) < 90
  ).length;
  const perluRemedialCount = results.filter((r) => (r.nilai ?? 0) < CONFIG.KKTP).length;

  const aoa: CellValue[][] = [
    ['Indikator Statistik', 'Nilai', 'Keterangan'],
    ['Sekolah', CONFIG.SEKOLAH, CONFIG.ALAMAT_SEKOLAH],
    ['Mata Pelajaran Utama', CONFIG.MATA_PELAJARAN, `Materi: ${CONFIG.MATERI}`],
    ['Kelas & KKTP Acuan', CONFIG.KELAS, `KKTP Minimal: ${CONFIG.KKTP}`],
    ['Total Peserta Ujian', total, 'Seluruh data siswa terekam'],
    ['Jumlah Tuntas (Lulus KKTP)', lulusCount, `${persenLulus}% dari total peserta`],
    ['Jumlah Remedial (Belum Tuntas)', remedialCount, `${persenRemedial}% dari total peserta`],
    ['Rata-Rata Nilai Kelas', avgNilai, `Batas KKTP: ${CONFIG.KKTP}`],
    [
      'Nilai Tertinggi',
      highestStudent ? Number(highestStudent.nilai ?? 0) : 0,
      highestStudent ? `Diraih oleh: ${highestStudent.nama} (Absen ${highestStudent.noAbsen})` : '-',
    ],
    [
      'Nilai Terendah',
      lowestStudent ? Number(lowestStudent.nilai ?? 0) : 0,
      lowestStudent ? `Diraih oleh: ${lowestStudent.nama} (Absen ${lowestStudent.noAbsen})` : '-',
    ],
    ['Distribusi Nilai 90 - 100 (Sangat Baik)', sangatBaikCount, 'Kategori Pengayaan'],
    [`Distribusi Nilai ${CONFIG.KKTP} - 89 (Tuntas)`, baikTuntasCount, 'Memenuhi KKTP'],
    [`Distribusi Nilai < ${CONFIG.KKTP} (Remedial)`, perluRemedialCount, 'Perlu Tindak Lanjut'],
    ['Total Kejadian Pindah Tab (Pelanggaran)', totalViolations, 'Akumulasi log keamanan siswa'],
    [
      'Siswa Auto-Submit (Batas Pelanggaran 3x)',
      autoSubmitViolationCount,
      'Dihentikan otomatis sistem',
    ],
  ];

  // Tambahkan ringkasan per Mata Pelajaran bila lebih dari 1 mapel
  const mapelGroups = new Map<string, ExamResult[]>();
  results.forEach((r) => {
    const key = r.mataPelajaran || CONFIG.MATA_PELAJARAN;
    const list = mapelGroups.get(key) || [];
    list.push(r);
    mapelGroups.set(key, list);
  });

  if (mapelGroups.size > 1) {
    mapelGroups.forEach((list, mapelName) => {
      const mAvg =
        Math.round((list.reduce((s, r) => s + (Number(r.nilai) || 0), 0) / list.length) * 100) /
        100;
      const mPass = list.filter((r) => r.status === 'Lulus' || (r.nilai ?? 0) >= CONFIG.KKTP).length;
      aoa.push([
        `Rekap Mapel: ${mapelName}`,
        mAvg,
        `${list.length} peserta • ${mPass} tuntas`,
      ]);
    });
  }

  return createStandardSheet(aoa, {
    enableAutoFilter: true,
    minWidth: 12,
    maxWidth: 48,
  });
}

/**
 * Membangun sheet Matriks Jawaban Per Butir (Tabel Standar Header Baris 1)
 */
function buildItemMatrixSheet(
  results: ExamResult[],
  questions: Question[],
  itemStats: QuestionItemAnalysis[],
  allResponses: Map<string, Record<number, StudentItemResponse>>
): XLSX.WorkSheet {
  const headers: CellValue[] = [
    'No',
    'Nomor Absen',
    'Nama Siswa',
    'Kelas',
    ...questions.map((_, i) => `S${i + 1}`),
    'Total Skor',
    'Nilai Akhir',
    'Status KKTP',
    'Durasi Pengerjaan',
    'Jumlah Pelanggaran',
    'Waktu Submit',
  ];

  const studentRows: CellValue[][] = results.map((r, idx) => {
    const key = `${r.nama}_${r.noAbsen}_${r.timestamp}`;
    const studentResponses = allResponses.get(key) || {};
    const itemScores = questions.map((q) => {
      const resp = studentResponses[q.id];
      if (!resp) return 0;
      if (resp.status === 'benar') return 1;
      if (resp.status === 'sebagian') return Number(resp.score);
      return 0;
    });

    const isLulus = r.status === 'Lulus' || (r.nilai ?? 0) >= CONFIG.KKTP;

    return [
      idx + 1,
      r.noAbsen || '-',
      r.nama || '-',
      r.kelas || CONFIG.KELAS,
      ...itemScores,
      Number(r.benar ?? 0),
      Number(r.nilai ?? 0),
      isLulus ? 'Tuntas' : 'Remedial',
      r.durasiPengerjaan || '-',
      Number(r.violationsCount ?? 0),
      r.timestamp || '-',
    ];
  });

  const summaryRows: CellValue[][] = [
    [
      '',
      '',
      'KUNCI JAWABAN',
      '',
      ...questions.map((q) => formatCorrectAnswer(q)),
      '',
      '',
      '',
      '',
      '',
      '',
    ],
    [
      '',
      '',
      'JUMLAH BENAR (Σ)',
      '',
      ...itemStats.map((s) => s.correctCount),
      '',
      '',
      '',
      '',
      '',
      '',
    ],
    [
      '',
      '',
      'KETERCAPAIAN (%)',
      '',
      ...itemStats.map((s) => s.persentaseKetercapaian),
      '',
      '',
      '',
      '',
      '',
      '',
    ],
    [
      '',
      '',
      'TINGKAT KESUKARAN',
      '',
      ...itemStats.map((s) => s.klasifikasiKesukaran),
      '',
      '',
      '',
      '',
      '',
      '',
    ],
    [
      '',
      '',
      'DAYA PEMBEDA (D)',
      '',
      ...itemStats.map((s) => s.dayaPembeda),
      '',
      '',
      '',
      '',
      '',
      '',
    ],
  ];

  const aoa = [headers, ...studentRows, ...summaryRows];
  return createStandardSheet(aoa, {
    enableAutoFilter: true,
    filterLastRowIndex: studentRows.length,
    minWidth: 8,
    maxWidth: 36,
    compactFromColIdx: 4,
    compactCount: questions.length,
  });
}

/**
 * Membangun sheet "Statistik Butir Soal" (Tabel Standar Header Baris 1)
 */
function buildStatistikButirSheet(
  questions: Question[],
  itemStats: QuestionItemAnalysis[]
): XLSX.WorkSheet {
  const headers = [
    'No Soal',
    'Elemen / CP',
    'Topik / Materi Pokok',
    'Bentuk Soal',
    'Level Kognitif',
    'Kunci Jawaban',
    'Bobot Skor',
    'Jumlah Benar',
    'Benar Sebagian',
    'Jumlah Salah',
    'Total Siswa',
    'Ketercapaian (%)',
    'Kesukaran Empiris',
    'Daya Pembeda (D)',
    'Interpretasi Daya Pembeda',
  ];

  const rows: CellValue[][] = itemStats.map((s, idx) => {
    const q = questions[idx];
    return [
      s.questionNumber,
      s.elemen,
      s.topic,
      formatQuestionTypeLabel(q?.type ?? s.type),
      s.levelKognitif,
      s.kunciJawaban,
      Number(s.bobot ?? 1),
      Number(s.correctCount),
      Number(s.partialCount),
      Number(s.wrongCount),
      Number(s.totalStudents),
      Number(s.persentaseKetercapaian),
      s.klasifikasiKesukaran,
      Number(s.dayaPembeda),
      s.dayaPembedaLabel,
    ];
  });

  const aoa = [headers, ...rows];
  return createStandardSheet(aoa, {
    enableAutoFilter: true,
    minWidth: 10,
    maxWidth: 40,
  });
}

/**
 * Membangun sheet "Rekap Nilai & Integritas" (Tabel Standar Header Baris 1)
 */
function buildIntegritySheet(results: ExamResult[]): XLSX.WorkSheet {
  const headers = [
    'No',
    'Nomor Absen',
    'Nama Siswa',
    'Kelas',
    'Mata Pelajaran',
    'Benar',
    'Salah',
    'Nilai Akhir',
    'Status KKTP',
    'Durasi Pengerjaan',
    'Jumlah Pelanggaran',
    'Status Pengerjaan',
    'Catatan Integritas',
    'Riwayat Waktu Pelanggaran',
    'Waktu Submit',
  ];

  const rows: CellValue[][] = results.map((r, idx) => {
    const isLulus = r.status === 'Lulus' || (r.nilai ?? 0) >= CONFIG.KKTP;
    const { statusSubmit, catatan, logHistory } = getIntegrityFields(r);

    return [
      idx + 1,
      r.noAbsen || '-',
      r.nama || '-',
      r.kelas || CONFIG.KELAS,
      r.mataPelajaran || CONFIG.MATA_PELAJARAN,
      Number(r.benar ?? 0),
      Number(r.salah ?? 0),
      Number(r.nilai ?? 0),
      isLulus ? 'Tuntas' : 'Remedial',
      r.durasiPengerjaan || '-',
      Number(r.violationsCount ?? 0),
      statusSubmit,
      catatan,
      logHistory,
      r.timestamp || '-',
    ];
  });

  const aoa = [headers, ...rows];
  return createStandardSheet(aoa, {
    enableAutoFilter: true,
    minWidth: 10,
    maxWidth: 48,
  });
}

/**
 * 1. Ekspor Rekap Nilai Siswa ke File Microsoft Excel Asli (.xlsx) Multi-Sheet
 * Sheet 1: Rekap Nilai | Sheet 2: Ringkasan Statistik | Sheet 3: Detail Jawaban & Butir
 */
export function exportResultsToXLSX(
  results: ExamResult[],
  questions: Question[]
): void {
  const wb = XLSX.utils.book_new();
  const { itemStats, allResponses } = computeItemAnalysis(results, questions);

  const wsRekap = buildRekapNilaiSheet(results);
  const wsStats = buildRingkasanStatistikSheet(results);
  const wsDetail = buildItemMatrixSheet(results, questions, itemStats, allResponses);

  XLSX.utils.book_append_sheet(wb, wsRekap, sanitizeSheetName('Rekap Nilai'));
  XLSX.utils.book_append_sheet(wb, wsStats, sanitizeSheetName('Ringkasan Statistik'));
  XLSX.utils.book_append_sheet(wb, wsDetail, sanitizeSheetName('Detail Jawaban & Butir'));

  const today = new Date().toISOString().slice(0, 10);
  const safeKelas = CONFIG.KELAS.replace(/[^a-zA-Z0-9_-]/g, '_');
  XLSX.writeFile(wb, `Rekap_Nilai_Tes_Sumatif_Kelas_${safeKelas}_${today}.xlsx`);
}

/**
 * 2. Ekspor Matriks Analisis Butir Soal ke File Microsoft Excel Asli (.xlsx) Multi-Sheet
 * Sheet 1: Matriks Jawaban | Sheet 2: Statistik Butir Soal | Sheet 3: Rekap Nilai & Integritas
 */
export function exportItemAnalysisMatrixToXLSX(
  results: ExamResult[],
  questions: Question[],
  precomputedStats?: QuestionItemAnalysis[],
  precomputedResponses?: Map<string, Record<number, StudentItemResponse>>
): void {
  const wb = XLSX.utils.book_new();

  let itemStats = precomputedStats;
  let allResponses = precomputedResponses;

  if (!itemStats || !allResponses) {
    const computed = computeItemAnalysis(results, questions);
    itemStats = computed.itemStats;
    allResponses = computed.allResponses;
  }

  const wsMatrix = buildItemMatrixSheet(results, questions, itemStats, allResponses);
  const wsButir = buildStatistikButirSheet(questions, itemStats);
  const wsIntegrity = buildIntegritySheet(results);

  XLSX.utils.book_append_sheet(wb, wsMatrix, sanitizeSheetName('Matriks Jawaban'));
  XLSX.utils.book_append_sheet(wb, wsButir, sanitizeSheetName('Statistik Butir Soal'));
  XLSX.utils.book_append_sheet(wb, wsIntegrity, sanitizeSheetName('Rekap Nilai & Integritas'));

  const today = new Date().toISOString().slice(0, 10);
  const safeKelas = CONFIG.KELAS.replace(/[^a-zA-Z0-9_-]/g, '_');
  XLSX.writeFile(wb, `Matriks_Analisis_Butir_Soal_Kelas_${safeKelas}_${today}.xlsx`);
}
