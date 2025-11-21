<?php
require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelImportKAP
{
    private $db;
    private $filename = 'format_kap.xlsx';

    // Konfigurasi API (static)
    private $endpoint = 'http://192.168.10.36:8090/api';
    private $apiToken = '797e9aa1-be97-4dc0-ae13-3ecd304a61a3';

    public function __construct()
    {
        $this->connectDB();
    }

    private function connectDB()
    {
        $host = 'localhost';
        $dbname = 'bpkp_upload';
        $username = 'root';
        $password = '';

        try {
            $this->db = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", $username, $password);
            $this->db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            echo "<div style='color: blue; padding: 10px; border: 1px solid blue; border-radius: 5px;'>";
            echo "<strong>INFO:</strong> Koneksi database berhasil";
            echo "</div>";
        } catch (PDOException $e) {
            die("<div style='color: red; padding: 10px; border: 1px solid red; border-radius: 5px;'><strong>ERROR Koneksi database:</strong> " . $e->getMessage() . "</div>");
        }
    }

    public function importData()
    {
        if (!file_exists($this->filename)) {
            return [
                'success' => false,
                'message' => "File {$this->filename} tidak ditemukan dalam folder yang sama"
            ];
        }

        try {
            // Baca file Excel
            $spreadsheet = IOFactory::load($this->filename);
            $worksheet = $spreadsheet->getActiveSheet();

            // Get highest row number
            $highestRow = $worksheet->getHighestRow();

            $importedCount = 0;
            $errors = [];

            echo "<div style='color: blue; padding: 10px; border: 1px solid blue; border-radius: 5px;'>";
            echo "<strong>INFO:</strong> File Excel ditemukan. Total baris: $highestRow";
            echo "</div>";

            // Get maxKaldikID dari API sekali saja di awal
            $maxKaldikIdFromApi = $this->getMaxKaldikIdFromApi();

            // Mulai dari baris 3 (baris 1 dan 2 di-skip karena merged/header)
            for ($row = 3; $row <= $highestRow; $row++) {
                // Ambil data dari kolom A sampai H
                $topik_nama = $worksheet->getCell('A' . $row)->getCalculatedValue();
                $keterangan_program_pembelajaran = $worksheet->getCell('B' . $row)->getCalculatedValue();
                $arahan_pimpinan = $worksheet->getCell('C' . $row)->getCalculatedValue();
                $prioritas_pembelajaran = $worksheet->getCell('D' . $row)->getCalculatedValue();
                $tujuan_program_pembelajaran = $worksheet->getCell('E' . $row)->getCalculatedValue();
                $kompetensi_dasar = $worksheet->getCell('F' . $row)->getCalculatedValue();
                $indikator_keberhasilan = $worksheet->getCell('G' . $row)->getCalculatedValue();
                $indikator_dampak_terhadap_kinerja_organisasi = $worksheet->getCell('H' . $row)->getCalculatedValue();
                $penugasan_yang_terkait_dengan_pembelajaran = $worksheet->getCell('I' . $row)->getCalculatedValue();
                $skill_group_owner = $worksheet->getCell('J' . $row)->getCalculatedValue();
                $metodeName = $worksheet->getCell('K' . $row)->getCalculatedValue();
                $diklatLocName = $worksheet->getCell('P' . $row)->getCalculatedValue();
                $detail_lokasi = $worksheet->getCell('Q' . $row)->getCalculatedValue();
                $diklatTypeName = $worksheet->getCell('R' . $row)->getCalculatedValue();
                $sasaran_peserta = $worksheet->getCell('S' . $row)->getCalculatedValue();
                $kriteria_peserta = $worksheet->getCell('T' . $row)->getCalculatedValue();
                $aktivitas_prapembelajaran = $worksheet->getCell('U' . $row)->getCalculatedValue();


                // Debug: Tampilkan data yang dibaca
                echo "<div style='background: #f0f0f0; padding: 10px; margin: 5px 0; border-radius: 5px;'>";
                echo "<strong>Baris $row:</strong> ";
                echo "Topik Nama: '$topik_nama', ";
                echo "Keterangan: '$keterangan_program_pembelajaran', ";
                echo "Tujuan: '$tujuan_program_pembelajaran'";
                echo "</div>";

                // Skip baris kosong
                if (empty(trim($topik_nama)) && empty(trim($keterangan_program_pembelajaran))) {
                    echo "<div style='color: gray; padding: 5px;'>Baris $row: Dikosongkan karena data kosong</div>";
                    continue;
                }

                // Process data row
                $result = $this->processRow([
                    'topik_nama' => $topik_nama,
                    'keterangan_program_pembelajaran' => $keterangan_program_pembelajaran,
                    'arahan_pimpinan' => $arahan_pimpinan,
                    'prioritas_pembelajaran' => $prioritas_pembelajaran,
                    'tujuan_program_pembelajaran' => $tujuan_program_pembelajaran,
                    'kompetensi_dasar' => $kompetensi_dasar,
                    'indikator_keberhasilan' => $indikator_keberhasilan,
                    'indikator_dampak_terhadap_kinerja_organisasi' => $indikator_dampak_terhadap_kinerja_organisasi,
                    'penugasan_yang_terkait_dengan_pembelajaran' => $penugasan_yang_terkait_dengan_pembelajaran,
                    'skill_group_owner' => $skill_group_owner,
                    'metodeName' => $metodeName,
                    'detail_lokasi' => $detail_lokasi,
                    'diklatTypeName' => $diklatTypeName,
                    'diklatLocName' => $diklatLocName,
                    'sasaran_peserta' => $sasaran_peserta,
                    'kriteria_peserta' => $kriteria_peserta,
                    'aktivitas_prapembelajaran' => $aktivitas_prapembelajaran
                ], $row, $maxKaldikIdFromApi);

                if ($result['success']) {
                    $importedCount++;
                    echo "<div style='color: green; padding: 5px;'>Baris $row: BERHASIL diimport</div>";
                } else {
                    $errors[] = "Baris $row: " . $result['message'];
                    echo "<div style='color: red; padding: 5px;'>Baris $row: GAGAL - " . $result['message'] . "</div>";
                }
            }

            return [
                'success' => true,
                'imported' => $importedCount,
                'total_rows' => $highestRow - 2,
                'errors' => $errors
            ];
        } catch (Exception $e) {
            return [
                'success' => false,
                'message' => "Error membaca file Excel: " . $e->getMessage()
            ];
        }
    }

    private function processRow($data, $rowNumber, $maxKaldikIdFromApi)
    {
        try {
            // Validasi data required
            if (empty($data['topik_nama'])) {
                return [
                    'success' => false,
                    'message' => 'Nama Topik Program Pembelajaran harus diisi'
                ];
            }

            // Cari topik_id berdasarkan nama topik
            $topik_id = $this->getTopikIdByName($data['topik_nama']);
            if (!$topik_id) {
                return [
                    'success' => false,
                    'message' => 'Topik dengan nama "' . $data['topik_nama'] . '" tidak ditemukan di database'
                ];
            }

            // Generate kode_pembelajaran otomatis
            $kode_pembelajaran = $this->generateKodePembelajaran($topik_id, $maxKaldikIdFromApi);

            // Data static sesuai permintaan
            $currentDateTime = date('Y-m-d H:i:s');
            $currentTimestamp = date('Y-m-d H:i:s');
            $judul = $data['topik_nama'];
            if (!empty($data['keterangan_program_pembelajaran']) && trim($data['keterangan_program_pembelajaran']) !== '') {
                $judul .= ' ' . trim($data['keterangan_program_pembelajaran']);
            }

            $metodeID = null;
            if (!empty($data['metodeName'])) {
                switch ($data['metodeName']) {
                    case 'Full Tatap Muka':
                        $metodeID = '1';
                        break;
                    case 'Blended Learning':
                        $metodeID = '2';
                        break;
                    case 'Full E-Learning':
                        $metodeID = '4';
                        break;
                    default:
                        $metodeID = null;
                }
            }

            $diklatTypeID = null;
            if (!empty($data['diklatTypeName'])) {
                switch ($data['diklatTypeName']) {
                    case 'Fungsional Auditor':
                        $diklatTypeID = '1';
                        break;
                    case 'TS Manajemen Pengawasan':
                        $diklatTypeID = '2';
                        break;
                    case 'Kedinasan':
                        $diklatTypeID = '3';
                        break;
                    case 'TS Pengawasan':
                        $diklatTypeID = '4';
                        break;
                    case 'Manajerial':
                        $diklatTypeID = '5';
                        break;
                    case 'Sertifikasi Non JFA':
                        $diklatTypeID = '6';
                        break;
                    case 'Micro Learning':
                        $diklatTypeID = '8';
                        break;
                    case 'MOOC':
                        $diklatTypeID = '7';
                        break;
                    default:
                        $diklatTypeID = null;
                }
            }

            $diklatLocID = null;
            if (!empty($data['diklatLocName'])) {
                $locations = [
                    'Denpasar' => '1',
                    'Karangasem' => '2',
                    'Pangkalpinang' => '3',
                    'Tanjung Pinang' => '4',
                    'Serang' => '5',
                    'Cilegon' => '6',
                    'Bengkulu' => '7',
                    'Jakarta' => '8',
                    'Cibubur' => '9',
                    'Gorontalo' => '10',
                    'Jambi' => '11',
                    'Pusdiklatwas BPKP' => '12',
                    'Bandung' => '13',
                    'Kota Bogor' => '14',
                    'Tangerang' => '15',
                    'Depok' => '16',
                    'Cianjur' => '17',
                    'Sukabumi' => '18',
                    'Tangerang Selatan' => '19',
                    'Indramayu' => '20',
                    'Tasikmalaya' => '21',
                    'Semarang' => '22',
                    'Baturaden' => '23',
                    'Banyumas' => '24',
                    'Pemalang' => '25',
                    'Surabaya' => '26',
                    'Bojonegoro' => '27',
                    'Pontianak' => '28',
                    'Ketapang' => '29',
                    'Banjarmasin' => '30',
                    'Banjarbaru' => '31',
                    'Kabupaten Tabalong' => '32',
                    'Tanjung' => '33',
                    'Palangkaraya' => '34',
                    'Samarinda' => '35',
                    'Lampung' => '36',
                    'Bandar Lampung' => '37',
                    'Ambon' => '38',
                    'Ternate' => '39',
                    'Banda Aceh' => '40',
                    'Mataram' => '41',
                    'Kupang' => '42',
                    'Ende' => '43',
                    'Jayapura' => '44',
                    'Asmat' => '45',
                    'Papua' => '46',
                    'Pekanbaru' => '47',
                    'Batam' => '48',
                    'Makassar' => '49',
                    'Palu' => '50',
                    'Kendari' => '51',
                    'Manado' => '52',
                    'Padang' => '53',
                    'Pariaman' => '54',
                    'Pasaman' => '55',
                    'Palembang' => '56',
                    'Bangka Belitung' => '57',
                    'Medan' => '58',
                    'Yogyakarta' => '59',
                    'Musi Rawas' => '60',
                    'Papua Barat' => '61',
                    'Manokwari' => '62',
                    'Subang' => '63',
                    'Mamuju' => '64',
                    'Kabupaten Bogor' => '65',
                    'Tarakan' => '66',
                    'Cirebon' => '67',
                    'Ciputat' => '68',
                    'Kuningan' => '69',
                    'Majalengka' => '70',
                    'Tanjung Selor' => '71',
                    'Balikpapan' => '72',
                    'Ciamis' => '73',
                    'Boyolali' => '74',
                    'Langsa' => '75',
                    'Badung' => '76',
                    'Sorong' => '77',
                    'Malang' => '78',
                    'Pembelajaran Jarak Jauh' => '79',
                    'Sidoarjo' => '80',
                    'Parapat' => '81',
                    'Bekasi' => '82',
                    'Magelang' => '83',
                    'Hybrid' => '84',
                    'Purwakarta' => '85',
                    'Seluruh Perwakilan BPKP' => '86',
                    'Buleleng' => '87',
                    'Trenggalek' => '88',
                    'Soreang' => '89',
                    'Tempat Kedudukan Masing-Masing Peserta' => '200',
                    'Trenggalek' => '201',
                    'Soreang' => '202'
                ];

                $diklatLocID = $locations[$data['diklatLocName']] ?? null;
            }




            // Prepare data untuk insert
            $insertData = [
                'kode_pembelajaran' => $kode_pembelajaran,
                'institusi_sumber' => 'Non BPKP',
                'jenis_program' => 'APIP',
                'frekuensi_pelaksanaan' => 'Tahunan',
                'referensi_indikator_kinerja' => null,
                'topik_id' => $topik_id,
                'keterangan_program_pembelajaran' => $data['keterangan_program_pembelajaran'],
                'judul' => $judul,
                'arahan_pimpinan' => $data['arahan_pimpinan'],
                'tahun' => '2026',
                'prioritas_pembelajaran' => 'Prioritas ' . $data['prioritas_pembelajaran'],
                'tujuan_program_pembelajaran' => $data['tujuan_program_pembelajaran'],
                'indikator_dampak_terhadap_kinerja_organisasi' => $data['indikator_dampak_terhadap_kinerja_organisasi'],
                'penugasan_yang_terkait_dengan_pembelajaran' => $data['penugasan_yang_terkait_dengan_pembelajaran'],
                'skill_group_owner' => $data['skill_group_owner'],
                'diklatLocID' => $diklatLocID,
                'diklatLocName' => $data['diklatLocName'],
                'detail_lokasi' => $data['detail_lokasi'],
                'kelas' => 1,
                'diklatTypeID' => $diklatTypeID,
                'diklatTypeName' => $data['diklatTypeName'],
                'metodeID' => $metodeID,
                'metodeName' => $data['metodeName'],
                'biayaID' => '3',
                'biayaName' => 'PNBP',
                'latsar_stat' => '0',
                'bentuk_pembelajaran' => 'Klasikal',
                'jalur_pembelajaran' => 'Pelatihan',
                'model_pembelajaran' => 'Pembelajaran terstruktur',
                'peserta_pembelajaran' => 'Eksternal',
                'sasaran_peserta' => $data['sasaran_peserta'],
                'kriteria_peserta' => $data['kriteria_peserta'],
                'aktivitas_prapembelajaran' => $data['aktivitas_prapembelajaran'],
                'penyelenggara_pembelajaran' => 'Pusdiklatwas BPKP',
                'fasilitator_pembelajaran' => json_encode(['Widyaiswara', 'Instruktur']),
                'sertifikat' => 'Sertifikat mengikuti pembelajaran',
                'tanggal_created' => $currentDateTime,
                'status_pengajuan' => 'Pending',
                'status_sync' => 'Waiting',
                'kompetensi_dasar' => $data['kompetensi_dasar'],
                'user_created' => null,
                'unit_kerja_id' => null,
                'nama_unit' => null,
                'current_step' => 1,
                'surat_permintaan' => null,
                'ruang_kelas_ceklist' => null,
                'ruang_kelas_keterangan' => null,
                'kurikulum_ceklist' => null,
                'kurikulum_keterangan' => null,
                'fasilitator_ceklist' => null,
                'fasilitator_keterangan' => null,
                'created_at' => $currentTimestamp,
                'updated_at' => $currentTimestamp
            ];

            // Debug: Tampilkan kode pembelajaran yang di-generate
            echo "<div style='color: purple; padding: 5px;'>Topik ID: $topik_id, Kode Pembelajaran: $kode_pembelajaran</div>";

            // Insert ke tabel pengajuan_kap
            $success = $this->insertPengajuanKAP($insertData);

            if ($success) {
                return ['success' => true, 'message' => 'Data berhasil diimport'];
            } else {
                return ['success' => false, 'message' => 'Gagal menyimpan data ke database'];
            }
        } catch (Exception $e) {
            return [
                'success' => false,
                'message' => 'Error processing row: ' . $e->getMessage()
            ];
        }
    }

    private function getTopikIdByName($nama_topik)
    {
        try {
            $query = $this->db->prepare("SELECT id FROM topik WHERE nama_topik = ? LIMIT 1");
            $query->execute([$nama_topik]);
            $result = $query->fetch(PDO::FETCH_ASSOC);

            if ($result) {
                return $result['id'];
            }

            // Jika tidak ditemukan, coba cari dengan LIKE (partial match)
            $query = $this->db->prepare("SELECT id FROM topik WHERE nama_topik LIKE ? LIMIT 1");
            $query->execute(['%' . $nama_topik . '%']);
            $result = $query->fetch(PDO::FETCH_ASSOC);

            return $result ? $result['id'] : null;
        } catch (Exception $e) {
            echo "<div style='color: red; padding: 5px;'>Error mencari topik: " . $e->getMessage() . "</div>";
            return null;
        }
    }

    private function getMaxKaldikIdFromApi()
    {
        try {
            $tahun = '2026';
            $prefixTahun = substr($tahun, -2); // 26

            $url = $this->endpoint . '/len-kaldik/max-id';
            $params = [
                'api_key' => $this->apiToken,
                'tahun' => $prefixTahun
            ];

            $fullUrl = $url . '?' . http_build_query($params);

            echo "<div style='color: blue; padding: 10px; border: 1px solid blue; border-radius: 5px;'>";
            echo "<strong>INFO API:</strong> Mengambil maxKaldikID dari: " . $fullUrl;
            echo "</div>";

            // Gunakan cURL untuk hit API
            $ch = curl_init();
            curl_setopt($ch, CURLOPT_URL, $fullUrl);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($ch, CURLOPT_TIMEOUT, 30);

            $response = curl_exec($ch);
            $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
            $error = curl_error($ch);
            curl_close($ch);

            if ($httpCode === 200 && $response) {
                $data = json_decode($response, true);

                if (isset($data['maxKaldikID'])) {
                    echo "<div style='color: green; padding: 10px; border: 1px solid green; border-radius: 5px;'>";
                    echo "<strong>SUKSES API:</strong> maxKaldikID = " . $data['maxKaldikID'];
                    echo "</div>";
                    return $data['maxKaldikID'];
                } else {
                    echo "<div style='color: orange; padding: 10px; border: 1px solid orange; border-radius: 5px;'>";
                    echo "<strong>WARNING API:</strong> maxKaldikID tidak ditemukan dalam response";
                    echo "</div>";
                }
            } else {
                echo "<div style='color: orange; padding: 10px; border: 1px solid orange; border-radius: 5px;'>";
                echo "<strong>WARNING API:</strong> Gagal hit API. HTTP Code: $httpCode, Error: $error";
                echo "</div>";
            }
        } catch (Exception $e) {
            echo "<div style='color: orange; padding: 10px; border: 1px solid orange; border-radius: 5px;'>";
            echo "<strong>WARNING API:</strong> Exception: " . $e->getMessage();
            echo "</div>";
        }

        return null;
    }

    private function generateKodePembelajaran($topik_id, $maxKaldikIdFromApi)
    {
        try {
            $tahun = '2026'; // Static sesuai permintaan
            $prefixTahun = substr($tahun, -2); // Ambil 2 digit terakhir => 26
            $topikIdFormatted = sprintf('%03d', $topik_id); // Format 3 digit

            // 1. Get last data dari DB berdasarkan tahun 2026
            $query = $this->db->prepare("SELECT kode_pembelajaran FROM pengajuan_kap WHERE tahun = ? ORDER BY id DESC LIMIT 1");
            $query->execute([$tahun]);
            $lastPengajuanFromDb = $query->fetch(PDO::FETCH_ASSOC);

            $kodePembelajaranDb = $lastPengajuanFromDb ? $lastPengajuanFromDb['kode_pembelajaran'] : null;

            // 2. Ambil 4 digit terakhir dari DB dan API
            $lastNoFromDb = $kodePembelajaranDb ? (int)substr($kodePembelajaranDb, -4) : 0;
            $lastNoFromApi = $maxKaldikIdFromApi ? (int)substr($maxKaldikIdFromApi, -4) : 0;

            // Jika tidak ada data di DB dan API, mulai dari 1
            if ($lastNoFromDb == 0 && $lastNoFromApi == 0) {
                $nextNo = 1;
            } else {
                $nextNo = max($lastNoFromApi, $lastNoFromDb) + 1;
            }

            $nextNoFormatted = sprintf('%04d', $nextNo);

            // 3. Generate final code with topic ID
            $prefixFinal = $prefixTahun . $topikIdFormatted;
            $kodePembelajaran = $prefixFinal . $nextNoFormatted;

            echo "<div style='color: blue; padding: 5px;'>";
            echo "Generate Kode: Tahun=$tahun, PrefixTahun=$prefixTahun, TopikID=$topikIdFormatted, ";
            echo "LastNoFromApi=$lastNoFromApi, LastNoFromDb=$lastNoFromDb, NextNo=$nextNoFormatted";
            echo "</div>";

            return $kodePembelajaran;
        } catch (Exception $e) {
            // Fallback jika query gagal
            echo "<div style='color: orange; padding: 5px;'>Fallback kode pembelajaran: " . $e->getMessage() . "</div>";
            $prefixTahun = '26';
            $topikIdFormatted = sprintf('%03d', $topik_id);
            return $prefixTahun . $topikIdFormatted . '0001';
        }
    }

    private function insertPengajuanKAP($data)
    {
        $fields = implode(', ', array_keys($data));
        $placeholders = ':' . implode(', :', array_keys($data));

        $sql = "INSERT INTO pengajuan_kap ($fields) VALUES ($placeholders)";

        echo "<div style='background: #fffacd; padding: 10px; margin: 5px 0; border: 1px solid #ffd700; border-radius: 5px;'>";
        echo "<strong>SQL Debug:</strong> " . $sql;
        echo "</div>";

        try {
            $stmt = $this->db->prepare($sql);
            $result = $stmt->execute($data);

            if ($result) {
                echo "<div style='color: green; padding: 5px;'>Insert BERHASIL</div>";
                return true;
            } else {
                $errorInfo = $stmt->errorInfo();
                echo "<div style='color: red; padding: 5px;'>Insert GAGAL: " . $errorInfo[2] . "</div>";
                return false;
            }
        } catch (PDOException $e) {
            echo "<div style='color: red; padding: 5px;'>PDO Exception: " . $e->getMessage() . "</div>";
            return false;
        }
    }
}

// Eksekusi import
echo "<div style='padding: 20px; font-family: Arial, sans-serif;'>";
echo "<h2>Hasil Import Data KAP</h2>";

$importer = new ExcelImportKAP();
$result = $importer->importData();

echo "<hr style='margin: 20px 0;'>";
echo "<h3>Ringkasan Akhir:</h3>";

if ($result['success']) {
    echo "<div style='color: green; padding: 10px; border: 1px solid green; border-radius: 5px;'>";
    echo "<strong>SUKSES:</strong> " . $result['imported'] . " dari " . $result['total_rows'] . " baris berhasil diimport";
    echo "</div>";

    if (!empty($result['errors'])) {
        echo "<div style='color: orange; margin-top: 10px;'>";
        echo "<strong>Peringatan:</strong>";
        echo "<ul>";
        foreach ($result['errors'] as $error) {
            echo "<li>$error</li>";
        }
        echo "</ul>";
        echo "</div>";
    }
} else {
    echo "<div style='color: red; padding: 10px; border: 1px solid red; border-radius: 5px;'>";
    echo "<strong>ERROR:</strong> " . $result['message'];
    echo "</div>";
}

echo "</div>";
