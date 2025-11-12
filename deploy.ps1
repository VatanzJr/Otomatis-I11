# 5. Create tulI11.js dengan hardcoded values (FIXED VERSION)
Show-Info "Membuat tulI11.js..."
$tulI11Content = @'
const puppeteer = require('puppeteer');
const axios = require('axios');
const path = require('path');
const qs = require('qs');

// === KONFIGURASI HARCODED ===
const APPS_SCRIPT_PKBA_URL = 'https://script.google.com/macros/s/AKfycbzuJYpjCG7YvDg40ImUMbH_vY1DIm7JBnhCP3APSddesyT6xB0pF5i6XhusEOWUcuTL1Q/exec';
const PLN_USERNAME = '9418672ZY';
const PLN_PASSWORD = 'mblendez';

console.log('========================================');
console.log('  PLN TUL I-11 Automation Tool');
console.log('========================================');
console.log('');

// Fungsi ambil data dari Apps Script
async function getAgendaFromPKBASheet() {
  try {
    console.log('📡 Mengambil data dari Google Apps Script...');
    const response = await axios.get(APPS_SCRIPT_PKBA_URL);
    console.log(`✅ Data berhasil diambil: ${response.data.length} records`);
    return response.data;
  } catch (error) {
    console.error('❌ Gagal mengambil data PKBA dari Apps Script:', error.message);
    return [];
  }
}

// Fungsi utama untuk memproses semua No Agenda
async function prosesCetakI11Loop(noAgendaList) {
  console.log(`🔄 Memproses ${noAgendaList.length} No Agenda...`);
  
  const browser = await puppeteer.launch({
    headless: "new",
    defaultViewport: null,
    args: [
      '--start-maximized',
      '--disable-web-security',
      '--disable-features=IsolateOrigins,site-per-process',
      '--disable-site-isolation-trials'
    ],
    userDataDir: path.resolve(__dirname, 'userdata')
  });

  const page = await browser.newPage();

  try {
    const url = `https://ap2t.pln.co.id/BillingTerpusatAP2TNew1-dr/redirect.jsp?user=${PLN_USERNAME}&password=${PLN_PASSWORD}&page=cetakulangI11`;

    console.log('🌐 Membuka halaman PLN...');
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await new Promise(resolve => setTimeout(resolve, 3000));

    // Pilih filter "No Agenda"
    await page.waitForSelector('#ext-gen24', { timeout: 10000 });
    await page.type('#ext-gen24', 'No Agenda');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    console.log('📝 Filter diset ke "No Agenda"');
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Loop semua No Agenda
    for (const noAgenda of noAgendaList) {
      try {
        console.log('\n' + '='.repeat(50));
        console.log(`🔄 Memproses NO AGENDA: ${noAgenda}`);
        console.log('='.repeat(50));

        // Bersihkan input sebelum mengetik baru
        await page.evaluate(() => { 
            const input = document.querySelector('#ext-gen22');
            if (input) input.value = ''; 
        });
        
        await page.type('#ext-gen22', noAgenda);
        console.log(`📝 Input No Agenda: ${noAgenda}`);

        // Klik tombol Search
        await page.click('#ext-gen103');
        console.log('🔍 Tombol Search diklik...');
        await new Promise(resolve => setTimeout(resolve, 5000));

        // Tunggu hasil tabel muncul
        try {
            await page.waitForSelector('.x-grid3-col-0', { timeout: 10000 });
            const noPDL = await page.$eval('.x-grid3-col-0', el => el.textContent.trim());
            console.log('✅ NOPDL ditemukan:', noPDL);

            // Update ke Apps Script
            console.log('📡 Mengupdate data ke Google Sheet...');
            await axios.post(
                APPS_SCRIPT_PKBA_URL,
                qs.stringify({ noAgenda, noPDL }),
                { 
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    timeout: 10000
                }
            )
            .then(res => console.log('✅ Data berhasil diupdate ke sheet'))
            .catch(err => console.error('❌ Error update:', err.message));

        } catch (error) {
            console.log('❌ Data tidak ditemukan untuk No Agenda:', noAgenda);
            // Update dengan status tidak ditemukan
            await axios.post(
                APPS_SCRIPT_PKBA_URL,
                qs.stringify({ noAgenda, noPDL: 'TIDAK DITEMUKAN' }),
                { 
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    timeout: 10000
                }
            )
            .then(res => console.log('✅ Status "TIDAK DITEMUKAN" diupdate ke sheet'))
            .catch(err => console.error('❌ Error update status:', err.message));
        }

        // Delay antar pencarian
        console.log('⏳ Menunggu 2 detik...');
        await new Promise(resolve => setTimeout(resolve, 2000));

      } catch (err) {
        console.error(`❌ Gagal proses No Agenda ${noAgenda}:`, err.message);
        // lanjut ke agenda berikutnya
      }
    }

    await browser.close();
    console.log('\n🎉 SEMUA NO AGENDA SELESAI DIPROSES!');

  } catch (err) {
    console.error('❌ Gagal memproses loop:', err.message);
    await browser.close();
  }
}

// Fungsi utama
async function main() {
  try {
    console.log('🚀 Memulai PLN TUL I-11 Automation Tool');
    console.log('🔑 Login dengan:');
    console.log('   Username:', PLN_USERNAME);
    console.log('   Password:', PLN_PASSWORD.substring(0, 2) + '******');
    console.log('');
    
    const rows = await getAgendaFromPKBASheet();

    if (!rows || rows.length === 0) {
      console.log('✅ Tidak ada data untuk diproses.');
      return;
    }

    // Ambil semua No Agenda dalam array
    const noAgendaList = rows
        .map(row => row['NOMOR AGENDA'] || row['noAgenda'] || row['No Agenda'])
        .filter(agenda => agenda && agenda.trim() !== '');

    if (noAgendaList.length === 0) {
        console.log('❌ Tidak ada Nomor Agenda yang valid.');
        return;
    }

    console.log(`📋 Ditemukan ${noAgendaList.length} Nomor Agenda:`);
    noAgendaList.forEach((agenda, index) => {
        console.log(`   ${index + 1}. ${agenda}`);
    });
    
    console.log('');
    console.log('⏳ Memulai proses dalam 3 detik...');
    await new Promise(resolve => setTimeout(resolve, 3000));

    // Proses semua No Agenda dalam 1 session
    await prosesCetakI11Loop(noAgendaList);
    
  } catch (error) {
    console.error('❌ Error utama:', error.message);
  }
  
  console.log('');
  console.log('👋 Program selesai. Tekan Ctrl+C untuk keluar.');
  
  // Tunggu user menutup
  process.stdin.resume();
  process.on('SIGINT', function() {
    console.log('\n👋 Program dihentikan oleh user');
    process.exit();
  });
}

// Handle errors
process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ Unhandled Rejection at:', promise, 'reason:', reason);
});

// Jalankan program
main();
'@
$tulI11Content | Out-File -FilePath "tulI11.js" -Encoding ASCII
Show-Success "tulI11.js created (Puppeteer v24 compatible)"
