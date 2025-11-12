# deploy.ps1 - PLN TUL I-11 Automation Tool
# Created by vatanzjr

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PLN TUL I-11 Automation Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Function untuk menampilkan pesan error
function Show-Error {
    param([string]$message)
    Write-Host "❌ $message" -ForegroundColor Red
}

# Function untuk menampilkan pesan sukses
function Show-Success {
    param([string]$message)
    Write-Host "✅ $message" -ForegroundColor Green
}

# Function untuk menampilkan informasi
function Show-Info {
    param([string]$message)
    Write-Host "📢 $message" -ForegroundColor Yellow
}

# Function untuk install dependencies dengan retry
function Install-Dependencies {
    $maxRetries = 2
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            Show-Info "Menginstall dependencies npm (Attempt $($retryCount + 1))..."
            
            # Jalankan npm install dengan timeout
            $process = Start-Process -FilePath "npm" -ArgumentList "install" -Wait -PassThru -NoNewWindow
            
            if ($process.ExitCode -eq 0) {
                Show-Success "Dependencies berhasil diinstall"
                return $true
            } else {
                Show-Error "npm install gagal dengan exit code: $($process.ExitCode)"
                $retryCount++
                
                if ($retryCount -lt $maxRetries) {
                    Show-Info "Retry dalam 5 detik..."
                    Start-Sleep -Seconds 5
                }
            }
        } catch {
            Show-Error "Error selama npm install: $($_.Exception.Message)"
            $retryCount++
            
            if ($retryCount -lt $maxRetries) {
                Show-Info "Retry dalam 5 detik..."
                Start-Sleep -Seconds 5
            }
        }
    }
    
    Show-Error "Gagal install dependencies setelah $maxRetries attempts"
    Show-Info "Coba jalankan manual: npm install"
    return $false
}

# Main installation process
try {
    Show-Info "Memulai instalasi PLN TUL I-11 Automation Tool..."
    
    # 1. Check Node.js installation
    Show-Info "Checking Node.js installation..."
    $nodeCheck = Get-Command node -ErrorAction SilentlyContinue
    if (-not $nodeCheck) {
        Show-Error "Node.js tidak terinstall!"
        Show-Info "Silakan download Node.js dari: https://nodejs.org"
        Show-Info "Atau jalankan: winget install OpenJS.NodeJS"
        Write-Host ""
        Show-Info "Setelah Node.js terinstall, jalankan script ini lagi."
        exit 1
    }
    
    # Check Node.js version
    $nodeVersion = node --version
    Show-Success "Node.js version: $nodeVersion"
    
    # 2. Check npm
    Show-Info "Checking npm..."
    $npmCheck = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npmCheck) {
        Show-Error "NPM tidak tersedia!"
        exit 1
    }
    
    $npmVersion = npm --version
    Show-Success "npm version: $npmVersion"
    
    # 3. Create project directory
    $projectPath = Join-Path $env:USERPROFILE "PLN-TULI11-Tool"
    Show-Info "Membuat direktori project: $projectPath"
    
    if (Test-Path $projectPath) {
        Show-Info "Direktori sudah ada, menghapus yang lama..."
        Remove-Item $projectPath -Recurse -Force
    }
    
    New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
    Set-Location $projectPath
    Show-Success "Direktori berhasil dibuat"
    
    # 4. Create package.json
    Show-Info "Membuat package.json..."
    $packageJsonContent = @'
{
  "name": "pln-tuli11-tool",
  "version": "1.0.0",
  "description": "PLN TUL I-11 Automation Tool",
  "main": "tulI11.js",
  "scripts": {
    "start": "node tulI11.js",
    "tuli11": "node tulI11.js"
  },
  "dependencies": {
    "puppeteer": "^21.5.2",
    "axios": "^1.5.0",
    "qs": "^6.11.2"
  },
  "keywords": ["pln", "automation", "tul-i11", "puppeteer"],
  "author": "vatanzjr",
  "license": "MIT"
}
'@
    $packageJsonContent | Out-File -FilePath "package.json" -Encoding UTF8
    Show-Success "package.json created"
    
    # 5. Create tulI11.js dengan hardcoded values
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
    await page.waitForTimeout(3000);

    // Pilih filter "No Agenda"
    await page.waitForSelector('#ext-gen24', { timeout: 10000 });
    await page.type('#ext-gen24', 'No Agenda');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    console.log('📝 Filter diset ke "No Agenda"');
    await page.waitForTimeout(1000);

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
        await page.waitForTimeout(5000);

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
        await page.waitForTimeout(2000);

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
    $tulI11Content | Out-File -FilePath "tulI11.js" -Encoding UTF8
    Show-Success "tulI11.js created"
    
    # 6. Create batch file untuk memudahkan user
    Show-Info "Membatu file batch untuk kemudahan penggunaan..."
    $batchContent = @'
@echo off
chcp 65001 >nul
color 0B
title PLN TUL I-11 Tool

echo ========================================
echo   PLN TUL I-11 Automation Tool
echo ========================================
echo.
echo [1] Jalankan TUL I-11 Tool
echo [2] Buka Folder Project
echo [3] Keluar
echo.
set /p choice="Pilih opsi (1/3): "

if "%choice%"=="1" (
  echo Memulai TUL I-11 Tool...
  node tulI11.js
  pause
) else if "%choice%"=="2" (
  explorer .
  exit
) else if "%choice%"=="3" (
  exit
) else (
  echo Pilihan tidak valid!
  timeout /t 2 >nul
  %0
)
'@
    $batchContent | Out-File -FilePath "start-tuli11.bat" -Encoding ASCII
    Show-Success "start-tuli11.bat created"
    
    # 7. Install npm dependencies DENGAN RETRY
    $dependenciesInstalled = Install-Dependencies
    
    if (-not $dependenciesInstalled) {
        Show-Error "Instalasi dependencies gagal, tetapi tool tetap bisa dicoba."
        Show-Info "Coba jalankan manual: npm install"
    }
    
    # 8. Verification
    Show-Info "Memverifikasi instalasi..."
    
    $requiredFiles = @("package.json", "tulI11.js", "start-tuli11.bat")
    $allFilesExist = $true
    
    foreach ($file in $requiredFiles) {
        if (Test-Path $file) {
            Show-Success "$file - OK"
        } else {
            Show-Error "$file - MISSING"
            $allFilesExist = $false
        }
    }
    
    if (-not $allFilesExist) {
        Show-Error "Beberapa file gagal dibuat"
        exit 1
    }
    
    # 9. Completed - SELALU TAMPIL MESKI DEPENDENCIES GAGAL
    Write-Host ""
    Show-Success "🎉 INSTALASI BERHASIL!"
    Show-Info "Project location: $projectPath"
    Write-Host ""
    Show-Info "File yang dibuat:"
    Get-ChildItem | ForEach-Object { 
        Write-Host "   📄 $($_.Name)" -ForegroundColor Gray 
    }
    
    Write-Host ""
    Show-Info "Konfigurasi TUL I-11:"
    Write-Host "   🔑 Username: 9418672ZY" -ForegroundColor White
    Write-Host "   🔒 Password: mblendez" -ForegroundColor White
    
    Write-Host ""
    if ($dependenciesInstalled) {
        Show-Success "✅ Semua dependencies berhasil diinstall"
    } else {
        Show-Error "❌ Dependencies gagal, tetapi tool bisa dicoba"
        Show-Info "Coba jalankan manual: npm install"
    }
    
    Write-Host ""
    Show-Info "🚀 Menjalankan PLN TUL I-11 Tool..."
    
    # Wait a moment before starting
    Start-Sleep -Seconds 3
    
    # Run the batch file
    Start-Process "start-tuli11.bat" -Wait
    
} catch {
    Show-Error "Terjadi error selama instalasi: $($_.Exception.Message)"
    exit 1
}

# Final message
Write-Host ""
Show-Info "CARA MENJALANKAN ULANG:"
Write-Host "1. Buka folder: $projectPath" -ForegroundColor White
Write-Host "2. Jalankan: start-tuli11.bat" -ForegroundColor White
Write-Host "3. Atau jalankan langsung: node tulI11.js" -ForegroundColor White
Write-Host ""
Show-Info "Untuk update, jalankan lagi script ini:"
Write-Host "   irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex" -ForegroundColor White
Write-Host ""
