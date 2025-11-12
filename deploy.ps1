# deploy.ps1 - PLN TUL I-11 Automation Tool
# Created by vatanzjr
# Puppeteer v24 Compatible Version

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

# Function untuk cleanup jika ada installasi sebelumnya yang gagal
function Cleanup-PreviousInstall {
    param([string]$projectPath)
    
    Show-Info "Membersihkan installasi sebelumnya..."
    
    # Hapus node_modules jika ada
    $nodeModulesPath = Join-Path $projectPath "node_modules"
    if (Test-Path $nodeModulesPath) {
        try {
            Remove-Item $nodeModulesPath -Recurse -Force -ErrorAction Stop
            Show-Success "node_modules lama berhasil dihapus"
        } catch {
            Show-Error "Tidak bisa hapus node_modules, tapi lanjut terus..."
        }
    }
    
    # Hapus package-lock.json jika ada
    $packageLockPath = Join-Path $projectPath "package-lock.json"
    if (Test-Path $packageLockPath) {
        try {
            Remove-Item $packageLockPath -Force -ErrorAction Stop
            Show-Success "package-lock.json lama berhasil dihapus"
        } catch {
            Show-Error "Tidak bisa hapus package-lock.json, tapi lanjut terus..."
        }
    }
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
        Show-Info "Pilih version LTS (Recommended for most users)"
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
        Show-Info "Pastikan Node.js terinstall dengan benar."
        exit 1
    }
    
    $npmVersion = npm --version
    Show-Success "npm version: $npmVersion"
    
    # 3. Create project directory
    $projectPath = Join-Path $env:USERPROFILE "PLN-TULI11-Tool"
    Show-Info "Membuat direktori project: $projectPath"
    
    if (Test-Path $projectPath) {
        Show-Info "Direktori sudah ada, membersihkan..."
        Cleanup-PreviousInstall -projectPath $projectPath
    } else {
        New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
    }
    
    Set-Location $projectPath
    Show-Success "Direktori berhasil dibuat"
    
    # 4. Create package.json dengan encoding yang benar
    Show-Info "Membuat package.json..."
    $packageJsonContent = @'
{
  "name": "pln-tuli11-tool",
  "version": "1.0.0",
  "description": "PLN TUL I-11 Automation Tool",
  "main": "tulI11.js",
  "scripts": {
    "start": "node tulI11.js",
    "tuli11": "node tulI11.js",
    "test-deps": "node -e \"try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer missing'); } try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios missing'); } try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS missing'); }\""
  },
  "dependencies": {
    "puppeteer": "^24.15.0",
    "axios": "^1.5.0",
    "qs": "^6.11.2"
  },
  "keywords": ["pln", "automation", "tul-i11", "puppeteer"],
  "author": "vatanzjr",
  "license": "MIT"
}
'@
    # Gunakan encoding ASCII tanpa BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText("$projectPath\package.json", $packageJsonContent, $utf8NoBom)
    Show-Success "package.json created dengan encoding yang benar"
    
    # 5. Create tulI11.js dengan Puppeteer v24 compatibility
    Show-Info "Membuat tulI11.js (Puppeteer v24 compatible)..."
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
    
    # 6. Create robust batch file dengan self-diagnose
    Show-Info "Membuat file batch dengan fitur self-diagnose..."
    $batchContent = @'
@echo off
chcp 65001 >nul
color 0B
title PLN TUL I-11 Tool

set "NODE_PATH=%~dp0node_modules"
set "NODE_ENV=production"

:menu
cls
echo ========================================
echo   PLN TUL I-11 Automation Tool
echo ========================================
echo.
echo [1] Jalankan TUL I-11 Tool
echo [2] Install Dependencies (npm install)
echo [3] Fix Module Path Issue
echo [4] Test Dependencies
echo [5] Buka Folder Project
echo [6] Keluar
echo.
set /p choice="Pilih opsi (1/6): "

if "%choice%"=="1" (
  cls
  echo ========================================
  echo   JALANKAN TUL I-11 TOOL
  echo ========================================
  echo.
  echo Memeriksa dependencies...
  echo.
  node -e "try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer missing'); process.exit(1); }" >nul 2>&1
  if errorlevel 1 (
    echo ❌ ERROR: Module puppeteer tidak ditemukan!
    echo.
    echo 💡 SOLUSI:
    echo   1. Pilih option [2] untuk install dependencies
    echo   2. Atau pilih [3] untuk fix module path
    echo   3. Pastikan sudah run sebagai Administrator
    echo.
    pause
    goto menu
  )
  
  node -e "try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios missing'); process.exit(1); }" >nul 2>&1
  if errorlevel 1 (
    echo ❌ ERROR: Module axios tidak ditemukan!
    echo.
    echo 💡 SOLUSI:
    echo   1. Pilih option [2] untuk install dependencies
    echo   2. Atau pilih [3] untuk fix module path
    echo   3. Pastikan sudah run sebagai Administrator
    echo.
    pause
    goto menu
  )
  
  node -e "try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS missing'); process.exit(1); }" >nul 2>&1
  if errorlevel 1 (
    echo ❌ ERROR: Module qs tidak ditemukan!
    echo.
    echo 💡 SOLUSI:
    echo   1. Pilih option [2] untuk install dependencies
    echo   2. Atau pilih [3] untuk fix module path  
    echo   3. Pastikan sudah run sebagai Administrator
    echo.
    pause
    goto menu
  )
  
  echo ✅ Semua dependencies OK!
  echo.
  echo 🚀 Menjalankan TUL I-11 Tool...
  echo.
  node tulI11.js
  echo.
  pause
  goto menu
) else if "%choice%"=="2" (
  cls
  echo ========================================
  echo   INSTALL DEPENDENCIES
  echo ========================================
  echo.
  echo Menginstall dependencies...
  echo ⏳ Ini mungkin butuh beberapa menit...
  echo.
  call npm install
  echo.
  if %errorlevel% equ 0 (
    echo ✅ Dependencies berhasil diinstall!
    echo.
    echo 💡 Sekarang test dengan option [4]
  ) else (
    echo ❌ Install dependencies gagal.
    echo.
    echo 💡 Coba:
    echo   1. Jalankan sebagai Administrator
    echo   2. Pastikan koneksi internet stabil
    echo   3. Coba option [3] untuk fix
  )
  echo.
  pause
  goto menu
) else if "%choice%"=="3" (
  cls
  echo ========================================
  echo   FIX MODULE PATH ISSUE
  echo ========================================
  echo.
  echo Memperbaiki module path...
  echo.
  echo 1. Membersihkan cache...
  call npm cache clean --force
  echo.
  echo 2. Hapus node_modules lama...
  rmdir /s /q node_modules 2>nul
  echo.
  echo 3. Install ulang dependencies...
  call npm install
  echo.
  echo 4. Verifikasi install...
  node -e "try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer missing'); }"
  node -e "try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios missing'); }"
  node -e "try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS missing'); }"
  echo.
  echo 💡 Sekarang test dengan option [4]
  echo.
  pause
  goto menu
) else if "%choice%"=="4" (
  cls
  echo ========================================
  echo   TEST DEPENDENCIES
  echo ========================================
  echo.
  echo Testing semua dependencies...
  echo.
  node -e "try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer missing'); }"
  node -e "try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios missing'); }"
  node -e "try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS missing'); }"
  echo.
  echo 💡 STATUS:
  echo   - Jika semua ✅ OK: Bisa jalankan option [1]
  echo   - Jika ada ❌ missing: Jalankan option [2] atau [3]
  echo.
  pause
  goto menu
) else if "%choice%"=="5" (
  explorer .
  goto menu
) else if "%choice%"=="6" (
  exit
) else (
  echo Pilihan tidak valid!
  timeout /t 2 >nul
  goto menu
)
'@
    $batchContent | Out-File -FilePath "start-tuli11.bat" -Encoding ASCII
    Show-Success "start-tuli11.bat created dengan fitur self-diagnose"
    
    # 7. Create quick test file
    Show-Info "Membuat file test cepat..."
    $testContent = @'
@echo off
echo Testing Node.js dan dependencies...
node --version
npm --version
echo.
echo Testing modules...
node -e "try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer error:', e.message); }"
node -e "try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios error:', e.message); }" 
node -e "try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS error:', e.message); }"
pause
'@
    $testContent | Out-File -FilePath "test-deps.bat" -Encoding ASCII
    Show-Success "test-deps.bat created"
    
    # 8. Verification
    Show-Info "Memverifikasi instalasi..."
    
    $requiredFiles = @("package.json", "tulI11.js", "start-tuli11.bat", "test-deps.bat")
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
    
    # 9. Completed
    Write-Host ""
    Show-Success "🎉 INSTALASI FILE BERHASIL!"
    Show-Info "Project location: $projectPath"
    Write-Host ""
    
    Show-Info "📋 FILE YANG DIBUAT:"
    Get-ChildItem | ForEach-Object { 
        Write-Host "   📄 $($_.Name)" -ForegroundColor Gray 
    }
    Write-Host ""
    
    Show-Info "🚀 CARA MENGGUNAKAN:"
    Write-Host "   1. Buka folder: $projectPath" -ForegroundColor White
    Write-Host "   2. Jalankan: start-tuli11.bat" -ForegroundColor White
    Write-Host "   3. Pilih [4] Test Dependencies TERLEBIH DAHULU" -ForegroundColor White
    Write-Host "   4. Jika ada yang missing, pilih [2] atau [3]" -ForegroundColor White
    Write-Host "   5. Setelah semua ✅ OK, pilih [1] untuk menjalankan" -ForegroundColor White
    Write-Host ""
    
    Show-Info "💡 TIPS PENTING:"
    Write-Host "   - Jalankan sebagai Administrator jika ada permission error" -ForegroundColor Yellow
    Write-Host "   - Selalu test dependencies dulu sebelum run tool" -ForegroundColor Yellow
    Write-Host "   - Menu punya fitur auto-diagnose dan fix" -ForegroundColor Yellow
    Write-Host "   - Tool sekarang compatible dengan Puppeteer v24+" -ForegroundColor Green
    Write-Host ""
    
    # 10. Tanya user apakah mau test sekarang
    $answer = Read-Host "Jalankan test dependencies sekarang? (y/n)"
    if ($answer -eq 'y' -or $answer -eq 'Y') {
        Show-Info "Menjalankan test dependencies..."
        Start-Sleep -Seconds 2
        
        # Test Node.js dan npm
        Write-Host "=== TEST NODE.JS & NPM ===" -ForegroundColor Cyan
        node --version
        npm --version
        
        # Test jika dependencies sudah ada
        Write-Host "`n=== TEST DEPENDENCIES ===" -ForegroundColor Cyan
        if (Test-Path "node_modules") {
            Write-Host "node_modules ditemukan..." -ForegroundColor Green
            node -e "try { require('puppeteer'); console.log('✅ Puppeteer OK'); } catch(e) { console.log('❌ Puppeteer missing'); }"
            node -e "try { require('axios'); console.log('✅ Axios OK'); } catch(e) { console.log('❌ Axios missing'); }"
            node -e "try { require('qs'); console.log('✅ QS OK'); } catch(e) { console.log('❌ QS missing'); }"
        } else {
            Write-Host "❌ node_modules belum ada" -ForegroundColor Red
            Write-Host "💡 Jalankan 'npm install' dulu" -ForegroundColor Yellow
        }
        
        Write-Host ""
        Show-Info "Tekan Enter untuk membuka menu tool..."
        Pause
        Start-Process "start-tuli11.bat" -Wait
    } else {
        Show-Info "Anda bisa jalankan tool nanti dengan:"
        Write-Host "   $projectPath\start-tuli11.bat" -ForegroundColor White
        Write-Host ""
        Show-Info "Tekan Enter untuk membuka folder..."
        Pause
        Start-Process "explorer" -ArgumentList $projectPath -Wait
    }
    
} catch {
    Show-Error "Terjadi error selama instalasi: $($_.Exception.Message)"
    Write-Host ""
    Show-Info "Silakan coba lagi atau hubungi support."
}

# Final message
Write-Host ""
Show-Info "📞 SUPPORT & UPDATE:"
Write-Host "   GitHub: https://github.com/vatanzjr/otomatis-I11" -ForegroundColor White
Write-Host ""
Show-Info "Untuk install ulang atau update, jalankan:"
Write-Host "   irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex" -ForegroundColor White
Write-Host ""
