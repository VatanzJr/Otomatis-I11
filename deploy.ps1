# deploy.ps1 - PLN TUL I-11 Automation Tool
# Created by vatanzjr
# One-Time Install Version

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

# Function untuk check jika tool sudah terinstall
function Check-ExistingInstallation {
    param([string]$projectPath)
    
    $requiredFiles = @("package.json", "tulI11.js", "start-tuli11.bat")
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path (Join-Path $projectPath $file))) {
            return $false
        }
    }
    return $true
}

# Function untuk check dependencies
function Check-Dependencies {
    param([string]$projectPath)
    
    $nodeModulesPath = Join-Path $projectPath "node_modules"
    if (-not (Test-Path $nodeModulesPath)) {
        return $false
    }
    
    # Check jika puppeteer ada
    $puppeteerPath = Join-Path $nodeModulesPath "puppeteer"
    if (-not (Test-Path $puppeteerPath)) {
        return $false
    }
    
    return $true
}

# Main installation process
try {
    Show-Info "Memulai PLN TUL I-11 Automation Tool..."
    
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
    
    # Check jika tool sudah terinstall
    $isAlreadyInstalled = Check-ExistingInstallation -projectPath $projectPath
    $hasDependencies = Check-Dependencies -projectPath $projectPath
    
    if ($isAlreadyInstalled -and $hasDependencies) {
        Show-Success "✅ Tool sudah terinstall dan siap digunakan!"
        Show-Info "Project location: $projectPath"
        Write-Host ""
        
        Show-Info "🚀 Menjalankan TUL I-11 Tool..."
        Start-Sleep -Seconds 2
        Set-Location $projectPath
        Start-Process "start-tuli11.bat" -Wait
        exit 0
    }
    
    if ($isAlreadyInstalled -and (-not $hasDependencies)) {
        Show-Info "Tool sudah terinstall tapi dependencies belum ada."
        Show-Info "Menginstall dependencies..."
        Set-Location $projectPath
        npm install
        if ($LASTEXITCODE -eq 0) {
            Show-Success "✅ Dependencies berhasil diinstall!"
            Show-Info "🚀 Menjalankan TUL I-11 Tool..."
            Start-Sleep -Seconds 2
            Start-Process "start-tuli11.bat" -Wait
            exit 0
        } else {
            Show-Error "❌ Gagal install dependencies"
            Show-Info "Silakan jalankan manual: npm install"
            exit 1
        }
    }
    
    # Jika belum terinstall, buat fresh install
    Show-Info "Membuat direktori project: $projectPath"
    
    if (Test-Path $projectPath) {
        Show-Info "Direktori sudah ada, membuat instalasi fresh..."
        # Hanya hapus file-file tool, biarkan node_modules jika ada
        $filesToRemove = @("package.json", "tulI11.js", "start-tuli11.bat", "test-deps.bat")
        foreach ($file in $filesToRemove) {
            $filePath = Join-Path $projectPath $file
            if (Test-Path $filePath) {
                Remove-Item $filePath -Force -ErrorAction SilentlyContinue
            }
        }
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
    "tuli11": "node tulI11.js"
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
    Show-Success "package.json created"
    
    # 5. Create tulI11.js dengan Puppeteer v24 compatibility
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
    Show-Success "tulI11.js created"
    
    # 6. Create simple batch file
    Show-Info "Membuat file batch..."
    $batchContent = @'
@echo off
chcp 65001 >nul
color 0B
title PLN TUL I-11 Tool

echo ========================================
echo   PLN TUL I-11 Automation Tool
echo ========================================
echo.
echo Menjalankan TUL I-11 Tool...
echo.
node tulI11.js
echo.
pause
'@
    $batchContent | Out-File -FilePath "start-tuli11.bat" -Encoding ASCII
    Show-Success "start-tuli11.bat created"
    
    # 7. Install dependencies
    Show-Info "Menginstall dependencies..."
    Write-Host "⏳ Ini mungkin butuh beberapa menit..." -ForegroundColor Yellow
    npm install
    
    if ($LASTEXITCODE -eq 0) {
        Show-Success "✅ Dependencies berhasil diinstall!"
    } else {
        Show-Error "❌ Gagal install dependencies"
        Show-Info "Silakan jalankan manual: npm install"
        exit 1
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
    
    # 9. Completed
    Write-Host ""
    Show-Success "🎉 INSTALASI BERHASIL!"
    Show-Info "Project location: $projectPath"
    Write-Host ""
    
    Show-Info "📋 FILE YANG DIBUAT:"
    Get-ChildItem | ForEach-Object { 
        Write-Host "   📄 $($_.Name)" -ForegroundColor Gray 
    }
    Write-Host ""
    
    Show-Info "🚀 Menjalankan TUL I-11 Tool..."
    Start-Sleep -Seconds 2
    
    # Run the tool
    Start-Process "start-tuli11.bat" -Wait
    
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
Show-Info "UNTUK NEXT TIME:"
Write-Host "   Buka folder: $projectPath" -ForegroundColor White
Write-Host "   Jalankan: start-tuli11.bat" -ForegroundColor White
Write-Host "   Atau langsung: node tulI11.js" -ForegroundColor White
Write-Host ""
Show-Info "Hanya install sekali, selanjutnya langsung run tool saja!"
Write-Host ""
