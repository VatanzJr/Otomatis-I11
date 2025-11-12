# Otomatis-I11
PLN TUL I-11 Automation Tool - By Vatanz

# otomatis-I11 - PLN TUL I-11 Automation Tool

Tool otomatis untuk proses TUL I-11 pada sistem PLN AP2T. Install dengan satu command PowerShell dan siap digunakan!

## 🚀 Cara Install & Menjalankan

### **Instalasi One-Click (Recommended)**
```powershell
irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex
```

### **Jika Error Execution Policy**
```powershell
# Jalankan sebagai Administrator terlebih dahulu
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Kemudian jalankan installer
irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex
```

### **Setelah Instalasi Berhasil**
```batch
# Buka folder instalasi
cd %USERPROFILE%\PLN-TULI11-Tool

# Jalankan tool
start-tuli11.bat
# atau
node tulI11.js
```

## ⚡ Fitur Utama

- ✅ **Auto login** ke sistem PLN AP2T
- ✅ **Ambil data otomatis** dari Google Sheet
- ✅ **Proses batch** multiple No Agenda sekaligus
- ✅ **Auto update** hasil ke Google Sheet
- ✅ **One-click installer** - Install sekali, pakai berkali-kali
- ✅ **Self-diagnose** - Auto detect dan fix issues
- ✅ **Offline mode** - Tetap bisa testing tanpa koneksi internet

## 🔧 System Requirements

- **Windows 10/11** dengan PowerShell 5.1+
- **Node.js 16+** ([Download](https://nodejs.org))
- **Koneksi internet** (untuk akses Google Sheet & sistem PLN)
- **Akses** ke sistem PLN AP2T

## 📋 Cara Kerja

1. **📥 Download & Install** - Auto download dependencies
2. **🔐 Login PLN** - Auto login ke sistem AP2T
3. **📋 Ambil Data** - Ambil daftar No Agenda dari Google Sheet
4. **🔄 Proses Data** - Cari NOPDL untuk setiap No Agenda
5. **📤 Update Hasil** - Auto update hasil ke Google Sheet

## 🛠 Troubleshooting

### **❌ Error: "Running scripts is disabled on this system"**
```powershell
# Jalankan PowerShell sebagai Administrator, lalu:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### **❌ Error: "Node.js is not recognized"**
- Download dan install Node.js dari [nodejs.org](https://nodejs.org)
- Pilih **LTS version**
- Restart komputer setelah install

### **❌ Error: "Cannot connect to Google Apps Script"**
1. Cek koneksi internet Anda
2. Buka browser dan test: https://script.google.com
3. Pastikan tidak ada firewall yang memblokir
4. Coba gunakan jaringan berbeda (Hotspot)

### **❌ Error: "Module not found"**
```batch
# Jalankan di folder tool:
npm install
# atau
start-tuli11.bat (pilih option 2)
```

### **❌ Error: Login gagal ke sistem PLN**
- Pastikan username dan password benar
- Pastikan akses ke https://ap2t.pln.co.id tidak diblokir
- Cek koneksi internet stabil

## 📁 Struktur Project

```
PLN-TULI11-Tool/
├── 📄 tulI11.js          # Main script
├── 📄 package.json       # Dependencies
├── 📄 start-tuli11.bat   # Launcher
└── 📄 test-deps.bat      # Dependency checker
```

## 🔄 Cara Update

```powershell
# Jalankan installer lagi - auto update
irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex
```

## 💡 Tips Penggunaan

### **Untuk Pertama Kali:**
1. Pastikan Node.js terinstall
2. Jalankan installer PowerShell
3. Tunggu sampai proses selesai
4. Tool akan otomatis running

### **Untuk Penggunaan Berikutnya:**
1. Buka folder `PLN-TULI11-Tool` di `C:\Users\[Username]\`
2. Jalankan `start-tuli11.bat`
3. Atau langsung `node tulI11.js`

### **Mode Offline/Testing:**
- Tool punya fallback data dummy
- Tetap bisa testing sistem meski tanpa koneksi internet
- Data dummy: TEST001, TEST002, TEST003

## 🎯 Konfigurasi

### **Credentials Sistem PLN:**
```javascript
Username: 9418672ZY
Password: mblendez
```

### **Google Apps Script:**
```
https://script.google.com/macros/s/AKfycbzuJYpjCG7YvDg40ImUMbH_vY1DIm7JBnhCP3APSddesyT6xB0pF5i6XhusEOWUcuTL1Q/exec
```

## 📞 Support

### **Laporkan Issues:**
1. Screenshot error message
2. Detail langkah yang dilakukan
3. Versi Node.js (`node --version`)

### **Update & Source Code:**
- GitHub: https://github.com/vatanzjr/otomatis-I11
- Issues: https://github.com/vatanzjr/otomatis-I11/issues

### **Contact:**
- Developer: vatanzjr
- Repository: otomatis-I11

## ⚠️ Disclaimer

- Tool ini untuk keperluan otomasi internal
- Pastikan memiliki akses yang sah ke sistem PLN
- Gunakan credentials yang benar dan authorized
- Developer tidak bertanggung jawab atas misuse

## 🆘 Emergency Fix

Jika semua troubleshooting gagal:

1. **Hapus folder lama:**
   ```powershell
   Remove-Item "$env:USERPROFILE\PLN-TULI11-Tool" -Recurse -Force
   ```

2. **Install ulang:**
   ```powershell
   irm https://raw.githubusercontent.com/vatanzjr/otomatis-I11/main/deploy.ps1 | iex
   ```

---

**⭐ Jika tool ini membantu, consider star repository ini!** 

**Happy Automating! 🚀**
