# Excel Data Automation

Aplikasi desktop offline (GUI) berbasis Python + PySide6 untuk memudahkan data entry / otomatisasi data Excel (.xlsx).

## Fitur
- Upload file .xlsx
- Deteksi data duplikat
- Deteksi baris dengan kolom kosong
- Sheet `Data_Duplikat`, `Data_Kosong`
- Sheet `Data_Bersih` (tanpa duplikat & kosong, diurutkan & dikelompokkan)
- Sheet per kategori (misal on/off)
- Sheet `Charts` (Bar, Histogram, Line, Stacked, Scatter, Heatmap, Funnel, Pie)
- File original tetap, hasil disimpan sebagai file baru

## Menjalankan dari source
```bash
pip install -r requirements.txt
python main.py
```

## Build binary
```bash
pyinstaller --onefile --windowed --name ExcelDataAutomation main.py
```

## Binary
- Linux/macOS: `ExcelDataAutomation` (tanpa ekstensi)
- Windows: `ExcelDataAutomation.exe`
