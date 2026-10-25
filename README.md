# AUTO-CAD-Import-Points
auto cad all files and points imports automatically



# Civil 3D COGO Points Batch Importer

This Python script automates the process of importing PNEZD points into multiple AutoCAD Civil 3D drawings (`.dwg`). It matches drawing files with their corresponding text files, imports the points as native COGO points, formats them, and saves the files automatically.

## 🌟 Key Features
- **Batch Processing:** Oka folder lo unna multiple `.dwg` files ni oke sari process chestundi.
- **Auto-Matching:** Drawing peru (`Project_A.dwg`) mariyu Text file peru (`Project_A.txt`) match ayithe automatic ga import chestundi.
- **Manual Fallback Popup:** Okavela exact match dorakakapothe, script aagi points file ni manual ga select chesukomani popup aduguthundi.
- **Native COGO Points:** Normal AutoCAD points laaga kakunda, direct ga Civil 3D API dwara Native COGO Points create chestundi.
- **Auto-Formatting:** 
  - Points anni `PNTS_asbuilt` layer loki velthayi.
  - Scale automatic ga `1:1` ki set avthundi.
  - Import ayyaka `PNTS_asbuilt` layer OFF ayyi, Zoom Extents (`Z -> E`) ayyi save & close avthundi.

---

## 📋 Prerequisites (Kavalsina Softwares)
Ee script run cheyyadaniki mee system lo ivi undali:
1. **AutoCAD Civil 3D** (Any version like 2020, 2022, 2024, etc.)
2. **Python 3.x** installed.
3. **pywin32 library:** Python nundi AutoCAD ni control cheyyadaniki idi kavali.

### Install Dependencies:
Command Prompt (CMD) leda VS Code terminal open chesi ee command run cheyyandi:
```bash
pip install pywin32
