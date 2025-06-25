# 🚛 Workers & Resources: Soviet Republic Vehicle Data Extractor

This script extracts vehicle data from *Workers & Resources: Soviet Republic* game files and exports it to a single, organized Excel spreadsheet.

---

## ✨ Features

- Parses all vehicle definitions from game and DLC folders.
- Extracts relevant vehicle data
- Supports filtering by a specific year. (Note: start and end date is chosen randomly within the year, so it may vary a bit)
- Currently only supports Early Start DLC as I don't own any others. Add other DLCs subfolders to the VEHICLE_SUBFOLDERS variable to also include them.

---

## 📁 Output

- A single Excel file: `vehicles.xlsx`  
  - If a year filter is used, the file will be named: `vehicles_<year>.xlsx`
- Sheets inside the Excel file:
  - `Road`, `Rail`, `Water`, `Air`
- Each row contains data for a single vehicle.

---

## 🛠 Setup

### 1. Requirements

- Python 3.7+
- Install dependencies:

```bash
pip install pandas openpyxl
```

**2. Prepare BTFTool.exe**
Download BTFTool.exe from: https://github.com/Nargon/BTFTool

Place it in a subfolder named tools next to the Python script:

📁 your_folder/
├── tools/
│   └── BTFTool.exe
│   └── BTFTool.exe.config
│   └── BTFTool.pdb
├── extract_vehicles.py

Change GAME_DATA_FOLDER inside the extract.py to your game data folder.

▶️ Usage
Run without filtering:
```
python extract_vehicles.py
```
Run with year filter (e.g., 1975):
```
python extract_vehicles.py 1975
```
✅ Only vehicles available in the given year will be included. As the start and end date is random within the year the vehicle might become available in a later month or has already gone out of service.

⚠️ Notes

    If a vehicle has no $RESOURCE_TRANSPORT_TYPE, but includes a skill like $SKILL_HARVESTING, that skill will be used as a proxy transport type and its level as capacity.
    One known issue is that fire trucks are somehow not getting labeled correctly.
    Currently no modded vehicles are supported. Might add that in the future.
