# 🚛 Workers & Resources: Soviet Republic Vehicle Data Extractor

This script extracts vehicle data from *Workers & Resources: Soviet Republic* game files and exports it to a single, organized Excel spreadsheet.

---

## ✨ Features

- Parses all vehicle definitions from game and DLC folders
- Extracts relevant vehicle data
- Supports filtering by a specific year. (Note: start and end date is chosen randomly within the year, so actual availability may vary a bit)
- Second script to visualize data per vehicle type, also warns about gaps without available vehicles of a type
![road_harvesting](https://github.com/user-attachments/assets/8c454b3c-06b9-41af-bfd5-4362882caf42)

---

## 🛠 Setup

### 1. Requirements

- Python 3.7+
- Install dependencies:

```bash
pip install pandas openpyxl matplotlib
```

**2. Prepare BTFTool.exe**
Download BTFTool.exe from: https://github.com/Nargon/BTFTool

Place it in a subfolder named tools next to the Python script:
```
📁 your_folder/
├── tools/
│   └── BTFTool.exe
│   └── BTFTool.exe.config
│   └── BTFTool.pdb
├── extract_vehicles.py
```
Change GAME_DATA_FOLDER inside the extract.py to your game data folder.

**3. Usage**
Run without filtering:
```
python extract_vehicles.py
```
Run with year filter (e.g., 1975):
```
python extract_vehicles.py 1975
```
Generate diagrams and check for gaps (after the first script):
```
python generate_vehicle_diagrams.py
```
✅ Only vehicles available in the given year will be included. As the start and end date is random within the year the vehicle might become available in a later month or has already gone out of service.

⚠️ Notes

    I could only test the Early Start DLC as I don't own any others. Added the other DLCs with the same scheme. If it differs you may have to modify the VEHICLE_SUBFOLDERS variable.
    Kyrillic letters are not displayed in the diagrams, as the default fonts are not able to.
    Currently no modded vehicles are supported. Might add that in the future.
