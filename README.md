# PCCS_DAILYSHEET-AUTOMATION

WhatsApp se aayi AWB (Air Waybill) PDFs ko auto-detect karke daily Excel sheet mein fill karta hai.
CCU (Kolkata) + BLR (Bengaluru) dual-origin support. Ek hi engine — `PCCS_MASTER.py` — watcher,
extractor, filler, tray icon sab handle karta hai.

## Quick start
1. `pip install watchdog pypdf openpyxl pywin32 pystray Pillow`
2. `START.bat` (silent) ya `python PCCS_MASTER.py`
3. Pehli baar: setup wizard paths puchega. Blank chhodne par `AUTO` = auto-detect.
4. System tray mein ✈ icon → right-click menu.

## Zero-config paths (`config.json`)
`paths` mein `"AUTO"` likho to script khud detect karega:
- **watch_folder** → Windows `~/Downloads`
- **base_data_dir** → project ke andar `AWB_TOOLS/`
- **processed / blr_processed / archive** → `AWB_TOOLS/PROCESSED_CCU`, `PROCESSED_BLR`, `ARCHIVE`
- **extract_excel** → missing ho to blank scaffold (CCU+BLR headers) auto-ban jaayegi

Absolute path bhi de sakte ho — tab wahi use hoga. Zaroori folders boot pe auto-create hote hain.

## Features
- **Async pipeline** — PDF detect → rename → queue; ek background worker (COM ek hi baar init)
  sequential fill karta hai. User lagataar PDFs download kar sakta hai bina UI freeze/file-lock ke.
- **Dynamic Financial Year** — sheet naam clock se (`ccu_{fy}` → `ccu_2026-27`). Agar us naam ka
  tab `.xlsm` mein na ho to config ke static naam pe safe fallback.
- **Weekly auto-cleanup** — har 7 din: processed folders dated ZIP mein archive → delete,
  aur `AWB_EXTRACT-DATA.xlsx` ki purani rows trim (full backup ke saath). Tray se "Run Cleanup Now".
- **Dynamic Agent Master** — rename dialog mein naya naam type karo → "➕ Add as New Agent" button →
  `AGENTS_MASTER.xlsx` ke sahi sheet (CCU/BLR) mein add + cache turant refresh.
- **Column maps in config** — saare magic numbers `config.json` → `column_maps` mein. Daily sheet
  ka layout badle to sirf config edit, code nahi.

## Boot (Windows Task Scheduler — "On Logon")
- `MASTER_STARTUP.bat` → Chrome (WhatsApp Web + Gmail) + PCCS engine launch. Project-relative.
- Purana `WHATSAPP_WATCHER.py` **obsolete** — `PCCS_MASTER.py` khud watchdog observer chalata hai.

## Files
| File | Role |
|------|------|
| `PCCS_MASTER.py`        | Engine — watcher + extract + fill + tray + cleanup |
| `config.json`           | Paths, settings, financial_year, cleanup, column_maps |
| `AGENTS_MASTER.xlsx`    | Agent naam list (CCU / BLR sheets) |
| `AWB_EXTRACT-DATA.xlsx` | Intermediate staging (CCU / BLR sheets) |
| `session_memory.json`   | Last AWB + processed list + last_cleanup |
| `START.bat`             | Silent launcher (pythonw) |
| `MASTER_STARTUP.bat`    | Full boot orchestration |
