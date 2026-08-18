# TransareApp

Aplicatie desktop pentru completarea rapoartelor salariale pe baza fisierelor Excel de lucru.

## Rulare locala

```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\python app.py
```

## Teste

```powershell
.\.venv\Scripts\python -m pytest
```

## Build executabil

```powershell
.\scripts\build_release.ps1
```

Build-ul genereaza executabilul in `dist\TransareApp\TransareApp.exe`.
Daca Inno Setup 6 este instalat, scriptul genereaza si installerul in `dist\installer\TransareApp-Setup.exe`.

## Fisiere exemplu

- `1. TRANSATORI+ DETINUTI 2026.xlsx`
- `fisiere de test\`
- `assets\salarii_template_preset.xlsx`

Folderele generate local (`build`, `dist`, `output`, `.venv`, cache-uri) nu sunt versionate.
