# Build executabil și setup

1. Rulează `.\scripts\build_release.ps1`.
2. Executabilul PyInstaller va fi generat în `dist\TransareApp\TransareApp.exe`.
3. Dacă Inno Setup 6 este instalat, scriptul generează și `dist\installer\TransareApp-Setup.exe`.

Observații:

- Iconița executabilului este generată din `letter-s2.png`.
- Template-ul Excel din `assets\salarii_template_preset.xlsx` este inclus automat în build.
- În build-ul instalat, output-ul implicit merge în `Documents\TransareApp\output`, nu în folderul de instalare.
