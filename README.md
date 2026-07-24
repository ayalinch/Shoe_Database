# KITE UHN Shoe Database

A desktop application for managing and viewing the KITE Research Institute's footwear evaluation database.

## How to Run

To run the standalone application:

1. Open the **`Compiled`** folder.
2. Make sure the `Photos and Reports` folder and the `WinterLab Master list of footwear.xlsx` Excel file are located inside this folder alongside the `.exe`.
3. Double-click **`KiteShoeDatabase.exe`** to start the application.

*(The application will automatically generate a `shoes.db` database and `images` folder upon running within the relative directory).*

## Development

To run the application from source code:

```bash
pip install PySide6 openpyxl
python Database.py
```

### Compiling to EXE

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name "KiteShoeDatabase" --add-data "68556ca78f14ebbed4120b97_Blue-KITE.png;." Database.py
```

After compiling, move the newly generated `KiteShoeDatabase.exe` from the `dist` folder into your `Compiled` folder. The leftover `build`, `dist`, and `.spec` files can be deleted without issue.
