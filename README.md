# Excel Phrase Picker

A searchable phrase picker for **Microsoft Excel desktop on Windows**. Select one or more saved phrases and insert them into the active cell, **one phrase per line**.

## How it works

- Works with any downloaded workbook you open in desktop Excel; no extra sheets or VBA code are added to the downloaded workbook.
- VBA code lives in Excel's hidden **Personal Macro Workbook** (`PERSONAL.XLSB`).
- Phrases live in a **separate** editable `PhraseLibrary.xlsx` file, with one phrase in each cell of column A on the `PhraseLibrary` sheet (starting at A2).
- Search is case-insensitive and matches any part of a phrase.
- Add multiple phrases before clicking **Insert into cell**.
- Text in the destination cell is replaced only after confirmation; Excel text wrapping is enabled.

## Files

| File | Purpose |
| --- | --- |
| [PhraseLibrary.xlsx](PhraseLibrary.xlsx) | Phrase library with sample entries to replace. |
| [PhrasePickerModule.bas](PhrasePickerModule.bas) | Macro to launch the picker and remember the library location. |
| [UserFormCode.txt](UserFormCode.txt) | Code to paste into the VBA UserForm. |
| [SETUP_README.txt](SETUP_README.txt) | Step-by-step setup and usage instructions. |

## Quick setup

1. Download `PhraseLibrary.xlsx`, save it in a permanent folder, and edit phrases in column A.
2. In desktop Excel, create `PERSONAL.XLSB` by recording and stopping a macro with **Store macro in: Personal Macro Workbook**.
3. Press **Alt + F11**. In `VBAProject (PERSONAL.XLSB)`, import `PhrasePickerModule.bas` using **File > Import File**.
4. Under the same VBA project, use **Insert > UserForm**. Name it `frmPhrasePicker`.
5. Add controls with the exact names: `txtSearch` (TextBox), `lstMatches` and `lstSelected` (ListBoxes), and `cmdAdd`, `cmdRemove`, `cmdInsert`, `cmdCancel` (CommandButtons).
6. Paste `UserFormCode.txt` into the UserForm's code window.
7. Save `PERSONAL.XLSB`. In Excel's Macro dialog (**Alt + F8**), assign a shortcut such as **Ctrl + Shift + J** to `PERSONAL.XLSB!OpenPhrasePicker`.

Open any downloaded workbook, select a cell, run `OpenPhrasePicker`, and choose the library file the first time. Search and add phrases; click **Insert into cell** to insert them as separate lines.

For full instructions, read [SETUP_README.txt](SETUP_README.txt).

## Requirements and limitations

- Requires Microsoft Excel **desktop** on Windows, with VBA macros permitted by your organisation. It does not run in Excel for the web.
- The picker writes only to the selected target cell and sets **Wrap Text** for that cell; it does not add anything to the downloaded workbook.
- The first-time installation does require creating a VBA UserForm manually.
- Library phrases are examples; edit the local `PhraseLibrary.xlsx` to store your own phrases.
- If your organisation blocks personal macros, Excel add-ins or VBA execution, ask IT before attempting installation.
