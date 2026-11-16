SEARCHABLE EXCEL PHRASE PICKER — DESKTOP EXCEL FOR WINDOWS

WHAT THIS DOES
- Works from your PERSONAL.XLSB macro workbook.
- Does NOT add any worksheets, macros or code to downloaded Excel files.
- Loads 30–60 phrases from a separate PhraseLibrary.xlsx workbook.
- Search phrases, add multiple, insert each on a separate line in one cell.
- The downloaded file can remain a normal .xlsx workbook.

STEP 1 — PLACE THE LIBRARY
- Save PhraseLibrary.xlsx somewhere permanent, such as Documents.
- Replace the sample phrases in column A, starting at A2, with your own.
- One complete phrase per row. Keep sheet name 'PhraseLibrary'.
- Save and close the library workbook.

STEP 2 — CREATE PERSONAL.XLSB (IF IT DOESN'T EXIST)
- Open Excel (desktop app, not browser).
- Developer > Record Macro (or View > Macros > Record Macro).
- Store macro in: Personal Macro Workbook.
- Click OK, then Stop Recording. This creates PERSONAL.XLSB.
- If prompted, save PERSONAL.XLSB when Excel closes.

STEP 3 — IMPORT MODULE
- Press Alt+F11 to open the VBA editor.
- In Project Explorer, select VBAProject (PERSONAL.XLSB).
- Choose File > Import File, choose PhrasePickerModule.bas.
- Check that the module appears under PERSONAL.XLSB, not downloaded workbook.

STEP 4 — MAKE THE USERFORM
- In the VBA editor, select VBAProject (PERSONAL.XLSB).
- Insert > UserForm.
- Change the UserForm (Name) property to frmPhrasePicker.
- Add these controls from the Toolbox, assigning each exact (Name) property:
    TextBox: txtSearch
    ListBox: lstMatches
    ListBox: lstSelected
    CommandButton: cmdAdd         Caption: Add phrase
    CommandButton: cmdRemove      Caption: Remove
    CommandButton: cmdInsert      Caption: Insert into cell
    CommandButton: cmdCancel      Caption: Cancel
- Make both ListBoxes wide enough to show your phrases.
- Set UserForm Caption to: Quick Phrase Picker.
- Right-click UserForm > View Code; paste all content of UserFormCode.txt.
- Save VBA changes; when Excel closes, confirm save to PERSONAL.XLSB.

STEP 5 — ASSIGN KEYBOARD SHORTCUT
- Open Excel > Developer > Macros (or Alt+F8).
- In the macro list select PERSONAL.XLSB!OpenPhrasePicker.
- Options > type capital J in Shortcut key box => Ctrl+Shift+J on Windows.
- Avoid a shortcut already important for your work.

EVERYDAY USE
1. Open downloaded Excel file.
2. Click destination cell.
3. Press Ctrl+Shift+J.
4. First time only, select your separate PhraseLibrary.xlsx when prompted.
5. Type a few letters; press Enter to add first matching phrase, or double-click.
6. Add more phrases using the search box.
7. Click Insert into cell. Phrases appear on separate lines in the same cell.
8. Save your downloaded workbook normally as .xlsx.

TO UPDATE YOUR PHRASES
- Edit and save PhraseLibrary.xlsx only. You can add up to many more rows.
- Reopen phrase picker — it rereads the phrase library automatically.

TO USE ANOTHER PHRASE LIBRARY
- Run PERSONAL.XLSB!ChooseDifferentPhraseLibrary in Excel's Macro dialog.

NOTES
- Requires desktop Excel with VBA/macros allowed. Does not run in Excel for the web.
- If a destination cell contains content, the picker asks before replacing it.
- It only writes the chosen cell and turns Wrap Text on for that cell.
- This version does not append to existing text or automatically resize the row.
- The macro is stored on the specific computer running Excel.
