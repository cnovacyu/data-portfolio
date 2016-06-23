'Automates the batching of file processing, while calculating & graphing deviation'
Sub ProcessFiles()

Dim Filename, Pathname As String
Dim wbOpen As Workbook
Dim i As String
Dim j, k, x, y As Integer
Dim emptyRow, LastRow As Long

'This sets the filepath of where the files are stored to be processed
'and the name of the files, which is set up as a wildcard *.xls to
'apply the macro to all the files in the folder. **MAKE SURE THE
'MACRO FILE IS SAVED IN THE SAME FILE AS THE TEST FILES**
Pathname = "C:\Users\Username\Documents\Test\Test Files\"
Filename = Dir(Pathname & "*.xls")
j = 2
k = 1
x = 11
y = 11
emptyRow = 11

'Turn off any alert messages that may pop up while running macro for
'all the files.
Application.DisplayAlerts = False

'Turn off screen flashing while macro is running
Application.ScreenUpdating = False

'Loop statement checks the file folder in the directory named above
'and applies the macro in the while loop to all the files.

'Loop will continue until there are no other files in the folder
'left with  the extension of .xls left
Do While Filename <> ""

    'Opens up each test file in the folder
    Set wbOpen = Workbooks.Open(Pathname & Filename, True, False, True)
    
    'Loop is set for 6 tabs, since each file is consistent
    Do While j < 7
    
        'Activate each test file, else it looks at the macro workbook
        'as the active workbook
        i = wbOpen.Name
        Workbooks(i).Activate
    
        'Activate Sheet2 to copy Column B to Sheet1. Loop will apply
        'to Sheets 3 & 4, and Sheets 5 & 6 pairs
        Worksheets(j).Activate
        ActiveSheet.Range("B:B").Copy

        'Active Sheet 1 to paste into Column C. Loop will apply to
        'Sheets 3 & 4, and Sheets 5 & 6 pairs
        Worksheets(k).Activate
        ActiveSheet.Range("C1").Select
        ActiveSheet.Paste
    
        'Clears clipboard memory
        Application.CutCopyMode = False

        'Enter in "Deviation" in Cell D10
        ActiveSheet.Range("D10") = "DEVIATION"

        'Enter in the formula for Column D (5 - (B + C). Loops
        'until last row available
        Do While ActiveSheet.Cells(x, 2).Value <> ""
            Cells(emptyRow, 4) = 5 - (Cells(x, 2) + Cells(y, 3))
            x = x + 1
            y = y + 1
            emptyRow = emptyRow + 1
        Loop
        
        'Finds the last row so we can graph Column D
        LastRow = Range("D10").End(xlDown).Row

        'Create line graph from all rows in Column D
        Range("D10").Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Shapes.AddChart2(227, xlLine).Select
        ActiveChart.SetSourceData Source:=Range(Cells(10, 4), Cells(LastRow, 4))
        ActiveChart.Parent.Cut
        Range("F10").Select
        ActiveSheet.Paste
        
        'Reset variables in loops above, so they can repeat
        j = j + 2
        k = k + 2
        x = 11
        y = 11
        emptyRow = 11
        
    Loop
    
    'Close each file and save changes
    wbOpen.Close SaveChanges:=True
    Filename = Dir()
    
    'Reset variables in loops above, so they can repeat
    j = 2
    k = 1
    x = 11
    y = 11
    emptyRow = 11
    
Loop

Application.ScreenUpdating = True

End Sub