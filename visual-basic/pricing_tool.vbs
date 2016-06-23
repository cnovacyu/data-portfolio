'Pricing Tool - Calculates pricing based on customer type for all products'
Dim num As String, i As Integer, j As Integer, prod As String

Sub UserForm_Initialize()

'Clear all the textboxes on the form
PartTextBox.Value = ""
QtyTextBox1.Value = ""
QtyTextBox2.Value = ""
QtyTextBox3.Value = ""
QtyTextBox4.Value = ""
QtyTextBox5.Value = ""
CostTextBox1.Value = ""
CostTextBox2.Value = ""
CostTextBox3.Value = ""
CostTextBox4.Value = ""
CostTextBox5.Value = ""

'Add the drop down options for Prod Group. This list needs to be manually
'maintained as new prod groups are added or prod groups are obsoleted
With ProdGroupComboBox
    .AddItem "A001  Product 1"
    .AddItem "B001  Product 2"
    .AddItem "C001  Product 3"
    .AddItem "D001  Product 4"
    .AddItem "E001  Product 5"
    .AddItem "F001  Product 6"
    .AddItem "G001  Product 7"
    .AddItem "H001  Product 8"
    .AddItem "I001  Product 9"
    .AddItem "J001  Product 10"
    .AddItem "K001  Product 11"
    .AddItem "L001  Product 12"
    .AddItem "M001  Product 13"
    
End With

'Set cursor in Part Number text box when user launches user form
PartTextBox.SetFocus

End Sub

Sub CalcCommandButton_Click()

FillForm

End Sub

Sub ClearCommandButton_Click()

ClearForm

End Sub

Sub CloseCommandButton_Click()

Unload Me

End Sub


Sub ProdGroupComboBox_Change()

Dim prod As String
Dim qty As String

        'Take the first 4 letters of the product group and store into the prod variable
        'to vlookup other information
        prod = Left(ProdGroupComboBox, 4)
    
        'Look up the prod code variable in the Product Group tab to find the
        'purchasing group associated with that product code
        qty = Application.WorksheetFunction.VLookup(prod, Sheets("Product Groups").Range("$A$2:$E$500"), 4, False)
    
        'Lookup the quantity information for each purchasing group in the Pricing Groups tab
        QtyTextBox1.Value = Application.WorksheetFunction.VLookup(qty, Sheets("Pricing Groups").Range("$A$2:$F$100"), 2, False)
        QtyTextBox2.Value = Application.WorksheetFunction.VLookup(qty, Sheets("Pricing Groups").Range("$A$2:$F$100"), 3, False)
        QtyTextBox3.Value = Application.WorksheetFunction.VLookup(qty, Sheets("Pricing Groups").Range("$A$2:$F$100"), 4, False)
        QtyTextBox4.Value = Application.WorksheetFunction.VLookup(qty, Sheets("Pricing Groups").Range("$A$2:$F$100"), 5, False)
        QtyTextBox5.Value = Application.WorksheetFunction.VLookup(qty, Sheets("Pricing Groups").Range("$A$2:$F$100"), 6, False)

End Sub
Sub FillForm()

    'If the user does not enter a part number, does not choose a product group,
    'or does not enter all cost information, it will prompt them to do so with a
    'message box and continue looping until all the fields are filled out.
    If UserForm1.PartTextBox.Text = "" Then
        PartTextBox.SetFocus
        MsgBox "Please enter a Part Number", vbOKOnly, "Empty Field"
    ElseIf UserForm1.ProdGroupComboBox.Text = "" Then
        ProdGroupComboBox.SetFocus
        MsgBox "Please choose a Product Group", vbOKOnly, "Empty Field"
    ElseIf UserForm1.CostTextBox1 = "" Then
        CostTextBox1.SetFocus
        MsgBox "Please enter in all Cost Information", vbOKOnly, "Empty Field"
    ElseIf UserForm1.CostTextBox2 = "" Then
        CostTextBox2.SetFocus
        MsgBox "Please enter in all Cost Information", vbOKOnly, "Empty Field"
    ElseIf UserForm1.CostTextBox3 = "" Then
        CostTextBox3.SetFocus
        MsgBox "Please enter in all Cost Information", vbOKOnly, "Empty Field"
    ElseIf UserForm1.CostTextBox4 = "" Then
        CostTextBox4.SetFocus
        MsgBox "Please enter in all Cost Information", vbOKOnly, "Empty Field"
    ElseIf UserForm1.CostTextBox5 = "" Then
        CostTextBox5.SetFocus
        MsgBox "Please enter in all Cost Information", vbOKOnly, "Empty Field"
    Else
        Calculate
    End If

End Sub

Sub Calculate()

Dim emptyRow As Long
Dim cost1 As Double
Dim cost2 As Double
Dim cost3 As Double
Dim cost4 As Double
Dim cost5 As Double
Dim land1 As Double
Dim land2 As Double
Dim land3 As Double
Dim land4 As Double
Dim land5 As Double
Dim bur As Double
Dim price As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

    'Store the cost inputs from the user form into these variables
    cost1 = CostTextBox1.Value
    cost2 = CostTextBox2.Value
    cost3 = CostTextBox3.Value
    cost4 = CostTextBox4.Value
    cost5 = CostTextBox5.Value
    
    'Take the first 4 letters of the product group and store into the prod variable
    'to vlookup other information
    prod = Left(ProdGroupComboBox, 4)
    
    'Look up the prod code variable in the Product Group tab to find the
    'pricing group associated with that product code
    price = Application.WorksheetFunction.VLookup(prod, Sheets("Product Groups").Range("$A$2:$E$500"), 4, False)

    'Activate Sheet2 (Pricing Calculator) since this tab is where we want to paste
    'in our values and calculate pricing
    Sheet2.Activate
    
    'Determine the next empty row. I typed in values in A1:A4 and changed font color
    'to white to force the count to be 5 when selecting Column A as a default. This
    'will force the "first" empty row to be Row 6.
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
    
    'Copy and paste the part number entered in the user form into Column A in the first
    'empty row in the Pricing Calculator tab
    Cells(emptyRow, 1).Value = PartTextBox.Text
    
    'Copy and paste the product group entered in the user form into Column B in the first
    'empty row in the Pricing Calculator tab
    Cells(emptyRow, 2).Value = ProdGroupComboBox.Text
    
    'Vlookup the material burden rate of the product group in the Product Groups sheet
    'and paste into Column C in the first empty row in the Pricing Calculator tab. Also store
    'the landed cost % into a variable for the landed cost calculation
    Cells(emptyRow, 3) = Application.WorksheetFunction.VLookup(prod, Sheets("Product Groups").Range("$A$2:$E$500"), 3, False)
    
    bur = Application.WorksheetFunction.VLookup(prod, Sheets("Product Groups").Range("$A$2:$E$500"), 3, False)
    
    'Calculate landed cost
    land1 = cost1 * (1 + bur)
    land2 = cost2 * (1 + bur)
    land3 = cost3 * (1 + bur)
    land4 = cost4 * (1 + bur)
    land5 = cost5 * (1 + bur)
    
    'Vlookup the price group of the product group in the Product Groups sheet and paste
    'into Column D of in the first empty row of the Pricing Calculator tab
    Cells(emptyRow, 4) = Application.WorksheetFunction.VLookup(prod, Sheets("Product Groups").Range("$A$2:$D$500"), 4, False)
    
    'Set j = 5 for the Intercompany pricing calculation in the if statement below. j is the column
    'number for the first price break of the Dist column. j will increase by 1, going to the
    'next Dist price break columns to calculate Intercompany pricing
    j = 5
    
    'If the pricing group is "Special Pricing - See PM" then output N/A on all columns for
    'Dist, OEM, and Intercompany pricing
    If price = "Special Pricing - See PM" Then
        For i = 5 To 28
        Cells(emptyRow, i) = "N/A"
        Next i
    Else
        'Else, calculate price by taking the user cost inputs / (1 - GM%) which is vlooked
        'up using the pricing group
        
        'Distribution price calculation based on the GM% set in the Pricing Groups tab. Need to
        'maintain this tab, add any changes, new groups, etc
        Cells(emptyRow, 5) = Round(land1 / (1 - (Application.WorksheetFunction.VLookup(price, Sheets("Pricing Groups").Range("$A$2:$K$50"), 7, False))), 3)
        Cells(emptyRow, 6) = Round(land2 / (1 - (Application.WorksheetFunction.VLookup(price, Sheets("Pricing Groups").Range("$A$2:$K$50"), 8, False))), 3)
        Cells(emptyRow, 7) = Round(land3 / (1 - (Application.WorksheetFunction.VLookup(price, Sheets("Pricing Groups").Range("$A$2:$K$50"), 9, False))), 3)
        Cells(emptyRow, 8) = Round(land4 / (1 - (Application.WorksheetFunction.VLookup(price, Sheets("Pricing Groups").Range("$A$2:$K$50"), 10, False))), 3)
        Cells(emptyRow, 9) = Round(land5 / (1 - (Application.WorksheetFunction.VLookup(price, Sheets("Pricing Groups").Range("$A$2:$K$50"), 11, False))), 3)
        Cells(emptyRow, 10) = Round(Cells(emptyRow, 9).Value * 0.96, 3)
        Cells(emptyRow, 11) = Round(Cells(emptyRow, 10).Value * 0.98, 3)
        Cells(emptyRow, 12) = Round(Cells(emptyRow, 11).Value * 0.98, 3)
        
        ' price calculation
        Cells(emptyRow, 13) = Round(Cells(emptyRow, 5).Value / (1 - 0.35), 3)
        Cells(emptyRow, 14) = Round(Cells(emptyRow, 6).Value / (1 - 0.3), 3)
        Cells(emptyRow, 15) = Round(Cells(emptyRow, 7).Value / (1 - 0.25), 3)
        Cells(emptyRow, 16) = Round(Cells(emptyRow, 8).Value / (1 - 0.25), 3)
        Cells(emptyRow, 17) = Round(Cells(emptyRow, 9).Value / (1 - 0.2), 3)
        Cells(emptyRow, 18) = Round(Cells(emptyRow, 10).Value / (1 - 0.18), 3)
        Cells(emptyRow, 19) = Round(Cells(emptyRow, 11).Value / (1 - 0.15), 3)
        Cells(emptyRow, 20) = Round(Cells(emptyRow, 12).Value / (1 - 0.1), 3)
        
        'Intercompany pricing calculation. For loop is applying the same calculation for
        'all Intercompany price breaks based on the distribution columns, which increase
        ' by 1 in the loop statement
        For k = 21 To 28
            Cells(emptyRow, k) = Round(Cells(emptyRow, j).Value * 0.85, 3)
            j = j + 1
        Next k
        
    End If

End Sub

Sub ClearForm()

    'If the user clears the form, all fields will be cleared out. I set the Prod Group
    'Combo box to the first one in the list, otherwise the ProdGroupComboBox_Change Sub
    'errors out since there is no prod group to lookup to display the purchase quantities
    UserForm1.Controls("PartTextBox").Text = ""
    
    UserForm1.Controls("ProdGroupComboBox").Text = "A001  Product 1"

For i = 1 To 5
    UserForm1.Controls("QtyTextBox" & i).Text = ""
Next i

For j = 1 To 5
    UserForm1.Controls("CostTextBox" & j).Text = ""
Next j

End Sub



'Open the user form when clicking on the "Calculate Pricing Button" on Pricing Calculator sheet
Sub Button6_Click()

UserForm1.Show

End Sub

'Clear Pricing Calculator worksheet when clicking on the "Clear" button on Pricing Calculator sheet
Sub Button3_Click()

Sheet2.Activate

Range("$A$7:$AE$2500") = ""

End Sub


'Copy horizontal prices from the Pricing Calculator Sheet and paste into the
'Convert to Vertical Layout sheet when clicking on the "Convert to Vertical
'Pricing" button on the Convert to Vertical Layout tab.
Sub Button1_Click()

    Dim i As Integer
    Dim j As Integer
    
    Sheet3.Activate
    
    'Bold Columns D, G, and J in Sheet3 (Convert to Vertical Layout) tab
    Columns("D").Font.Bold = True
    Columns("G").Font.Bold = True
    Columns("J").Font.Bold = True
    
    'Right align Columns D7:K2500 in Sheet 3 (Convert to Vertical Layout) tab
    Range("D7:K2500").HorizontalAlignment = xlRight

    i = 7
    j = 7
    Do While Sheets("Pricing Calculator").Cells(i, 1).Value <> ""
        Sheets("Convert to Vertical Layout").Cells(j, 1).Value = Sheets("Pricing Calculator").Cells(i, 1).Value
        Sheets("Convert to Vertical Layout").Cells(j, 2).Value = 1
        Sheets("Convert to Vertical Layout").Cells(j, 5).Value = 1
        Sheets("Convert to Vertical Layout").Cells(j, 8).Value = 1
        Sheets("Convert to Vertical Layout").Cells(j, 3) = Sheets("Pricing Calculator").Cells(i, 5).Value
        Sheets("Convert to Vertical Layout").Cells(j, 6) = Sheets("Pricing Calculator").Cells(i, 13).Value
        Sheets("Convert to Vertical Layout").Cells(j, 9) = Sheets("Pricing Calculator").Cells(i, 21).Value
        Sheets("Convert to Vertical Layout").Cells(j + 1, 2).Value = 50
        Sheets("Convert to Vertical Layout").Cells(j + 1, 5).Value = 50
        Sheets("Convert to Vertical Layout").Cells(j + 1, 8).Value = 50

        i = i + 1
        j = j + 9
    Loop
    
End Sub
'Clear Convert to Vertical Layout worksheet when clicking on the "Clear" button on Convert to Vertical Layout sheet
Sub Button2_Click()

    Sheet3.Activate

    Range("A7:K2500") = ""

End Sub