
Private Sub CommandButton3_Click()
Unload UserForm1
End Sub


Private Sub CommandButton5_Click()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Sheet2")

    Dim r As Long, lastRow As Long
    lastRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row

    Dim targetID As Long
    If IsNumeric(Me.TextBox11.Value) Then
        targetID = CLng(Me.TextBox11.Value)
    Else
        MsgBox "Invalid internal ID", vbCritical
        Exit Sub
    End If

    For r = 2 To lastRow
        If Ws.Cells(r, 1).Value = targetID Then
            Ws.Cells(r, 2).Value = Me.ComboBox2.Value
            Ws.Cells(r, 3).Value = Me.ComboBox3.Value
            Ws.Cells(r, 4).Value = Me.ComboBox1.Value
            Ws.Cells(r, 5).Value = Me.ComboBox5.Value
            Ws.Cells(r, 6).Value = Me.TextBox6.Value
            Ws.Cells(r, 7).Value = Me.ComboBox4.Value
            Ws.Cells(r, 8).Value = Me.TextBox8.Value
            MsgBox "Entry updated successfully!", vbInformation
            Exit Sub
        End If
    Next r

    MsgBox "ID not found to update!", vbExclamation
End Sub

Private Sub TextBox17_Change()
    If Trim(Me.TextBox17.Value) = "" Then Exit Sub

    If Not IsNumeric(Me.TextBox17.Value) Then
        MsgBox "Please enter value in numeric format", vbExclamation
        Exit Sub
    End If

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Sheet2")

    Dim r As Long, lastRow As Long
    Dim found As Boolean
    found = False

    lastRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Ws.Cells(r, 1).Value = CLng(Me.TextBox17.Value) Then
            Me.TextBox11.Value = Ws.Cells(r, 1).Value
            Me.ComboBox2.Value = Ws.Cells(r, 2).Value
            Me.ComboBox3.Value = Ws.Cells(r, 3).Value
            Me.ComboBox1.Value = Ws.Cells(r, 4).Value
            Me.ComboBox5.Value = Ws.Cells(r, 5).Value
            Me.TextBox6.Value = Ws.Cells(r, 6).Value
            Me.ComboBox4.Value = Ws.Cells(r, 7).Value
            Me.TextBox8.Value = Ws.Cells(r, 8).Value
            found = True
            Exit For
        End If
    Next r

    If Not found Then
        MsgBox "No Entry Found", vbExclamation
        Me.TextBox11.Value = ""
        Me.ComboBox2.Value = ""
        Me.ComboBox3.Value = ""
        Me.ComboBox5.Value = ""
        Me.TextBox6.Value = ""
        Me.ComboBox4.Value = ""
        Me.TextBox8.Value = ""
    End If
End Sub




Private Sub UserForm_Initialize()
    Dim Ws As Worksheet
    Dim LastColumn As Long
    Dim ColmnLoop As Long
    Dim lastRow As Long
    Dim RowLoop As Long

    Set Ws = ThisWorkbook.Worksheets("Sheet1")
    'ComboBox2
    LastColumn = Ws.Cells(1, Ws.Columns.Count).End(xlToLeft).Column
    For ColmnLoop = 6 To LastColumn
        If Ws.Cells(1, ColmnLoop).Value <> "" Then
            UserForm1.ComboBox2.AddItem Ws.Cells(1, ColmnLoop).Value
        End If
    Next ColmnLoop
    ' ComboBox3
    With Me.ComboBox3
        .AddItem "Week 1"
        .AddItem "Week 2"
        .AddItem "Week 3"
        .AddItem "Week 4"
    End With
    ' ComboBox1
    Dim ColumnRng As Range
    Dim CRowLoop As Long
    Set ColumnRng = Ws.Range("D4")
    For CRowLoop = ColumnRng.Row To 20
        If Ws.Cells(CRowLoop, ColumnRng.Column).Value <> "" Then
            UserForm1.ComboBox1.AddItem Ws.Cells(CRowLoop, ColumnRng.Column).Value
        End If
    Next CRowLoop
    'ComboBox4
    With Me.ComboBox4
        .AddItem "Cash"
        .AddItem "Card"
        .AddItem "Irish UPI"
    End With
    'TextBox 1
Dim WSTWO As Worksheet
Dim LastRowTWO As Long
Dim LastSerial As Variant
Set WSTWO = ThisWorkbook.Worksheets("Sheet2")
LastRowTWO = WSTWO.Cells(WSTWO.Rows.Count, 1).End(xlUp).Row
If LastRowTWO <= 1 Then
    Me.TextBox11.Value = 1
Else
    LastSerial = WSTWO.Cells(LastRowTWO, 1).Value
    If IsNumeric(LastSerial) Then
        Me.TextBox11.Value = CLng(LastSerial) + 1
    Else
        Me.TextBox11.Value = 1
    End If
End If
Me.TextBox11.Enabled = False
End Sub

Private Sub ComboBox1_Change()
    Dim GroceriesRange As Range
    Dim TransportRange As Range
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Worksheets("Sheet1")
    Dim cell As Range
    Me.ComboBox5.Clear
    Set GroceriesRange = Ws.Range("E6:E15")
    Set TransportRange = Ws.Range("E16:E19")
    If Me.ComboBox1.Value = "Groceries" Then
    For Each cell In GroceriesRange
        If cell.Value <> "" Then Me.ComboBox5.AddItem cell.Value
    Next cell
    ElseIf Me.ComboBox1.Value = "Transport" Then
    For Each cell In TransportRange
        If cell.Value <> "" Then Me.ComboBox5.AddItem cell.Value
    Next cell
End If
End Sub
Private Sub TextBox6_AfterUpdate()
If IsNumeric(Me.TextBox6.Value) Then
        Me.TextBox6.Value = Format(Me.TextBox6.Value, "€#,##0.00")
    Else
        MsgBox "Please enter a valid number in Number Format."
        Me.TextBox6.Value = ""
    End If
End Sub
Private Sub CommandButton1_Click()
Dim PastingValuesRow As Long
Dim Ws As Worksheet
Set Ws = ThisWorkbook.Worksheets("Sheet2")
PastingValuesRow = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row + 1

    Ws.Cells(PastingValuesRow, 1).Value = Me.TextBox11.Value
    Ws.Cells(PastingValuesRow, 2).Value = Me.ComboBox2.Value
    Ws.Cells(PastingValuesRow, 3).Value = Me.ComboBox3.Value
    Ws.Cells(PastingValuesRow, 4).Value = Me.ComboBox1.Value
    Ws.Cells(PastingValuesRow, 5).Value = Me.ComboBox5.Value
    Ws.Cells(PastingValuesRow, 6).Value = Me.TextBox6.Value
    Ws.Cells(PastingValuesRow, 7).Value = Me.ComboBox4.Value
    Ws.Cells(PastingValuesRow, 8).Value = Me.TextBox8.Value

End Sub

Private Sub CommandButton4_Click()
Dim Ws As Worksheet
Dim Number As Long
Set Ws = ThisWorkbook.Worksheets("Sheet2")
Number = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row
Me.ComboBox2.Value = ""
Me.ComboBox3.Value = ""
Me.ComboBox1.Value = ""
Me.ComboBox5.Value = ""
Me.TextBox6.Value = ""
Me.ComboBox4.Value = ""
Me.TextBox8.Value = ""
Me.TextBox17.Value = ""
Me.TextBox11.Value = Number
End Sub
