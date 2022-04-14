Attribute VB_Name = "Module1"
'Force explicit variable declaration
Option Explicit

'Global variable that controls the writing index of components to the "Aux" sheet
Public AuxIndex As Integer

'Boolean function that checks existance of worksheet
Function CheckIfSheetExists(SheetName As String) As Boolean
    Dim WS As Worksheet
    
    CheckIfSheetExists = False
    For Each WS In Worksheets
        If SheetName = WS.name Then
            CheckIfSheetExists = True
            Exit Function
        End If
    Next WS
End Function

'Start button execute code: Displays UserForm and creates "Aux" sheet for the first time it executes
Sub Botão4_Click()
    UserForm1.Show False
    If CheckIfSheetExists("Aux") <> True Then
        Sheets.Add.name = "Aux"
        Sheets("Aux").Move After:=Sheets("Alunos")
        'Sheets("Aux").Visible = False
        'Store number of students in Cell "E1" of "Aux" sheet
        Sheets("Aux").Range("E1").formula = "=COUNTA(Alunos!A:A)-1"
        AuxIndex = 1
    End If
End Sub

'Returns the index of an item in a ComboBox based on it's name
Function SelectItemComboBox(Cb As ComboBox, name As String) As Integer
    Dim size As Integer
    Dim idx As Integer
    
    SelectItemComboBox = -1
    size = Cb.ListCount
    For idx = 0 To size - 1
        If Cb.List(idx) = name Then
            SelectItemComboBox = idx
            Exit Function
        End If
    Next idx
End Function

'Returns the column of the cell in sheet "Grupos" that a group was stored in
Function FindGrouping(sheet As Worksheet, name As String) As Integer
    Dim col As Integer
    Dim idx As Integer
    
    FindGrouping = -1
    col = sheet.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).column - 1
    For idx = 1 To col
        If sheet.Cells(1, idx).Value = name Then
            FindGrouping = idx
            Exit Function
        End If
    Next idx
End Function

'Returns the column of the cell in sheet "Aux" that a comp was stored in
Function FindCompIndex(sheet As Worksheet, name As String) As Integer
    Dim idx As Integer
    
    FindCompIndex = -1
    idx = 1
    Do While idx < AuxIndex
        If sheet.Cells(1, idx).Value = name Then
            FindCompIndex = idx
            Exit Function
        Else
            idx = idx + 5
        End If
    Loop
End Function

'Returns the row of the cell in sheet "Aux" that a parameter was stored in based on a comp
Function FindParIndex(sheet As Worksheet, idxcomp As Integer, par As String) As Integer
    Dim idx As Integer
    Dim n As Integer
    
    FindParIndex = -1
    n = sheet.Cells(1, idxcomp + 3).Value
    For idx = 2 To n + 1
        If sheet.Cells(idx, idxcomp).Value = par Then
            FindParIndex = idx
            Exit Function
        End If
    Next idx
End Function

'Function that replaces old_group for new_group for all components in listbox with old_group
Sub Correct(Lb As listbox, name As String, New_Name As String)
    Dim size As Integer
    Dim idx As Integer
    
    size = Lb.ListCount
    For idx = 0 To size - 1
        If Lb.List(idx, 2) = name Then
            Lb.List(idx, 2) = New_Name
        End If
    Next idx
End Sub

'Function that fills a list box with all the parameters stored in memory (aka the "Aux" sheet)
Sub FillListBox(sheet As Worksheet, comp As String, listbox As MSForms.listbox)
    Dim idxcomp As Integer
    Dim n As Integer
    Dim i As Integer
    Dim idx As Integer
    Dim parametro As String
    Dim peso As Double
    
    idxcomp = FindCompIndex(sheet, comp)
    listbox.Clear
    n = sheet.Cells(1, idxcomp + 3).Value
    For i = 2 To n + 1
        idx = listbox.ListCount
        parametro = sheet.Cells(i, idxcomp).Value
        peso = sheet.Cells(i, idxcomp + 1).Value
        listbox.AddItem
        listbox.List(idx, 0) = parametro
        listbox.List(idx, 1) = peso
    Next i
End Sub
'Function that creates and populates an array with the correct values given a component
Function CreateArray(idxcomp As Integer, sheetAux As Worksheet, sheetAlunos As Worksheet, n_down As Integer, n_across As Integer) As Variant
    Dim TempArray() As Variant
    Dim i As Integer
    
    ReDim TempArray(1 To n_down, 1 To n_across)
    TempArray(2, 1) = "Alunos"
    
    For i = 3 To n_down
        TempArray(i, 1) = sheetAlunos.Cells(i + 5, 1).Value
    Next i
    
    For i = 2 To n_across - 1
        TempArray(1, i) = sheetAux.Cells(i, idxcomp + 1).Value
        TempArray(2, i) = sheetAux.Cells(i, idxcomp).Value
    Next i
    
    CreateArray = TempArray
End Function
'Function that creates and populates an array with the correct values given a component with a grouping
Function CreateArrayGroup(idxcomp As Integer, sheetAux As Worksheet, n_down As Integer, n_across As Integer) As Variant
    Dim TempArray() As Variant
    Dim i As Integer
    
    ReDim TempArray(1 To n_down, 1 To n_across)
    
    For i = 1 To n_across - 1
        TempArray(1, i) = sheetAux.Cells(i + 1, idxcomp + 1).Value
        TempArray(2, i) = sheetAux.Cells(i + 1, idxcomp).Value
    Next i
    
    CreateArrayGroup = TempArray
End Function

'Function that converts column number to column letter
Function Number2Letter(column As Integer) As String
Dim ColumnNumber As Long
Dim ColumnLetter As String

Number2Letter = Split(Cells(1, column).Address, "$")(1)
End Function

'Function that adds a component column to the "Alunos" sheet
Sub UpdateAlunos(sheetAlunos As Worksheet, n_down As Integer, n_across As Integer, weight As Double, name As String)
    Dim col As Integer
    Dim i As Integer
    Dim formula As String
    
    col = sheetAlunos.Cells(7, Columns.Count).End(xlToLeft).Offset(0, 1).column
    sheetAlunos.Cells(6, col).Value = weight
    sheetAlunos.Cells(7, col).Value = name
    For i = 8 To 8 + n_down - 3
        formula = "=VLOOKUP(A" & CStr(i) & "," & name & "!A3:" & Number2Letter(n_across) & CStr(n_down) & "," & CStr(n_across) & ",FALSE)*" & Number2Letter(col) & "6/" & name & "!" & Number2Letter(n_across) & "1"
        sheetAlunos.Cells(i, col).formula = formula
    Next i
End Sub

'Function that adds a component column to the "Alunos" sheet with a grouping
Sub UpdateAlunosGroup(sheetAlunos As Worksheet, sheetGroup As Worksheet, n_down As Integer, n_across As Integer, weight As Double, name As String, agrupamento As String, n_alunos As Integer)
    Dim col As Integer
    Dim i As Integer
    Dim formula As String
    Dim idxgroup As Integer
    Dim group_col As Integer
    
    group_col = sheetAlunos.Cells(7, Columns.Count).End(xlToLeft).Offset(0, 1).column
    sheetAlunos.Cells(7, group_col).Value = agrupamento
    idxgroup = FindGrouping(sheetGroup, agrupamento)
    For i = 8 To 8 + n_alunos - 1
        formula = "=Grupos!" & Number2Letter(idxgroup) & CStr(i - 6)
        sheetAlunos.Cells(i, group_col).formula = formula
    Next i
    
    col = sheetAlunos.Cells(7, Columns.Count).End(xlToLeft).Offset(0, 1).column
    sheetAlunos.Cells(6, col).Value = weight
    sheetAlunos.Cells(7, col).Value = name
    For i = 8 To 8 + n_alunos - 1
        formula = "=VLOOKUP(" & Number2Letter(group_col) & CStr(i) & "," & name & "!A3:" & Number2Letter(n_across) & CStr(n_down) & "," & CStr(n_across) & ",FALSE)*" & Number2Letter(col) & "6/" & name & "!" & Number2Letter(n_across) & "1"
        sheetAlunos.Cells(i, col).formula = formula
    Next i
End Sub
