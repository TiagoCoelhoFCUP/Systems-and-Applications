VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Avaliação"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20670
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Force explicit variable declaration
Option Explicit

Private Sub AddComp_Click()
    Dim idx As Integer
    Dim componente As String
    Dim peso As Double
    Dim idxcomp As Integer
    
    'If we got valid inputs
    If TxtComp.Value <> "" And TxtPesoC.Value <> "" Then
        'If we are correcting a component
        If AddComp.Caption = "Corrigir" Then
            idx = ListBoxC.ListIndex
            componente = TxtComp.Value
            peso = TxtPesoC.Value
        
            'Get that component's position in the aux table
            idxcomp = FindCompIndex(Sheets("Aux"), ListBoxC.List(idx, 0))
            'Write to that position the new name and weight values
            Sheets("Aux").Cells(1, idxcomp).Value = componente
            Sheets("Aux").Cells(1, idxcomp + 1).Value = peso
            'Write to the ListBox the new namee and weight values
            ListBoxC.List(idx, 0) = componente
            ListBoxC.List(idx, 1) = peso
        
            'If the modification has a grouping
            If CheckBoxAgr.Value = True And ComboBoxAgr.ListIndex <> -1 Then
                'Write to the ListBox the new grouping value
                ListBoxC.List(idx, 2) = ComboBoxAgr.List(ComboBoxAgr.ListIndex)
                'Write to the "Aux" sheet the new grouping value
                Sheets("Aux").Cells(1, idxcomp + 2).Value = ComboBoxAgr.List(ComboBoxAgr.ListIndex)
            Else
                'Else Write "NA" to both the "Aux" sheet and the ListBox
                ListBoxC.List(idx, 2) = "NA"
                Sheets("Aux").Cells(1, idxcomp + 2).Value = "NA"
            End If
        
            'Rewrite the button caption and blank the input spaces
            AddComp.Caption = "Adicionar"
            ListBoxC.Enabled = True
            TxtComp.Value = ""
            TxtPesoC.Value = ""
            ListBoxC.ListIndex = -1
            CheckBoxAgr.Value = False
            ComboBoxAgr.ListIndex = -1
            ListBoxP.ListIndex = -1
            ListBoxC.ListIndex = -1
        Else
            'If we are adding a component
            idx = ListBoxC.ListCount
            componente = TxtComp.Value
            peso = TxtPesoC.Value
            ListBoxC.ColumnCount = 3
         
         
            'Add new component to ListBox and write it's name and weight values
            ListBoxC.AddItem
            ListBoxC.List(idx, 0) = componente
            ListBoxC.List(idx, 1) = peso
            
            'Write to "Aux" sheet the added component's name and weight values, using the global varibale AuxIndex
            Sheets("Aux").Cells(1, AuxIndex).Value = componente
            Sheets("Aux").Cells(1, AuxIndex + 1).Value = peso
            
            'If component has a grouping associated write it's grouping value to the "Aux" sheet and to the ListBox
            If CheckBoxAgr.Value = True And ComboBoxAgr.ListIndex <> -1 Then
                ListBoxC.List(idx, 2) = ComboBoxAgr.List(ComboBoxAgr.ListIndex)
                Sheets("Aux").Cells(1, AuxIndex + 2).Value = ComboBoxAgr.List(ComboBoxAgr.ListIndex)
            Else
            'Else, write "NA" to both the "Aux" sheet and the ListBox
                ListBoxC.List(idx, 2) = "NA"
                Sheets("Aux").Cells(1, AuxIndex + 2).Value = "NA"
            End If
             
            'Write the number of parameters this component has to "Aux" sheet (0 when created)
            Sheets("Aux").Cells(1, AuxIndex + 3).Value = 0
            
            'Increase the global variable and blank the input spaces
            AuxIndex = AuxIndex + 5
            TxtComp.Value = ""
            TxtPesoC.Value = ""
            CheckBoxAgr.Value = False
            ComboBoxAgr.ListIndex = -1
            
            ListBoxC.ListIndex = -1
            ListBoxP.ListIndex = -1
            
        End If
    End If
End Sub

Private Sub AddPar_Click()
    Dim parametro As String
    Dim old_parametro As String
    Dim peso As Double
    Dim comp As String
    Dim idxcomp As Integer
    Dim idxpar As Integer
    Dim n As Integer
    
    'If we got valid inputs
    If TxtPar.Value <> "" And TxtPesoP.Value <> "" And ListBoxC.ListIndex <> -1 Then
        parametro = TxtPar.Value
        peso = TxtPesoP.Value
        'Get the name of the component we're working on
        comp = ListBoxC.List(ListBoxC.ListIndex)
        'Find that component's position on the "Aux" sheet
        idxcomp = FindCompIndex(Sheets("Aux"), comp)
        
        If AddPar.Caption = "Corrigir" Then
            'Find parameter index and replace it's values in the "Aux" sheet
            old_parametro = ListBoxP.List(ListBoxP.ListIndex, 0)
            idxpar = FindParIndex(Sheets("Aux"), idxcomp, old_parametro)
            Sheets("Aux").Cells(idxpar, idxcomp).Value = parametro
            Sheets("Aux").Cells(idxpar, idxcomp + 1).Value = peso
            AddPar.Caption = "Adicionar"
        Else
            'Write the parameter components and increase the value of the cell with the number of parameters
            n = Sheets("Aux").Cells(1, idxcomp + 3).Value + 1
            Sheets("Aux").Cells(1 + n, idxcomp).Value = parametro
            Sheets("Aux").Cells(1 + n, idxcomp + 1).Value = peso
            Sheets("Aux").Cells(1, idxcomp + 3).Value = n
        End If
        'Force a change to ListBoxC to refresh ListBoxP
        ListBoxC.ListIndex = ListBoxC.ListIndex - 1
        ListBoxC.ListIndex = ListBoxC.ListIndex + 1
        'Blank the input spaces
        TxtPar.Value = ""
        TxtPesoP.Value = ""
        ListBoxP.ListIndex = -1
        ListBoxC.Enabled = True
        ListBoxP.Enabled = True
    End If
End Sub

Private Sub AgrCriar_Click()
    Dim idx As Integer
    Dim old_agrupamento As String
    Dim agrupamento As String
    Dim col As Integer
    Dim n_alunos As Integer
    Dim i As Integer
    
    'If we have valid inputs
    If TxtAgr.Value <> "" Then
        'If we are correcting a grouping
        If AgrCriar.Caption = "Corrigir" Then
            idx = ListBoxA.ListIndex
            old_agrupamento = ListBoxA.List(idx)
            agrupamento = TxtAgr.Value
            'Replace old grouping with new grouping in ComboBox and ListBox
            ListBoxA.List(idx) = agrupamento
            ComboBoxAgr.List(idx) = agrupamento
            'Get the grouping's position in the "Grupos" table
            col = FindGrouping(Sheets("Grupos"), old_agrupamento)
            'Replace the old group name with the new group name in the "Grupos" sheet
            Sheets("Grupos").Cells(1, col).Value = agrupamento
            'Correct the group names of all the components associated with this group
            If AuxIndex <> 1 Then
                Call Correct(ListBoxC, old_agrupamento, agrupamento)
            End If
            'Blank all the input boxes
            AgrCriar.Caption = "Criar"
            ListBoxA.Enabled = True
            TxtAgr.Value = ""
            ListBoxA.ListIndex = -1
        
        'If we are creating a grouping
        Else
            'Create the sheet "Grupos" if it doesn't exist and move it to the right of the sheet "Alunos"
            If CheckIfSheetExists("Grupos") <> True Then
                Sheets.Add.name = "Grupos"
                Sheets("Grupos").Move After:=Sheets("Alunos")
                Sheets("Grupos").Cells(1, 1).Value = "Estudantes"
                Sheets("Grupos").Cells(1, 1).Interior.ColorIndex = 15
                
                n_alunos = Sheets("Aux").Range("E1").Value
                
                For i = 2 To n_alunos + 1
                    Sheets("Grupos").Cells(i, 1).Value = Sheets("Alunos").Cells(i + 6, 1).Value
                Next i
                
            End If
            'Add the new grouping to the ListBox and the ComboBox
            n_alunos = Sheets("Aux").Range("E1").Value
            agrupamento = TxtAgr.Value
            ComboBoxAgr.AddItem (agrupamento)
            idx = ListBoxA.ListCount
            ListBoxA.ColumnCount = 1
            ListBoxA.AddItem (agrupamento)
            'Find the next available column to write to in the "Grupos" sheet
            col = Sheets("Grupos").Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).column
            'Write the new grouping to the "Grupos" sheet
            Sheets("Grupos").Cells(1, col).Value = agrupamento
            Sheets("Grupos").Cells(1, col).Interior.ColorIndex = 15
            Sheets("Grupos").Range(Cells(1, 1), Cells(n_alunos + 1, col)).Borders.LineStyle = xlContinuous
            Sheets("Grupos").Range(Cells(1, 1), Cells(n_alunos + 1, col)).HorizontalAlignment = xlCenter
            Sheets("Grupos").Range(Cells(1, 1), Cells(n_alunos + 1, col)).ColumnWidth = 15
            Sheets("Grupos").Range(Cells(1, 1), Cells(n_alunos + 1, col)).VerticalAlignment = xlVAlignCenter
            Sheets("Grupos").Range(Cells(1, 1), Cells(1, col)).WrapText = True
            'Blank input boxes
            TxtAgr.Value = ""
            ListBoxA.ListIndex = -1
        End If
    End If
End Sub

Private Sub EditarAgr_Click()
    Dim idx As Integer
    
    'If there are elements to be removed in the listbox and the listbox has one element selected
    If ListBoxA.ListCount <> 0 Then
         idx = ListBoxA.ListIndex
         If ListBoxA.ListIndex <> -1 Then
            'Change button label caption and place selected grouping in the textbox
            AgrCriar.Caption = "Corrigir"
            TxtAgr.Value = ListBoxA.List(idx)
            ListBoxA.Enabled = False
         End If
    End If
End Sub

Private Sub EditarComp_Click()
    Dim idx As Integer
    Dim grupo As String
    
    'If there are elements to be edited in the listbox and the listbox has one element selected
    If ListBoxC.ListCount <> 0 And ListBoxC.ListIndex <> -1 Then
        'Change button label and place selected component's values in the textboxes
         AddComp.Caption = "Corrigir"
         idx = ListBoxC.ListIndex
         TxtComp.Value = ListBoxC.List(idx, 0)
         TxtPesoC.Value = ListBoxC.List(idx, 1)
         'Diplay ComboBoxAgr and CheckBoxAgr based on component's grouping
         If ListBoxC.List(idx, 2) <> "NA" Then
            CheckBoxAgr.Value = True
            grupo = SelectItemComboBox(ComboBoxAgr, ListBoxC.List(idx, 2))
            ComboBoxAgr.ListIndex = grupo
         Else
            CheckBoxAgr.Value = False
            ComboBoxAgr.ListIndex = -1
         End If
            ListBoxC.Enabled = False
    End If
End Sub

Private Sub EditarPar_Click()
    Dim idx As Integer
    
    'If there are elements to be edited in the listbox and the listbox has one element selected
    If ListBoxC.ListCount <> 0 And ListBoxP.ListCount <> 0 And ListBoxC.ListIndex <> -1 And ListBoxP.ListIndex <> -1 Then
        'Change button label and place selected parameters's values in the textboxes
        AddPar.Caption = "Corrigir"
        idx = ListBoxP.ListIndex
        TxtPar.Value = ListBoxP.List(idx, 0)
        TxtPesoP.Value = ListBoxP.List(idx, 1)
        ListBoxC.Enabled = False
        ListBoxP.Enabled = False
    End If
End Sub

Private Sub Executar_Click()
    Dim idx As Integer
    'Disable Screen Updating to increase performance
    Application.ScreenUpdating = False
    'For each component do
    idx = 1
    Do While idx < AuxIndex
        Dim agrupamento As String
        Dim name As String
        Dim n_par As Integer
        Dim TempArray() As Variant
        Dim i As Integer
        Dim letter As String
        Dim n_alunos As Integer
        Dim col As Integer
        Dim cond1 As FormatCondition
        
        'get it's name, number of parameters, grouping and number of students
        name = Sheets("Aux").Cells(1, idx).Value
        agrupamento = Sheets("Aux").Cells(1, idx + 2).Value
        n_par = Sheets("Aux").Cells(1, idx + 3).Value
        n_alunos = Sheets("Aux").Range("E1").Value
        
        'if component has no grouping
        If agrupamento = "NA" Then
            
            'create a new sheet fot that component and add it to the end of the workbook
            Sheets.Add(After:=Sheets(Sheets.Count)).name = name
            
            'Fill an array with data acording to the component and write that array to the worksheet
            TempArray = CreateArray(idx, Sheets("Aux"), Sheets("Alunos"), n_alunos + 2, n_par + 2)
            Sheets(name).Range(Cells(1, 1), Cells(n_alunos + 2, n_par + 2)).Value = TempArray
            
            'Write the "Total" column formulas directly to the cell
            letter = Number2Letter(n_par + 1)
            Sheets(name).Cells(1, n_par + 2).formula = "=SUM(B1:" & letter & "1)"
            Sheets(name).Cells(2, n_par + 2).Value = "Total"
            
            For i = 3 To n_alunos + 2
                Sheets(name).Cells(i, n_par + 2).formula = "=SUM(B" & CStr(i) & ":" & letter & CStr(i) & ")"
            Next i
            
            Call UpdateAlunos(Sheets("Alunos"), n_alunos + 2, n_par + 2, Sheets("Aux").Cells(1, idx + 1).Value, name)
            
            'Format the cells with the right colors and text alignments
            Sheets(name).Range(Cells(2, 1), Cells(2, n_par + 2)).WrapText = True
            Sheets(name).Range(Cells(1, 1), Cells(n_alunos + 2, n_par + 2)).Borders.LineStyle = xlContinuous
            Sheets(name).Range(Cells(1, 1), Cells(n_alunos + 2, n_par + 2)).HorizontalAlignment = xlCenter
            Sheets(name).Range(Cells(1, 1), Cells(n_alunos + 2, n_par + 2)).VerticalAlignment = xlVAlignCenter
            Sheets(name).Range(Cells(2, 1), Cells(2, n_par + 2)).Interior.Color = RGB(211, 226, 235)
            Sheets(name).Range(Cells(3, 1), Cells(n_alunos + 2, 1)).Interior.Color = RGB(224, 224, 222)
            Sheets(name).Range(Cells(3, n_par + 2), Cells(n_alunos + 2, n_par + 2)).Interior.ColorIndex = 19
        Else
            'if component has agrouping
            Dim idxgroup As Integer
            Dim group_letter As String
            Dim n_grupos As Integer
            
            'Create a new sheet for that component and add it to the end of the workbook
            idxgroup = FindGrouping(Sheets("Grupos"), agrupamento)
            Sheets.Add(After:=Sheets(Sheets.Count)).name = name
            group_letter = Number2Letter(idxgroup)
            
            'Get the number of unique users in the grouping and write them to the component's sheet
            Sheets("Grupos").Range(group_letter & "1:" & group_letter & "65536").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets(name).Range("A2"), Unique:=True
            Sheets(name).Cells(2, 1).Value = "Grupos"
            n_grupos = Application.WorksheetFunction.CountA(Sheets(name).Range("A:A")) - 1
            
            'Fill an array with data acording to the component and write that array to the worksheet
            TempArray = CreateArrayGroup(idx, Sheets("Aux"), n_grupos + 2, n_par + 1)
            Sheets(name).Range(Cells(1, 2), Cells(n_grupos + 2, n_par + 2)).Value = TempArray
            
            'Write the "Total" column formulas directly to the cell
            letter = Number2Letter(n_par + 1)
            Sheets(name).Cells(1, n_par + 2).formula = "=SUM(B1:" & letter & "1)"
            Sheets(name).Cells(2, n_par + 2).Value = "Total"
            
            For i = 3 To n_grupos + 2
                Sheets(name).Cells(i, n_par + 2).formula = "=SUM(B" & CStr(i) & ":" & letter & CStr(i) & ")"
            Next i
            
            Call UpdateAlunosGroup(Sheets("Alunos"), Sheets("Grupos"), n_grupos + 2, n_par + 2, Sheets("Aux").Cells(1, idx + 1).Value, name, agrupamento, n_alunos)
            
            'Format the cells with the right colors and text alignment
            Sheets(name).Range(Cells(1, 1), Cells(n_grupos + 2, n_par + 2)).Borders.LineStyle = xlContinuous
            Sheets(name).Range(Cells(1, 1), Cells(n_grupos + 2, n_par + 2)).HorizontalAlignment = xlCenter
            Sheets(name).Range(Cells(1, 1), Cells(n_grupos + 2, n_par + 2)).VerticalAlignment = xlVAlignCenter
            Sheets(name).Range(Cells(2, 1), Cells(2, n_par + 2)).WrapText = True
            Sheets(name).Range(Cells(2, 1), Cells(2, n_par + 2)).Interior.Color = RGB(211, 226, 235)
            Sheets(name).Range(Cells(3, 1), Cells(n_grupos + 2, 1)).Interior.Color = RGB(224, 224, 222)
            Sheets(name).Range(Cells(3, n_par + 2), Cells(n_grupos + 2, n_par + 2)).Interior.ColorIndex = 19
            
        End If
        idx = idx + 5
    Loop
    
    'Get first empty column in sheet "Alunos"
    col = Sheets("Alunos").Cells(7, Columns.Count).End(xlToLeft).Offset(0, 1).column
    
    'Write the sum column to the "Alunos" sheet
    Sheets("Alunos").Cells(6, col).formula = "=SUM(B6:" & Number2Letter(col - 1) & "6)"
    Sheets("Alunos").Cells(7, col).Value = "Total"
    
    For i = 8 To 8 + n_alunos - 1
        Sheets("Alunos").Cells(i, col).formula = "=SUMIF(B" & CStr(i) & ":" & Number2Letter(col - 1) & CStr(i) & ","">=0"")"
    Next i
    
    'Write the Nota column to the "Alunos" sheet
    Sheets("Alunos").Cells(6, col + 1).formula = "=" & Number2Letter(col) & CStr(6)
    Sheets("Alunos").Cells(7, col + 1).Value = "Nota"
    
    'Conditional formating on the Nota column of "Alunos" sheet
    Sheets("Alunos").Select
    Set cond1 = Sheets("Alunos").Range(Cells(8, col + 1), Cells(n_alunos + 7, col + 1)).FormatConditions.Add(xlCellValue, xlLess, "=" & CStr(Sheets("Alunos").Cells(6, col + 1).Value) & "/2")
    With cond1
    .Interior.Color = RGB(230, 170, 177)
    .Font.Color = vbRed
    End With
    
    'Format the cells with the right colors and text alignment
    Sheets("Alunos").Range(Cells(7, 1), Cells(7, col + 1)).Interior.ColorIndex = 15
    Sheets("Alunos").Range(Cells(8, col), Cells(n_alunos + 7, col)).Interior.ColorIndex = 19
    Sheets("Alunos").Range(Cells(7, 1), Cells(n_alunos + 7, col + 1)).Borders.LineStyle = xlContinuous
    Sheets("Alunos").Cells(6, 1).Borders.LineStyle = xlLineStyleNone
    Sheets("Alunos").Range(Cells(6, 2), Cells(n_alunos + 7, col + 1)).HorizontalAlignment = xlCenter
    Sheets("Alunos").Range(Cells(8, 2), Cells(n_alunos + 7, col)).NumberFormat = "0.00"
    
    For i = 2 To col + 1
        If IsEmpty(Sheets("Alunos").Cells(6, i).Value) = False Then
            Sheets("Alunos").Cells(6, i).Borders.LineStyle = xlContinuous
        End If
    Next i
    
    'Create the "Síntese" sheet
    Sheets.Add(Before:=Sheets("Alunos")).name = "Síntese"
    
    'Write all the possible grades to the "Síntese" sheet
    Sheets("Síntese").Cells(1, 1).Value = "Notas"
    For i = 0 To Sheets("Alunos").Cells(6, col + 1).Value
        Sheets("Síntese").Cells(i + 2, 1).Value = i
    Next i
    
    'Write all frequencies to the grades that were previously added
    Sheets("Síntese").Cells(1, 2).Value = "Freq"
    For i = 2 To Sheets("Alunos").Cells(6, col + 1).Value + 2
        Sheets("Síntese").Cells(i, 2).formula = "=COUNTIF('Alunos'!" & Number2Letter(col + 1) & "8:" & Number2Letter(col + 1) & CStr(n_alunos + 7) & ",A" & CStr(i) & ")"
    Next i
    
    'Write final statistics'
    For i = 4 To 7
        Sheets("Síntese").Range(Cells(Sheets("Alunos").Cells(6, col + 1).Value + i, 1), Cells(Sheets("Alunos").Cells(6, col + 1).Value + i, 2)).Merge
    Next i
    
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 4, 1).Value = "Nota mais alta"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 4, 3).formula = "=MAX('Alunos'!" & Number2Letter(col + 1) & "8:" & Number2Letter(col + 1) & CStr(n_alunos + 7) & ")"
    
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 5, 1).Value = "Nota mais baixa"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 5, 3).formula = "=MIN('Alunos'!" & Number2Letter(col + 1) & "8:" & Number2Letter(col + 1) & CStr(n_alunos + 7) & ")"
    
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 6, 1).Value = "Média"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 6, 3).NumberFormat = "0.0"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 6, 3).formula = "=AVERAGE('Alunos'!" & Number2Letter(col + 1) & "8:" & Number2Letter(col + 1) & CStr(n_alunos + 7) & ")"
    
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 7, 1).Value = "#Aprov/#Avaliados"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 7, 3).NumberFormat = "0.00"
    Sheets("Síntese").Cells(Sheets("Alunos").Cells(6, col + 1).Value + 7, 3).formula = "=COUNTIF('Alunos'!" & Number2Letter(col + 1) & "8:" & Number2Letter(col + 1) & CStr(n_alunos + 7) & ", "">=" & CStr(Sheets("Alunos").Cells(6, col + 1).Value / 2) & """) / " & CStr(n_alunos)
    
    Sheets("Síntese").Range(Cells(Sheets("Alunos").Cells(6, col + 1).Value + 4, 1), Cells(Sheets("Alunos").Cells(6, col + 1).Value + 7, 3)).Borders.LineStyle = xlContinuous
    Application.ScreenUpdating = True
    
    'Create a bar chart with the grade distribution
    ActiveSheet.Shapes.AddChart(, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Síntese!$B$2:$B$" & CStr(Sheets("Alunos").Cells(6, col + 1).Value + 2))
    ActiveChart.SeriesCollection(1).XValues = Range("Síntese!$A$2:$A$" & CStr(Sheets("Alunos").Cells(6, col + 1).Value + 2))
    
    'Move graph to correct position and give it a title
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveSheet.Shapes("Gráfico 1").IncrementLeft 100.5
    ActiveSheet.Shapes("Gráfico 1").IncrementTop -53.25
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Distribuição das Notas"
    
    'Give labels to the axis
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Notas"
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Frequência"
    
    'Define the graph axis scale
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MajorUnit = 1
    ActiveChart.Axes(xlValue).MinorUnit = 1
End Sub

Private Sub ListBoxC_Change()
    Dim comp As String
    
    'If theres any component selected
    If ListBoxC.ListIndex <> -1 Then
        'Get that component's name
        comp = ListBoxC.List(ListBoxC.ListIndex, 0)
        'Fill ListBoxP with all it's parameters stored in the "Aux" sheet
         Call FillListBox(Sheets("Aux"), comp, ListBoxP)
    'If there is no component selected, clear the Listbox
    Else
        ListBoxP.Clear
    End If
End Sub

Private Sub UserForm_Initialize()
    'Disable ComboBoxAgr by default
    ComboBoxAgr.Enabled = False
End Sub

Private Sub CheckBoxAgr_Click()
    'If CheckBoxAgr is ticked then ComboBoxAgr is enabled
    ComboBoxAgr.Enabled = CheckBoxAgr.Value
End Sub

