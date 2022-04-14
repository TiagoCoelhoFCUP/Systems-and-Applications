VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'Declare the recordsets to be used during the form execution
Public rst_pessoas As New ADODB.Recordset
Public rst_pais As New ADODB.Recordset
Public rst_irmãos As New ADODB.Recordset
Public rst_casamentos As New ADODB.Recordset
Public pai_id As Integer
Public mae_id As Integer

'Method to add a new person to the database
Private Sub AddPessoa_Click()
    Dim dirs As Variant
    Dim filename As String
    'If there's no other person selected and the mandatory input fields are all filled
     If Me!DropDownSelector.ListIndex = -1 And Me!TxtNome.Value <> "" And (Me!RadioMasc.Value Or Me!RadioFem.Value) And Me!TxtDataN.Value <> "" And Me!TxtNac.Value <> "" Then
        'Create a new entry on the "Pessoas" table
        rst_pessoas.AddNew
        'Fill in the respective fields values
        rst_pessoas.Fields("Nome") = Me!TxtNome.Value
        If Me!RadioMasc.Value = True Then
            rst_pessoas.Fields("Genero") = "Masculino"
        Else
            rst_pessoas.Fields("Genero") = "Feminino"
        End If
        
        rst_pessoas.Fields("DataNasc") = Me!TxtDataN.Value
        rst_pessoas.Fields("DataObito") = Me!TxtDataO.Value
        rst_pessoas.Fields("Nacionalidade") = Me!TxtNac.Value
        rst_pessoas.Fields("Residencia") = Me!TxtRes.Value
        rst_pessoas.Fields("Profissao") = Me!TxtProf.Value
        
        'If no picture is selected the "default" picture is selected
        If Me!Imagem.Picture = "(nenhum)" Then
            rst_pessoas.Fields("Foto") = "default.png"
        Else
            dirs = Split(Me!Imagem.Picture, "\", , vbTextCompare)
            filename = dirs(UBound(dirs))
            rst_pessoas.Fields("Foto") = filename
        End If
        'Update the table/recordset
        rst_pessoas.Update
        
        'Add the new person to the ComboBox at the top of the form
        Me!DropDownSelector.AddItem (rst_pessoas.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
        
        'Clear form
        Call LimparForm_Click
    'If conditions are not met, display message box
    Else
        MsgBox "Para criar uma pessoa, não podes ter nenhuma pessoa selecionada e tens de inserir os 4 campos obrigatorios: Nome, Género, Data de Nascimento e Nacionalidade"
    End If
End Sub

'Method to add a relation
Private Sub AddRelacao_Click()
    'Verify that there's 2 people selected and a relation selected
    If Me!DropDownSelector.ListIndex <> -1 And Me!DropDownRelacao.ListIndex <> -1 And Me!DropDownPessoa.ListIndex <> -1 Then
        Dim Value As Integer
        Dim genero1 As String
        Dim genero2 As String
        Dim genero3 As String
        
        'If the relation is "Irmão de" then add a new entry to the "Irmão_de" table
        If Me!DropDownRelacao.Value = "Irmão de" Then
            rst_irmãos.AddNew
            rst_irmãos.Fields("Id_P1") = Me!DropDownSelector.Value
            rst_irmãos.Fields("Id_P2") = Me!DropDownPessoa.Value
            rst_irmãos.Update
            
        'If the reltion is "Pai de" then we need to make sure that you can create a new entry
        ElseIf Me!DropDownRelacao.Value = "Pai de" Then
            'Find the gender of the parent in the relation
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & Me!DropDownSelector.Value, 0, adSearchForward
            genero1 = rst_pessoas.Fields("Genero")
            genero2 = "NA"
            genero3 = "NA"
            
            'If the "Pai_de" table is empty
            If rst_pais.EOF And rst_pais.BOF Then
                'Add a new entry to that table with the selected relation
                rst_pais.AddNew
                rst_pais.Fields("IdPai") = Me!DropDownSelector.Value
                rst_pais.Fields("IdFilho") = Me!DropDownPessoa.Value
                rst_pais.Update
            'If Not, find the first entry on the "Pai_de" table for the specific son
            Else
                rst_pais.MoveFirst
                rst_pais.Find "IdFilho=" & Me!DropDownPessoa.Value, 0, adSearchForward
                'If we find a match, get that parent's gender
                If Not rst_pais.EOF Then
                    rst_pessoas.MoveFirst
                    rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
                    genero2 = rst_pessoas.Fields("Genero")
                    'Find the next entry, if it exists
                    rst_pais.MoveNext
                    rst_pais.Find "IdFilho=" & Me!DropDownPessoa.Value, 0, adSearchForward
                    'If it does exist, get that second parent's gender
                    If Not rst_pais.EOF Then
                        rst_pessoas.MoveFirst
                        rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
                        genero3 = rst_pessoas.Fields("Genero")
                    End If
                End If
                
                'If there are no parents with the same gender of the new parent
                If (genero1 <> genero2 And genero1 <> genero3) Then
                    'Add a new entry to the "Pai_de" table with the selected relation
                    rst_pais.AddNew
                    rst_pais.Fields("IdPai") = Me!DropDownSelector.Value
                    rst_pais.Fields("IdFilho") = Me!DropDownPessoa.Value
                    rst_pais.Update
                'If there is already a parent of that gender, display a message box
                Else
                    MsgBox "O Filho desta relação já possui um progenitor do mesmo género. Remova a relação anterior para conseguir adicionar uma nova"
                End If
            End If
            
        'If the relation is "Filho then we need to make sure that you can create a new entry
        ElseIf Me!DropDownRelacao.Value = "Filho de" Then
            'Find the gender of the parent in the relation
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & Me!DropDownPessoa.Value, 0, adSearchForward
            genero1 = rst_pessoas.Fields("Genero")
            genero2 = "NA"
            genero3 = "NA"
            
            'If the "Pai_de" table is empty
            If rst_pais.EOF And rst_pais.BOF Then
                'Add a new entry to that table with the selected relation
                rst_pais.AddNew
                rst_pais.Fields("IdPai") = Me!DropDownPessoa.Value
                rst_pais.Fields("IdFilho") = Me!DropDownSelector.Value
                rst_pais.Update
            'If Not, find the first entry on the "Pai_de" table for the specific son
            Else
                rst_pais.MoveFirst
                rst_pais.Find "IdFilho=" & Me!DropDownSelector.Value, 0, adSearchForward
                'If we find a match, get that parent's gender
                If Not rst_pais.EOF Then
                    rst_pessoas.MoveFirst
                    rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
                    genero2 = rst_pessoas.Fields("Genero")
                    'Find the next entry, if it exists
                    rst_pais.MoveNext
                    rst_pais.Find "IdFilho=" & Me!DropDownSelector.Value, 0, adSearchForward
                    'If it does exist, get that second parent's gender
                    If Not rst_pais.EOF Then
                        rst_pessoas.MoveFirst
                        rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
                        genero3 = rst_pessoas.Fields("Genero")
                    End If
                End If
                
                'If there are no parents with the same gender of the new parent
                If (genero1 <> genero2 And genero1 <> genero3) Then
                    rst_pais.AddNew
                    rst_pais.Fields("IdPai") = Me!DropDownPessoa.Value
                    rst_pais.Fields("IdFilho") = Me!DropDownSelector.Value
                    rst_pais.Update
                'If there is already a parent of that gender, display a message box
                Else
                    MsgBox "O Filho desta relação já possui um progenitor do mesmo género. Remova a relação anterior para conseguir adicionar uma nova"
                End If
            End If
            
        'If the relation is "Casado com" then add a new entry to the "Casado_com" table
        ElseIf Me!DropDownRelacao.Value = "Casado com" Then
            'If provided a wedding date
            If TxtDataC.Value <> "" Then
                rst_casamentos.AddNew
                rst_casamentos.Fields("Id_P1") = Me!DropDownPessoa.Value
                rst_casamentos.Fields("Id_P2") = Me!DropDownSelector.Value
                rst_casamentos.Fields("DataInicio") = TxtDataC.Value
                rst_casamentos.Fields("DataFim") = TxtDataD.Value
                rst_casamentos.Update
            'Id there's no wedding date, display a message box
            Else
                MsgBox "Para adicionar a relação precisas de preencher o campo obrigatório 'Data de Casamento'"
            End If
        End If
        
        'Call the "DropDownSelector_Change" method to update the displayed relations
        Call DropDownSelector_Change
    'If there aren't at least 2 people selected and a relation selected, display a message box
    Else
        MsgBox "Para adicionar uma relação precias de ter o tipo de relação selecionado e as duas pessoas que participam na mesma"
    End If
End Sub

'Method to change the photo displayed at the centre of the form
Private Sub AlterarFoto_Click()
    Dim lngCount As Long
    Dim dirs As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        strPath = .SelectedItems(1)
    End With
    
    'get the filename from the fullpath of the selected photo
    dirs = Split(strPath, "\", , vbTextCompare)
    strFileName = dirs(UBound(dirs))
        
    'copy the file to the "Fotos" folder within the project's directory
    FileCopy strPath, CurrentProject.Path & "\Fotos\" & strFileName
    'Set the centre image as the selected image
    Imagem.Picture = CurrentProject.Path & "\Fotos\" & strFileName
End Sub

'Method to update a selected person's information
Private Sub AtualizarPessoa_Click()
    Dim dirs As Variant
    Dim filename As String
    Dim selecionado As Integer
    
    'If there's a person selected and the mandatory fields are filled
    If Me!DropDownSelector.ListIndex <> -1 And Me!TxtNome.Value <> "" And (Me!RadioMasc.Value Or Me!RadioFem.Value) And Me!TxtDataN.Value <> "" And Me!TxtNac.Value <> "" Then
        selecionado = Me!DropDownSelector.Value
        'Find the selected person in the "Pessoas" table
        rst_pessoas.MoveFirst
        rst_pessoas.Find "Id='" & selecionado & "'", 0, adSearchForward
        'Replace the field values with the values in the input boxes
        rst_pessoas.Fields("Nome") = Me!TxtNome.Value
        If Me!RadioMasc.Value = True Then
            rst_pessoas.Fields("Genero") = "Masculino"
        Else
            rst_pessoas.Fields("Genero") = "Feminino"
        End If
        
        rst_pessoas.Fields("DataNasc") = Me!TxtDataN.Value
        rst_pessoas.Fields("DataObito") = Me!TxtDataO.Value
        rst_pessoas.Fields("Nacionalidade") = Me!TxtNac.Value
        rst_pessoas.Fields("Residencia") = Me!TxtRes.Value
        rst_pessoas.Fields("Profissao") = Me!TxtProf.Value
        
        dirs = Split(Me!Imagem.Picture, "\", , vbTextCompare)
        filename = dirs(UBound(dirs))
        rst_pessoas.Fields("Foto") = filename
        rst_pessoas.Update
        
        'Replace the entry on the Combobox with the new entry
        Me!DropDownSelector.AddItem Item:=rst_pessoas.Fields("Id") & ";" & rst_pessoas.Fields("Nome"), Index:=Me!DropDownSelector.ListIndex
        Me!DropDownSelector.RemoveItem (Me!DropDownSelector.ListIndex + 1)
        
        Call LimparForm_Click
    'If conditions are not met, display a message box
    Else
        MsgBox "Para atualizar uma pessoa precisas de ter uma pessoa selecionada e não podes deixar em branco qualquer um dos 4 campos obrigatórios: Nome, Género, Data de Nascimento e Nacionalidade"
    End If
End Sub

'Method to de-select all other relations when "CheckBoxMãe" is selected
Private Sub CheckBoxMãe_Click()
    If CheckBoxMãe.Value = True Then
        ListIrmaos.Value = ""
        ListCasamentos.Value = ""
        ListFilhos.Value = ""
        CheckBoxPai.Value = False
    End If
End Sub

'Method to de-select all other relations when "CheckBoxPai" is selected
Private Sub CheckBoxPai_Click()
    If CheckBoxPai.Value = True Then
        ListIrmaos.Value = ""
        ListCasamentos.Value = ""
        ListFilhos.Value = ""
        CheckBoxMãe.Value = False
    End If
End Sub

'Method to delete a selected relation
Private Sub DelRelacao_Click()
    Dim selecionado As Integer
    'If there's a person selected
    If DropDownSelector.ListIndex <> -1 Then
        selecionado = Me!DropDownSelector.Value
        'If the selected relation is "Pai de"
        If CheckBoxPai.Value = True Then
            'Find that relation
            rst_pais.MoveFirst
            rst_pais.Find "Id=" & pai_id, 0, adSearchForward
            rst_pais.Delete
            'Delete it
            rst_pais.Close
            rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            pai_id = 0
            CheckBoxPai.Value = False
            TxtPai = ""
        'If the selected relation is "Mãe de"
        ElseIf CheckBoxMãe.Value = True Then
            'Find that relation
            rst_pais.MoveFirst
            rst_pais.Find "Id=" & mae_id, 0, adSearchForward
            'Delete it
            rst_pais.Delete
            rst_pais.Close
            rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            mae_id = 0
            CheckBoxMãe.Value = False
            TxtMae = ""
        'If the selected relation is "Irmão de"
        ElseIf ListIrmaos.ListIndex <> -1 Then
            'Find that relation
            rst_irmãos.MoveFirst
            rst_irmãos.Find "Id=" & ListIrmaos.Value, 0, adSearchForward
            'Delete it
            rst_irmãos.Delete
            rst_irmãos.Close
            rst_irmãos.Open "Irmão_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            ListIrmaos.RemoveItem (ListIrmaos.ListIndex)
        'If the selected relation is "Filho de"
        ElseIf ListFilhos.ListIndex <> -1 Then
            'Find that relation
            rst_pais.MoveFirst
            rst_pais.Find "Id=" & ListFilhos.Value, 0, adSearchForward
            'Delete it
            rst_pais.Delete
            rst_pais.Close
            rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            ListFilhos.RemoveItem (ListFilhos.ListIndex)
        'If the selected relation is "Casado com"
        ElseIf ListCasamentos.ListIndex <> -1 Then
            'Find that relation
            rst_casamentos.MoveFirst
            rst_casamentos.Find "Id_Casamento=" & ListCasamentos.Value, 0, adSearchForward
            'Delete it
            rst_casamentos.Delete
            rst_casamentos.Close
            rst_casamentos.Open "Casado_com", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            ListCasamentos.RemoveItem (ListCasamentos.ListIndex)
        End If
    'If not person is selected, display a message box
    Else
        MsgBox "Para eliminar uma relação precisas de ter uma pessoa e uma relação selecionadas"
    End If
End Sub

'Method that disables the data fields for a relation when they are not needed
Private Sub DropDownRelacao_Change()
    If Me!DropDownRelacao.Value = "Casado com" Then
        Me!TxtDataC.Value = Null
        Me!TxtDataD.Value = Null
        Me!TxtDataC.Enabled = True
        Me!TxtDataD.Enabled = True
    Else
        Me!TxtDataC.Enabled = False
        Me!TxtDataD.Enabled = False
    End If
End Sub

'Method that opens the recordsets and fills the comboboxes when the form is loaded
Private Sub Form_Load()
    'Set some inputs as locked/disabled
    Me!TxtPai.Locked = True
    Me!TxtMae.Locked = True
    Me!TxtDataC.Enabled = False
    Me!TxtDataD.Enabled = False
    Me!DropDownRelacao.Enabled = False
    Me!DropDownPessoa.Enabled = False
    
    'Add relation types to "DropDownRelacao"
    Me!DropDownRelacao.AddItem ("Pai de")
    Me!DropDownRelacao.AddItem ("Filho de")
    Me!DropDownRelacao.AddItem ("Irmão de")
    Me!DropDownRelacao.AddItem ("Casado com")
    
    'Open the connections to the tables in the database
    rst_pessoas.Open "Pessoas", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    rst_casamentos.Open "Casado_com", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    rst_irmãos.Open "Irmão_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    
    'Load every person to the "DropDownSelector" combobox
    Me!DropDownSelector.ColumnCount = 2
    Do While Not rst_pessoas.EOF
        Me!DropDownSelector.AddItem (rst_pessoas.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
        If IsNull(rst_pessoas.Fields("Foto").Value) Then
            rst_pessoas.Fields("Foto").Value = "default.png"
            rst_pessoas.Update
        End If
        
        rst_pessoas.MoveNext
    Loop
End Sub

'Method that displays the information regarding a selected person
Private Sub DropDownSelector_Change()

    'Clear some inputs
    ListFilhos.Value = ""
    ListIrmaos.Value = ""
    ListCasamentos.Value = ""
    CheckBoxPai.Value = False
    CheckBoxMãe.Value = False
    
    Dim selecionado As Integer
    selecionado = Me!DropDownSelector.Value
    'Find selected person on the "Pessoas" table
    rst_pessoas.MoveFirst
    rst_pessoas.Find "Id=" & selecionado, 0, adSearchForward
    'Fill the input boxes with the person's information
    Me!TxtNome.Value = rst_pessoas.Fields("Nome")
    If rst_pessoas.Fields("Genero") = "Masculino" Then
        Me!RadioMasc.Value = True
        Me!RadioFem.Value = False
    Else
        Me!RadioFem.Value = True
        Me!RadioMasc.Value = False
    End If
    Me!TxtDataN.Value = rst_pessoas.Fields("DataNasc")
    Me!TxtDataO.Value = rst_pessoas.Fields("DataObito")
    Me!TxtNac.Value = rst_pessoas.Fields("Nacionalidade")
    Me!TxtRes.Value = rst_pessoas.Fields("Residencia")
    Me!TxtProf.Value = rst_pessoas.Fields("Profissao")
    Me!Imagem.Picture = CurrentProject.Path & "\Fotos\" & rst_pessoas.Fields("Foto").Value
    
    'If the "Pai_de" table is not empty
    If Not (rst_pais.EOF And rst_pais.BOF) Then
        'Find the first entry that has the selected person as the son in the relation
        rst_pais.MoveFirst
        rst_pais.Find "IdFilho=" & selecionado, 0, adSearchForward
        'If there was a match
        If Not rst_pais.EOF Then
            'Find the parent's corresponding entry in the "Pessoas" table
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
            'If his gender is male, fill "TxtPai" with his name
            If rst_pessoas.Fields("Genero") = "Masculino" Then
                pai_id = rst_pais.Fields("Id")
                Me!TxtPai = rst_pessoas.Fields("Nome")
                Me!TxtMae = ""
                Me!CheckBoxPai.Enabled = True
            'Else, fill "TxtMãe" with her name
            Else
                mae_id = rst_pais.Fields("Id")
                Me!TxtMae = rst_pessoas.Fields("Nome")
                Me!TxtPai = ""
                Me!CheckBoxMãe.Enabled = True
            End If
            
            'Move the cursor to the next entry and search for a new relation
            rst_pais.MoveNext
            rst_pais.Find "IdFilho=" & selecionado, 0, adSearchForward
            'If there was a match
            If Not rst_pais.EOF Then
                'Find the parent's corresponding entry in the "Pessoas" table
                rst_pessoas.MoveFirst
                rst_pessoas.Find "Id=" & rst_pais.Fields("IdPai"), 0, adSearchForward
                'If his gender is male, fill "TxtPai" with his name
                If rst_pessoas.Fields("Genero") = "Masculino" Then
                    pai_id = rst_pais.Fields("Id")
                    Me!TxtPai.Value = rst_pessoas.Fields("Nome")
                    Me!CheckBoxPai.Enabled = True
                'Else, fill "TxtMãe" with her name
                Else
                    mae_id = rst_pais.Fields("Id")
                    Me!TxtMae.Value = rst_pessoas.Fields("Nome")
                    Me!CheckBoxMãe.Enabled = True
                End If
            'If the second entry was not found
            Else
                'If the entry that was found was the male parent
                If TxtPai.Value <> "" Then
                    'Clear the female parent's info
                    Me!TxtMae.Value = ""
                    mae_id = 0
                    Me!CheckBoxMãe.Enabled = False
                'If the entry found was the female parent
                Else
                    'Clear the male parent's info
                    Me!TxtPai.Value = ""
                    pai_id = 0
                    Me!CheckBoxPai.Enabled = False
                End If
            End If
        'If there were no matches found
        Else
            'Clear both parent's info
            Me!TxtPai.Value = ""
            Me!TxtMae.Value = ""
            mae_id = 0
            pai_id = 0
            Me!CheckBoxPai.Enabled = False
            Me!CheckBoxMãe.Enabled = False
        End If
        rst_pais.MoveFirst
    End If
    
    'Clear the previous Listbox
    ListFilhos.rowsource = ""
    
    'If the "Pai_de" table is not empty
    If Not (rst_pais.EOF And rst_pais.BOF) Then
        'Find the first entry that has the selected person as the their parent in the relation
        rst_pais.MoveFirst
        rst_pais.Find "IdPai=" & selecionado, 0, adSearchForward
        'While we keep finding matches
        While Not rst_pais.EOF
            'Find the entry on the "Pessoas" table that matches with the son of the relationship
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_pais.Fields("IdFilho"), 0, adSearchForward
            'Add that son to the "ListFilhos" listbox
            ListFilhos.AddItem (rst_pais.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
            rst_pais.MoveNext
            rst_pais.Find "IdPai=" & selecionado, 0, adSearchForward
        Wend
        rst_pais.MoveFirst
    End If
    
    'Clear the previous Listbox
    ListIrmaos.rowsource = ""
    
    'If the 'Irmão_de' table is not empty
    If Not (rst_irmãos.EOF And rst_irmãos.BOF) Then
        'Find the first entry that has the selected person's Id in the "P1_Id" field on the "Irmão_de" table
        rst_irmãos.MoveFirst
        rst_irmãos.Find "Id_P1=" & selecionado, 0, adSearchForward
        'While we keep finding matches
        While Not rst_irmãos.EOF
            'Find the entry of the "Pessoas" table that matches witht the brother
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_irmãos.Fields("Id_P2"), 0, adSearchForward
            'Add that brother to the "ListIrmaos" listbox
            ListIrmaos.AddItem (rst_irmãos.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
            rst_irmãos.MoveNext
            rst_irmãos.Find "Id_P1=" & selecionado, 0, adSearchForward
        Wend
        
        'Find the first entry that has the selected person's Id in the "P2_Id" field on the "Irmão_de" table
        rst_irmãos.MoveFirst
        rst_irmãos.Find "Id_P2=" & selecionado, 0, adSearchForward
        'While we keep finding matches
        While Not rst_irmãos.EOF
            'Find the entry of the "Pessoas" table that matches witht the brother
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_irmãos.Fields("Id_P1"), 0, adSearchForward
            'Add that brother to the "ListIrmaos" listbox
            ListIrmaos.AddItem (rst_irmãos.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
            rst_irmãos.MoveNext
            rst_irmãos.Find "Id_P2=" & selecionado, 0, adSearchForward
        Wend
        rst_irmãos.MoveFirst
    End If
    
    'Clear the previous Listbox
    ListCasamentos.rowsource = ""
    
    'If the "Casado_com" table is not empty
    If Not (rst_casamentos.EOF And rst_casamentos.BOF) Then
        'Find the first entry that has the selected person's Id in the "P1_Id" field on the "Casado_com" table
        rst_casamentos.MoveFirst
        rst_casamentos.Find "Id_P1=" & selecionado, 0, adSearchForward
        'While we keep finding matches"
        While Not rst_casamentos.EOF
            'Find the entry of the "Pessoas" table that matches witht the spouse
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_casamentos.Fields("Id_P2"), 0, adSearchForward
            'Add that marriage to the "ListCasamentos" listbbox
            ListCasamentos.AddItem (rst_casamentos.Fields("Id_Casamento") & ";" & rst_pessoas.Fields("Nome") & ";" & rst_casamentos.Fields("DataInicio") & ";" & rst_casamentos.Fields("DataFim"))
            rst_casamentos.MoveNext
            rst_casamentos.Find "Id_P1=" & selecionado, 0, adSearchForward
        Wend
        
        'Find the first entry that has the selected person's Id in the "P2_Id" field on the "Casado_com" table
        rst_casamentos.MoveFirst
        rst_casamentos.Find "Id_P2=" & selecionado, 0, adSearchForward
        'While we keep finding matches"
        While Not rst_casamentos.EOF
            'Find the entry of the "Pessoas" table that matches witht the spouse
            rst_pessoas.MoveFirst
            rst_pessoas.Find "Id=" & rst_casamentos.Fields("Id_P1"), 0, adSearchForward
            'Add that marriage to the "ListCasamentos" listbbox
            ListCasamentos.AddItem (rst_casamentos.Fields("Id_Casamento") & ";" & rst_pessoas.Fields("Nome") & ";" & rst_casamentos.Fields("DataInicio") & ";" & rst_casamentos.Fields("DataFim"))
            rst_casamentos.MoveNext
            rst_casamentos.Find "Id_P2=" & selecionado, 0, adSearchForward
        Wend
        rst_casamentos.MoveFirst
    End If
    
    'Enable the inputs related to adding a new relation
    Me!DropDownRelacao.Enabled = True
    Me!DropDownPessoa.Enabled = True
    Me!TxtDataC.Enabled = False
    Me!TxtDataD.Enabled = False
    Me!DropDownRelacao.Value = ""
    Me!DropDownPessoa.Value = ""
    
    'Populate the "DropDownPessoa" combobox with all the persons the selected individual can form relations with
    Me!DropDownPessoa.ColumnCount = 2
    Me!DropDownPessoa.rowsource = ""
    rst_pessoas.MoveFirst
    Do While Not rst_pessoas.EOF
        If rst_pessoas.Fields("Id") <> selecionado Then
            Me!DropDownPessoa.AddItem (rst_pessoas.Fields("Id") & ";" & rst_pessoas.Fields("Nome"))
        End If
        rst_pessoas.MoveNext
    Loop
End Sub

'Method that clears the form
Private Sub LimparForm_Click()
    'set "mae_id" and "pai_id" to 0
    mae_id = 0
    pai_id = 0
    'set the resultsets to the first positions
    If Not (rst_pessoas.EOF And rst_pessoas.BOF) Then
        rst_pessoas.MoveFirst
    End If
    If Not (rst_pais.EOF And rst_pais.BOF) Then
        rst_pais.MoveFirst
    End If
    If Not (rst_irmãos.EOF And rst_irmãos.BOF) Then
        rst_irmãos.MoveFirst
    End If
    If Not (rst_casamentos.EOF And rst_casamentos.BOF) Then
        rst_casamentos.MoveFirst
    End If
    
    'Clear the rest of the inputs
    Me!DropDownSelector.Value = ""
    Me!TxtNome.Value = ""
    Me!RadioMasc.Value = False
    Me!RadioFem.Value = False
    Me!TxtDataN.Value = Null
    Me!TxtDataO.Value = Null
    Me!TxtNac.Value = ""
    Me!TxtRes.Value = ""
    Me!TxtProf.Value = ""
    Me!Imagem.Picture = "(nenhum)"
    Me!TxtPai.Value = ""
    Me!TxtMae.Value = ""
    Me!ListFilhos.rowsource = ""
    Me!ListIrmaos.rowsource = ""
    Me!ListCasamentos.rowsource = ""
    Me!DropDownRelacao.Value = ""
    Me!DropDownRelacao.Enabled = False
    Me!DropDownPessoa.Value = ""
    Me!DropDownPessoa.Enabled = False
    Me!TxtDataC.Value = Null
    Me!TxtDataD.Value = Null
    Me!TxtDataC.Enabled = False
    Me!TxtDataD.Enabled = False
    Me!CheckBoxMãe.Value = False
    Me!CheckBoxPai.Value = False
    Me!CheckBoxPai.Enabled = False
    Me!CheckBoxMãe.Enabled = False
End Sub

'Method to de-select all other relations when a "ListFilhos" entry is selected
Private Sub ListFilhos_Click()
    If ListFilhos.ListIndex <> -1 Then
        ListIrmaos.Value = ""
        ListCasamentos.Value = ""
        CheckBoxPai.Value = False
        CheckBoxMãe.Value = False
    End If
End Sub

'Method to de-select all other relations when a "ListCasamentos" entry is selected
Private Sub ListCasamentos_Click()
    If ListCasamentos.ListIndex <> -1 Then
        ListIrmaos.Value = ""
        ListFilhos.Value = ""
        CheckBoxPai.Value = False
        CheckBoxMãe.Value = False
    End If
End Sub

'Method to de-select all other relations when a "ListIrmaos" entry is selected
Private Sub ListIrmaos_Click()
    If ListIrmaos.ListIndex <> -1 Then
        ListFilhos.Value = ""
        ListCasamentos.Value = ""
        CheckBoxPai.Value = False
        CheckBoxMãe.Value = False
    End If
End Sub

'Method that "moves" the form to the previous person
Private Sub PessoaAnterior_Click()
    'If there's a person selected
    If Me!DropDownSelector.ListIndex <> -1 Then
        Dim selecionado As Integer
        selecionado = Me!DropDownSelector.Value
        rst_pessoas.MoveFirst
        rst_pessoas.Find "Id=" & selecionado, 0, adSearchForward
        'set the recordset to the previous person
        rst_pessoas.MovePrevious
    'If there's not
    Else
        'set the recordset to the last person
        rst_pessoas.MoveLast
    End If
    
    'If the recordset is not at BOF
    If Not rst_pessoas.BOF Then
        'Set the "DropDownSelector" value as the current person's Id
        Me!DropDownSelector.Value = rst_pessoas.Fields("Id")
        'Call "DropDownSelector_Change"
        Call DropDownSelector_Change
    End If
End Sub

'Method that "moves" the form to the next person
Private Sub PessoaSeguinte_Click()
    'If there's a person selected
    If Me!DropDownSelector.ListIndex <> -1 Then
        Dim selecionado As Integer
        selecionado = Me!DropDownSelector.Value
        rst_pessoas.MoveFirst
        rst_pessoas.Find "Id=" & selecionado, 0, adSearchForward
        'set the recordset to the next person
        rst_pessoas.MoveNext
    'If there's not
    Else
        'set the recordset to the first person
        rst_pessoas.MoveFirst
    End If
    
    'If the recordset is not at EOF
    If Not rst_pessoas.EOF Then
        'Set the "DropDownSelector" value as the current person's Id
        Me!DropDownSelector.Value = rst_pessoas.Fields("Id")
        'Call "DropDownSelector_Change"
        Call DropDownSelector_Change
    End If
End Sub

'Method that de-selects the female radio button when the male is selected
Private Sub RadioFem_Click()
    RadioFem.Value = True
    RadioMasc.Value = False
End Sub

'Method that de-selects the male radio buttin when the female is selected
Private Sub RadioMasc_Click()
    RadioMasc.Value = True
    RadioFem.Value = False
End Sub

'Method that removes a person form the database
Private Sub RemoverPessoa_Click()
    'If there's a person selected
    If Me!DropDownSelector.ListIndex <> -1 Then
        Dim selecionado As Integer
        selecionado = Me!DropDownSelector.Value
        
        'Remove it's entries in the "Pai_de" table where he/she is the son
        If Not (rst_pais.EOF And rst_pais.BOF) Then
            rst_pais.MoveFirst
            While Not rst_pais.EOF
                rst_pais.Find "IdFilho=" & selecionado, 0, adSearchForward
                If Not rst_pais.EOF Then
                    rst_pais.Delete
                    rst_pais.Close
                    rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If

        
        'Remove it's entries in the "Pai_de" table where he/she is the parent
        If Not (rst_pais.EOF And rst_pais.BOF) Then
            rst_pais.MoveFirst
            While Not rst_pais.EOF
                rst_pais.Find "IdPai=" & selecionado, 0, adSearchForward
                If Not rst_pais.EOF Then
                    rst_pais.Delete
                    rst_pais.Close
                    rst_pais.Open "Pai_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If
        
        'Remove the entries in the "Irmão_de" table where the person's Id in the the "P1_Id" field
        If Not (rst_irmãos.EOF And rst_irmãos.BOF) Then
            rst_irmãos.MoveFirst
            While Not rst_irmãos.EOF
                rst_irmãos.Find "Id_P1=" & selecionado, 0, adSearchForward
                If Not rst_irmãos.EOF Then
                    rst_irmãos.Delete
                    rst_irmãos.Close
                    rst_irmãos.Open "Irmão_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If
        
        'Remove the entries in the "Irmão_de" table where the person's Id in the the "P2_Id" field
        If Not (rst_irmãos.EOF And rst_irmãos.BOF) Then
            rst_irmãos.MoveFirst
            While Not rst_irmãos.EOF
                rst_irmãos.Find "Id_P2=" & selecionado, 0, adSearchForward
                If Not rst_irmãos.EOF Then
                    rst_irmãos.Delete
                    rst_irmãos.Close
                    rst_irmãos.Open "Irmão_de", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If
        
        'Remove the entries in the "Casado_com" table where the person's Id in the the "P1_Id" field
        If Not (rst_casamentos.EOF And rst_casamentos.BOF) Then
            rst_casamentos.MoveFirst
            While Not rst_casamentos.EOF
                rst_casamentos.Find "Id_P1=" & selecionado, 0, adSearchForward
                If Not rst_casamentos.EOF Then
                    rst_casamentos.Delete
                    rst_casamentos.Close
                    rst_casamentos.Open "Casado_com", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If
        
        'Remove the entries in the "Casado_com" table where the person's Id in the the "P2_Id" field
        If Not (rst_casamentos.EOF And rst_casamentos.BOF) Then
            rst_casamentos.MoveFirst
            While Not rst_casamentos.EOF
                rst_casamentos.Find "Id_P2=" & selecionado, 0, adSearchForward
                If Not rst_casamentos.EOF Then
                    rst_casamentos.Delete
                    rst_casamentos.Close
                    rst_casamentos.Open "Casado_com", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
                End If
            Wend
        End If
        
        'Remove person entry from the "Pessoas" table
        rst_pessoas.MoveFirst
        rst_pessoas.Find "Id=" & selecionado, 0, adSearchForward
        rst_pessoas.Delete
        rst_pessoas.Close
        rst_pessoas.Open "Pessoas", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        Me!DropDownSelector.RemoveItem (Me!DropDownSelector.ListIndex)
        
        'Clear form
        Call LimparForm_Click
    'If there's no person selected, display a message box
    Else
        MsgBox "Para eliminar uma pessoa precisa de ter uma pessoa selecionada"
    End If
End Sub


