Option Compare Database
Option Explicit

'Private Sub DescricaoDoProduto_NotInList(NewData As String, Response As Integer)
''Permite adicionar a editora à lista
'Dim db As DAO.Database
'Dim rst As DAO.Recordset
'
'On Error GoTo ErrHandler
'
''Pergunta se deseja acrescentar o novo item
'If Confirmar("O Produtos: " & NewData & "  não faz parte da " & _
'"lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
'    Set db = CurrentDb()
'    'Abre a tabela, adiciona o novo item e atualiza a combo
'    Set rst = db.OpenRecordset("Produtos")
'    With rst
'        .AddNew
'        !codProduto = NovoCodigo("Produtos", "codProduto")
'        !DescricaoDoProduto = NewData
'        .Update
'        Response = acDataErrAdded
'        .Close
'    End With
'Else
'    Response = acDataErrDisplay
'End If
'
'ExitHere:
'Set rst = Nothing
'Set db = Nothing
'Exit Sub
'
'ErrHandler:
'MsgBox Err.Description & vbCrLf & Err.Number & _
'vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
'Resume ExitHere
'End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

Private Sub ValorUnitario_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub
