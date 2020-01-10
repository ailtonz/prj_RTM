VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
On Error GoTo Workbook_BeforeSave_err

Dim strMSG As String
Dim strTitulo As String
Dim retVal As Variant

Dim dbOrcamento As DAO.Database
Dim rstOrcamento As DAO.Recordset
Dim rstIntervalos As DAO.Recordset
Dim rstNotificacoes As DAO.Recordset

Dim strOrcamento As String
Dim strIntervalos As String
Dim strNotificacoes As String


Dim strControle As String: strControle = Range("J3")
Dim strVendedor As String: strVendedor = Range("C3")
Dim strBaseDeDados As String: strBaseDeDados = Range("B1")

'   BLOQUEIO TELA
Application.ScreenUpdating = False

'   Campos Obrigatórios
'   PRODUTO
'If Range("G5") = "" Then
'    MsgBox "ATENÇÃO: Favor preencher o Campo PRODUTO!", vbCritical + vbOKOnly, "Campo Obrigatório"
'    Range("G5").Select
'    Exit Sub
'End If

''   PROJETO
'If Range("B7") = "" Then
'    MsgBox "ATENÇÃO: Favor preencher o Campo PROJETO!", vbCritical + vbOKOnly, "Campo Obrigatório"
'    Range("B7").Select
'    Exit Sub
'End If
'
''   LICENCIADO
'If Range("G6") = "" Then
'    MsgBox "ATENÇÃO: Favor preencher o Campo LICENCIADO!", vbCritical + vbOKOnly, "Campo Obrigatório"
'    Range("G6").Select
'    Exit Sub
'End If

    
If strControle <> "" And strVendedor <> "" Then
    
    '   BASE DE DADOS
    strOrcamento = "SELECT Orcamentos.* " & _
         " FROM Orcamentos  " & _
         " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "
    Set dbOrcamento = DBEngine.OpenDatabase(strBaseDeDados)
    
    '   ORÇAMENTOS
    Set rstOrcamento = dbOrcamento.OpenRecordset(strOrcamento)
    
    '   INTERVALOS
    strIntervalos = "Select * from qryEtapasIntervalosEdicoes where Departamento = '" & rstOrcamento.Fields("Departamento") & "' and Status = '" & rstOrcamento.Fields("Status") & "'"
    Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)
  
    '   QUESTIONA TERMINIO DE ETAPA
    strMSG = "ATENÇÃO: Esta etapa está finalizada?. " & Chr(10) & Chr(13) & Chr(13) & _
             "Click em SIM caso deseje enviar o Orçamento para a proxima etapa." & Chr(10) & Chr(13) & Chr(13) & _
             "Click em NÃO caso deseje continuar alterando no futuro."
    strTitulo = "Fluxo de etapas dos orçamentos"
    retVal = MsgBox(strMSG, vbQuestion + vbYesNo, strTitulo)
        
        
    If retVal = vbYes Then
        
        '   ENVIAR ORÇAMENTO PARA PROXIMA ETAPA
        admOrcamentoEtapaAvancar strBaseDeDados, strControle, strVendedor
        
        '   ENVIAR NOTIFICAÇÕES
        strNotificacoes = "SELECT DISTINCT qryPermissoesUsuarios.Selecionado AS Status, qryPermissoesUsuarios.eMail, " & _
                          " qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                          "(((qryPermissoesUsuarios.Selecionado)='" & rstOrcamento.Fields("Status") & "') AND ((qryPermissoesUsuarios.DPTO)<>'Vendas') AND " & _
                          "((qryPermissoesUsuarios.Categoria)='Notificações'))"
        
        strNotificacoes = strNotificacoes + "UNION SELECT qryPermissoesUsuarios.Selecionado AS Status, qryPermissoesUsuarios.eMail, " & _
                          "qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                          "(((qryPermissoesUsuarios.Selecionado)='" & rstOrcamento.Fields("Status") & "') AND ((qryPermissoesUsuarios.DPTO)='Vendas') " & _
                          "AND ((qryPermissoesUsuarios.Usuario)='" & strVendedor & "') AND ((qryPermissoesUsuarios.Categoria)='Notificações'))"
        
'        Saida strNotificacoes, "ENVIO.LOG"
        
        Set rstNotificacoes = dbOrcamento.OpenRecordset(strNotificacoes)
        While Not rstNotificacoes.EOF
        
            EnviarEmail rstNotificacoes.Fields("eMail"), rstOrcamento.Fields("Status") & " : " & rstOrcamento.Fields("Controle") & " - " & rstOrcamento.Fields("Vendedor")
            rstNotificacoes.MoveNext
            
        Wend
        
        rstNotificacoes.Close
    
    End If
               
    ' SALVAR CADASTROS DISPONIVEIS
    While Not rstIntervalos.EOF
    
        If rstIntervalos.Fields("Intervalo") = "ORÇAMENTO" Then
        
            CadastroOrcamento strBaseDeDados, strControle, strVendedor
            
        ElseIf rstIntervalos.Fields("Intervalo") = "ESPECIAL" Then
        
            CadastroEspecial strBaseDeDados, strControle, strVendedor
        
        ElseIf rstIntervalos.Fields("Intervalo") = "VENDA" Then
        
            CadastroVenda strBaseDeDados, strControle, strVendedor
        
        ElseIf rstIntervalos.Fields("Intervalo") = "CUSTOS" Then
        
            CadastroPrevisao strBaseDeDados, strControle, strVendedor
            
        ElseIf rstIntervalos.Fields("Intervalo") = "RENDIMENTO" Then
        
            CadastroRendimento strBaseDeDados, strControle, strVendedor
            copiarProduzido
            CadastroFinanceiro strBaseDeDados, strControle, strVendedor
        
        ElseIf rstIntervalos.Fields("Intervalo") = "FINANCEIRO" Then
        
            CadastroFinanceiro strBaseDeDados, strControle, strVendedor
            copiarAnalise
            CadastroFinal strBaseDeDados, strControle, strVendedor
            CadastroIndice strBaseDeDados, strControle, strVendedor
            
        ElseIf rstIntervalos.Fields("Intervalo") = "LIBERAÇÃO" Then
        
            CadastroLiberacao strBaseDeDados, strControle, strVendedor
            
        End If
        
            
        rstIntervalos.MoveNext
    
    Wend
            
    rstOrcamento.Close
    rstIntervalos.Close
    dbOrcamento.Close

End If

'   POSICIONA CURSOR
Range("C3").Select

'   LIMPAR FORMULARIOS
DesbloqueioDeGuia SenhaBloqueio
admOrcamentoLimpar
BloqueioDeGuia SenhaBloqueio

'   LIBERAR TELA
Application.ScreenUpdating = True


Workbook_BeforeSave_Fim:
    Set dbOrcamento = Nothing
    Set rstOrcamento = Nothing
    Set rstIntervalos = Nothing
    Set rstNotificacoes = Nothing
    
    Exit Sub
Workbook_BeforeSave_err:

    MsgBox Err.Description
    Resume Workbook_BeforeSave_Fim


End Sub


Private Sub Workbook_Open()
        
'    DesbloqueioDeGuia SenhaBloqueio
'
'    admOrcamentoLimpar
'
'    BloqueioDeGuia SenhaBloqueio
    
End Sub
