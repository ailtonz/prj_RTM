VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPesquisar 
   Caption         =   "Pesquisa de Orçamentos - Ver.07"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11055
   OleObjectBlob   =   "frmPesquisar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Public strPesquisar As String
Public strSQL As String
Public strUsuarios As String

Private Sub LiberarFormularios(strUsuario As String)
On Error GoTo LiberarFormularios_err
Dim strBanco As String: strBanco = Range("B1")
Dim dbOrcamento As DAO.Database
Dim rstLiberarFormularios As DAO.Recordset
Dim rstBloquearFormularios As DAO.Recordset


Dim matriz As Variant
'Dim bloqueio As Variant

Set dbOrcamento = DBEngine.OpenDatabase(strBanco)
Set rstLiberarFormularios = dbOrcamento.OpenRecordset("Select * from qryUsuariosFormularios where Usuario = '" & strUsuario & "'")
Set rstBloquearFormularios = dbOrcamento.OpenRecordset("qryFormularios")

matriz = Array()
'bloqueio = Array()

While Not rstBloquearFormularios.EOF
    matriz = Split(rstBloquearFormularios.Fields("VALOR_02"), "-")
    
    OcultarLinhas (matriz(0)), (matriz(1)), True
    rstBloquearFormularios.MoveNext

Wend


While Not rstLiberarFormularios.EOF
    matriz = Split(rstLiberarFormularios.Fields("Formulario"), "-")

    OcultarLinhas CStr(matriz(0)), CStr(matriz(1)), False
    rstLiberarFormularios.MoveNext

Wend


LiberarFormularios_Fim:
    rstBloquearFormularios.Close
    rstLiberarFormularios.Close
    dbOrcamento.Close
    Set rstLiberarFormularios = Nothing
    Set dbOrcamento = Nothing
    
    Exit Sub
LiberarFormularios_err:
    MsgBox Err.Description
    Resume LiberarFormularios_Fim

End Sub


Private Sub cmdControleDeUsuarios_Click()
    frmControleDeUsuarios.Show
End Sub

Private Sub cmdUsuarioPadrao_Click()
Dim strBanco As String: strBanco = Range("B1")

DesbloqueioDeGuia SenhaBloqueio
Application.ScreenUpdating = False

Range("B91") = Me.cboUsuario

CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL

LiberarFormularios Range("B91")

ActiveSheet.Name = Me.cboUsuario

'POSICIONA CURSOR
Range("C3").Select

Application.ScreenUpdating = True
BloqueioDeGuia SenhaBloqueio

UserForm_Initialize

End Sub

Private Sub cmdVoltarEtapa_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String
Dim retVal As Variant

    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Voltar Etapa"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        
        retVal = MsgBox("ATENÇÃO: Você deseja voltar o orçamento para etapa anterior", vbInformation + vbOKOnly, "Fluxo de etapas dos orçamentos")
        
        matriz = Array()
        matriz = Split(lstPesquisa.Value, " - ")
        
        '   VOLTA ORÇAMENTO PARA ETAPA ANTERIOR
        admOrcamentoEtapaVoltar strBanco, CStr(matriz(0)), CStr(matriz(2))
        
        MontarPesquisa

        CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
        
        
    End If
End Sub

''#########################################
''  FORMULARIO
''#########################################

Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range("B1")
Dim sqlUsuarios As String: strUsuarios = Range("B91")
       
MontarPesquisa

sqlUsuarios = "SELECT admCategorias.Descricao01 as Usuario FROM admCategorias WHERE (((admCategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "'))) AND ((admCategorias.Categoria)='Usuarios')) ORDER BY admCategorias.Descricao01"

CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL

CarregarComboBox strBanco, Me.cboUsuario, "Usuario", sqlUsuarios

Me.cboUsuario.Text = strUsuarios
       
DesbloqueioDeFuncoes strBanco, Me, "Select * from qryUsuariosFuncoes Where Usuario = '" & strUsuarios & "'", "Funcao"

CarregarClientes strBanco

CarregarLicenciados strBanco

'POSICIONA CURSOR
Range("C3").Select
              
End Sub

''#########################################
''  COMANDOS
''#########################################

Private Sub cmdPesquisar_Click()
Dim strBanco As String: strBanco = Range("B1")

Dim retValor As Variant

retValor = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", strPesquisar, 0, 0)
strPesquisar = retValor

MontarPesquisa

CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
Me.Repaint

End Sub

Private Sub MontarPesquisa()

strSQL = "SELECT qryOrcamentosListar.Pesquisa FROM qryOrcamentosListar WHERE ((qryOrcamentosListar.Pesquisa) Like '*" & strPesquisar & "*')"
strSQL = strSQL + " AND ((qryOrcamentosListar.VENDEDOR) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Usuarios'))"
strSQL = strSQL + " AND ((qryOrcamentosListar.DEPARTAMENTO) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Departamentos')) "
strSQL = strSQL + " AND ((qryOrcamentosListar.STATUS) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE admCategorias.Categoria = '" & strUsuarios & "') and Categoria = 'Status')) "
strSQL = strSQL + "ORDER BY qryOrcamentosListar.CONTROLE DESC , qryOrcamentosListar.VENDEDOR"

End Sub


Private Sub cmdNovo_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim sqlUsuario As String: strUsuarios = Range("B91")

    admOrcamentoNovo strBanco, Me.cboUsuario
    CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    
End Sub

Private Sub cmdAlterar_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String

    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "ALTERAR ORÇAMENTO!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        matriz = Array()
        matriz = Split(lstPesquisa.Value, " - ")

        Application.ScreenUpdating = False

        DesbloqueioDeGuia SenhaBloqueio
        
        admOrcamentoLimpar
        
        CarregarOrcamento strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarPrevisao strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarRendimento strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarFinanceiro strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarFinal strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarIndice strBanco, CStr(matriz(0)), CStr(matriz(2))
        
                
        Range("C4").Select
        
        BloqueioDeGuia SenhaBloqueio
    
        Application.ScreenUpdating = True
        
        frmPesquisar.Hide
        
    End If

End Sub

Private Sub cmdCopiar_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String

    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "CÓPIA!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        matriz = Array()
        matriz = Split(lstPesquisa.Value, " - ")
        
        admOrcamentoCopiar strBanco, CStr(matriz(0)), CStr(matriz(2))
        CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    End If
End Sub

Private Sub cmdExcluir_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String
Dim varResposta As Variant


    If IsNull(Me.lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        strMSG = "ATENÇÃO: Você deseja realmente EXCLUIR este registro?. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        varResposta = MsgBox(strMSG, vbInformation + vbYesNo, strTitulo)
    
        If varResposta = vbYes Then
            matriz = Array()
            matriz = Split(lstPesquisa.Value, " - ")
    
    
            varResposta = InputBox("Informe o motivo pelo qual o Orçamento foi Excluido.", "Motivo da exclusão")
    
            If varResposta <> "" Then
            
                If admOrcamentoExcluir(strBanco, CStr(matriz(0)), CStr(matriz(2)), CStr(varResposta)) Then
                    strMSG = "Exclusão concluida com sucesso!" & Chr(10) & Chr(13) & Chr(13)
                    strTitulo = "EXCLUIR!"
                    
                    CarregarListBox strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
                Else
                    strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
                    strTitulo = "EXCLUIR!"
                End If
            
            Else
                strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
                strTitulo = "EXCLUIR!"
            End If
            
            MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
        Else
            strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
            strTitulo = "EXCLUIR!"
            
            MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
        End If
        
    End If
    
End Sub



Private Sub cmdBanco_Click()
' VINCULAR BANCO DE DADOS

Dim strMSG As String
Dim strTitulo As String
Dim strBanco As String

strBanco = SelecionarBanco

    If strBanco <> "" Then
    
        DesbloqueioDeGuia SenhaBloqueio
        Range("B1").Value = strBanco
        BloqueioDeGuia SenhaBloqueio
        
    Else
        strMSG = "ATENÇÃO: Por favor Informe onde está o banco. "
        strTitulo = "Seleção do Banco de dados"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    End If
    

End Sub

''#########################################
''  LISTAS
''#########################################

Private Sub lstPesquisa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strMSG As String
Dim strTitulo As String
    
    
    If Me.cmdAlterar.Enabled Then
        cmdAlterar_Click
    Else
        strMSG = "Função Bloqueada! " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Alterar!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    
    End If
    
End Sub

''#########################################
''  PROCEDIMENTOS
''#########################################

Private Function admOrcamentoExcluir(BaseDeDados As String, strControle As String, strNOME As String, strMotivo As String) As Boolean: admOrcamentoExcluir = True
On Error GoTo admOrcamentoExcluir_err
Dim dbOrcamento As DAO.Database
Dim qdfadmOrcamentoExcluir As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmOrcamentoExcluir = dbOrcamento.QueryDefs("admOrcamentoExcluir")

With qdfadmOrcamentoExcluir

    .Parameters("NM_VENDEDOR") = strNOME
    .Parameters("NM_CONTROLE") = strControle
    .Parameters("NM_MOTIVO") = strMotivo
    
    .Execute
    
End With

qdfadmOrcamentoExcluir.Close
dbOrcamento.Close

admOrcamentoExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoExcluir = Nothing
    
    Exit Function
admOrcamentoExcluir_err:
    admOrcamentoExcluir = False
    MsgBox Err.Description
    Resume admOrcamentoExcluir_Fim
End Function

Private Function admOrcamentoNovo(BaseDeDados As String, strVendedor As String) As Boolean: admOrcamentoNovo = True
' CADASTRAR NOVO ORÇAMENTO

On Error GoTo admOrcamentoNovo_err
Dim dbOrcamento As DAO.Database
Dim qdfOrcamentoNovo As DAO.QueryDef
Dim qdfOrcamentoNovoCustos As DAO.QueryDef

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)

'' ORÇAMENTO
Set qdfOrcamentoNovo = dbOrcamento.QueryDefs("admOrcamentoNovo")
With qdfOrcamentoNovo

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With

'' PREVISÕES DE CUSTOS
Set qdfOrcamentoNovoCustos = dbOrcamento.QueryDefs("admOrcamentoNovoCustos")
With qdfOrcamentoNovoCustos

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With



admOrcamentoNovo_Fim:
    dbOrcamento.Close
    qdfOrcamentoNovo.Close
    qdfOrcamentoNovoCustos.Close
    
    Set dbOrcamento = Nothing
    Set qdfOrcamentoNovo = Nothing
    Set qdfOrcamentoNovoCustos = Nothing
    
    Exit Function
admOrcamentoNovo_err:
    admOrcamentoNovo = False
    MsgBox Err.Description
    Resume admOrcamentoNovo_Fim

End Function

Private Function admOrcamentoCopiar(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: admOrcamentoCopiar = True
' CRIAR CÓPIA DE ORÇAMENTO

On Error GoTo admOrcamentoCopiar_err
Dim dbOrcamento As DAO.Database
Dim qdfadmOrcamentoCopiar As DAO.QueryDef
Dim qdfNovoOrcamentoPrevisoesDeCustos As DAO.QueryDef

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)

'' ORÇAMENTO
Set qdfadmOrcamentoCopiar = dbOrcamento.QueryDefs("admOrcamentoCopiar")
With qdfadmOrcamentoCopiar

    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle
    
    .Execute
    
End With

'' PREVISÕES DE CUSTOS
Set qdfNovoOrcamentoPrevisoesDeCustos = dbOrcamento.QueryDefs("admOrcamentoNovoCustos")
With qdfNovoOrcamentoPrevisoesDeCustos

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With


admOrcamentoCopiar_Fim:
    dbOrcamento.Close
    qdfadmOrcamentoCopiar.Close
    qdfNovoOrcamentoPrevisoesDeCustos.Close
    
    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoCopiar = Nothing
    Set qdfNovoOrcamentoPrevisoesDeCustos = Nothing
    
    Exit Function
admOrcamentoCopiar_err:
    admOrcamentoCopiar = False
    MsgBox Err.Description
    Resume admOrcamentoCopiar_Fim
End Function



