VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmControleDeUsuarios 
   Caption         =   "Controle de Usuários"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340.001
   OleObjectBlob   =   "frmControleDeUsuarios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmControleDeUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Dim sqlPermissoes As String
Dim sqlSelecao As String

Private Sub cboPermissoes_Click()
Dim strBanco As String: strBanco = Range("B1")
    
    sqlSelecao = "SELECT Selecionado FROM qryPermissoesUsuarios WHERE USUARIO = '" & Me.cboUsuario.Text & "' AND Categoria = '" & Me.cboPermissoes.Text & "'"
        
    sqlPermissoes = "Select * from qryPermissoesItens where Grupo = '" & Me.cboPermissoes.Text & "' and Item not in (" & sqlSelecao & ")"
       
    CarregarListBox strBanco, Me, Me.lstItensEmUso.Name, "Selecionado", sqlSelecao
    
    CarregarListBox strBanco, Me, Me.lstItensDisponiveis.Name, "ITEM", sqlPermissoes
        
End Sub

Private Sub cboPermissoes_Enter()
Dim strBanco As String: strBanco = Range("B1")
Dim strSQL As String

    strSQL = "qryPermissoesGrupos"
    
    CarregarComboBox strBanco, Me.cboPermissoes, "Grupo", strSQL

End Sub

Private Sub cboUsuario_Enter()
Dim strBanco As String: strBanco = Range("B1")
Dim strSQL As String

    strSQL = "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No)) Order By Usuario"

    CarregarComboBox strBanco, Me.cboUsuario, "Usuario", strSQL
    
    Me.cboPermissoes.Clear

    Me.lstItensDisponiveis.Clear
    
    Me.lstItensEmUso.Clear

End Sub

Private Sub lstItensDisponiveis_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strBanco As String: strBanco = Range("B1")
Dim strMSG As String
Dim strTitulo As String

    If IsNull(Me.lstItensDisponiveis.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Seleção de Item disponivel"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
    
        admUsuariosPermissoes strBanco, Me.cboUsuario, Me.lstItensDisponiveis, Me.cboPermissoes
        
        cboPermissoes_Click
    End If
    
End Sub

Private Sub lstItensEmUso_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strBanco As String: strBanco = Range("B1")
Dim strMSG As String
Dim strTitulo As String

    If IsNull(Me.lstItensEmUso.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Remoção de Item em uso"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else

        admUsuariosPermissoesExcluir strBanco, Me.cboUsuario, Me.lstItensEmUso, Me.cboPermissoes
        
        cboPermissoes_Click
        
    End If
    
End Sub

''#########################################
''  FORMULARIO
''#########################################

Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range("B1")

    Me.cboDepartamento.AddItem "ADM"
    
    CarregarComboBox strBanco, Me.cboDepartamento, "Departamento", "qryDepartamentos"

    CarregarListBox strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
    CarregarListBox strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"

End Sub

''#########################################
''  COMANDOS
''#########################################

Private Sub cmdSalvar_Enter()
    Me.txtEmail = LCase(Me.txtEmail)
End Sub

Private Sub cmdSalvar_Click()
Dim strBanco As String: strBanco = Range("B1")

    If ExistenciaUsuario(Range("B1"), Me.txtCodigo, Me.txtNome) Then
        admUsuarioSalvar Range("B1"), Me.cboDepartamento, Me.txtCodigo, Me.txtNome, Me.txtEmail
    Else
        admUsuarioNovo Range("B1"), Me.cboDepartamento, Me.txtCodigo, Me.txtNome, Me.txtEmail
    End If
    
    LimparCampos
    
    CarregarListBox strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
    
End Sub

Private Sub cmdCancelar_Click()
    LimparCampos
End Sub

Private Sub cmdExcluir_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String
Dim strSelecao As String


    If Me.lstUsuarios.Value = "" Or IsNull(Me.lstUsuarios.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
      
        matriz = Array()
        matriz = Split(Me.lstUsuarios.Value, " - ")
        
        admUsuarioExcluir Range("B1"), CStr(matriz(1)), True
        
        CarregarListBox strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
        CarregarListBox strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
                
        LimparCampos
        
    End If

End Sub

Private Sub cmdRestaurar_Click()
Dim strBanco As String: strBanco = Range("B1")
Dim matriz As Variant
Dim strMSG As String
Dim strTitulo As String
Dim strSelecao As String

    If Me.lstUsuariosExcluidos.Value = "" Or IsNull(Me.lstUsuariosExcluidos.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "RESTAURAR!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else

        matriz = Array()
        matriz = Split(Me.lstUsuariosExcluidos.Value, " - ")
        
        admUsuarioExcluir Range("B1"), CStr(matriz(1)), False
        
        CarregarListBox strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
        CarregarListBox strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
        
        LimparCampos

    End If

    
End Sub

''#########################################
''  CAMPOS
''#########################################

Private Sub txtNome_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtNome = UCase(Me.txtNome)
End Sub
Private Sub txtCodigo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtCodigo = UCase(Me.txtCodigo)
End Sub

Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtEmail = LCase(Me.txtEmail)
End Sub

''#########################################
''  LISTAS
''#########################################

Private Sub lstUsuarios_Click()
Dim matriz As Variant

    matriz = Array()
    matriz = Split(Me.lstUsuarios.Value, " - ")
    
    Me.cboDepartamento.Text = CStr(matriz(0))
    Me.txtNome = CStr(matriz(1))
    Me.txtEmail = CStr(matriz(2))
    Me.txtCodigo = CStr(matriz(3))

    Me.cmdSalvar.Enabled = True

End Sub

Private Sub lstUsuarios_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdExcluir_Click
End Sub

Private Sub lstUsuariosExcluidos_Click()
Dim matriz As Variant

    matriz = Array()
    matriz = Split(Me.lstUsuariosExcluidos.Value, " - ")
    
    Me.cboDepartamento.Text = CStr(matriz(0))
    Me.txtNome = CStr(matriz(1))
    Me.txtEmail = CStr(matriz(2))
    Me.txtCodigo = CStr(matriz(3))
    
    Me.cmdSalvar.Enabled = False
End Sub

Private Sub lstUsuariosExcluidos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdRestaurar_Click
End Sub

''#########################################
''  PROCEDIMENTOS
''#########################################

Private Sub LimparCampos()

    Me.cboDepartamento.Text = "DPTO"
    Me.txtCodigo.Text = "CODIGO"
    Me.txtNome.Text = "NOME"
    Me.txtEmail.Text = "E-MAIL"
    
End Sub

Public Function ExistenciaUsuario(BaseDeDados As String, strCODIGO As String, strNOME As String) As Boolean: ExistenciaUsuario = False
On Error GoTo ExistenciaUsuario_err
Dim dbOrcamento As DAO.Database
Dim rstExistenciaUsuario As DAO.Recordset
Dim strSQL As String
Dim retVal As Variant

retVal = Dir(BaseDeDados)

If retVal = "" Then

    ExistenciaUsuario = True
    
Else
   
    strSQL = "SELECT * FROM qryUsuarios WHERE Usuario = '" & strNOME & "' AND  Codigo = '" & strCODIGO & "' "
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
    Set rstExistenciaUsuario = dbOrcamento.OpenRecordset(strSQL)
      
    If rstExistenciaUsuario.EOF Then
        ExistenciaUsuario = False
    Else
        ExistenciaUsuario = True
    End If
        
    rstExistenciaUsuario.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstExistenciaUsuario = Nothing
    
End If

ExistenciaUsuario_Fim:
  
    Exit Function
ExistenciaUsuario_err:
    ExistenciaUsuario = True
    MsgBox Err.Description
    Resume ExistenciaUsuario_Fim
End Function

Private Function admUsuarioNovo(BaseDeDados As String, strDPTO As String, strCODIGO As String, strNOME As String, strEMAIL As String) As Boolean: admUsuarioNovo = True
On Error GoTo admUsuarioNovo_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovo As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovo = dbOrcamento.QueryDefs("admUsuarioNovo")

With qdfadmUsuarioNovo

    .Parameters("CODUSUARIO") = strCODIGO
    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EMAIL_USUARIO") = strEMAIL
    .Parameters("DPTO_USUARIO") = strDPTO
    
    .Execute
    
End With

admUsuarioNovoDepartamentos BaseDeDados, strNOME

admUsuarioNovoFuncoes BaseDeDados, strNOME

admUsuarioNovoNotificacoes BaseDeDados, strNOME

admUsuarioNovoStatus BaseDeDados, strNOME

admUsuarioNovoUsuarios BaseDeDados, strNOME

qdfadmUsuarioNovo.Close
dbOrcamento.Close

admUsuarioNovo_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovo = Nothing
    
    Exit Function
admUsuarioNovo_err:
    admUsuarioNovo = False
    MsgBox Err.Description
    Resume admUsuarioNovo_Fim
End Function

Private Function admUsuarioSalvar(BaseDeDados As String, strDPTO As String, strCODIGO As String, strNOME As String, strEMAIL As String) As Boolean: admUsuarioSalvar = True
On Error GoTo admUsuarioSalvar_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioSalvar As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioSalvar = dbOrcamento.QueryDefs("admUsuarioSalvar")

With qdfadmUsuarioSalvar

    .Parameters("CODUSUARIO") = strCODIGO
    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EMAIL_USUARIO") = strEMAIL
    .Parameters("DPTO_USUARIO") = strDPTO
    
    .Execute
    
End With

qdfadmUsuarioSalvar.Close
dbOrcamento.Close

admUsuarioSalvar_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioSalvar = Nothing
    
    Exit Function
admUsuarioSalvar_err:
    admUsuarioSalvar = False
    MsgBox Err.Description
    Resume admUsuarioSalvar_Fim
End Function

Private Function admUsuarioExcluir(BaseDeDados As String, strNOME As String, Excluir As Boolean) As Boolean: admUsuarioExcluir = True
On Error GoTo admUsuarioExcluir_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioExcluir As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioExcluir = dbOrcamento.QueryDefs("admUsuarioExcluir")

With qdfadmUsuarioExcluir

    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EXCLUSAO") = Excluir
    
    .Execute
    
End With

qdfadmUsuarioExcluir.Close
dbOrcamento.Close

admUsuarioExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioExcluir = Nothing
    
    Exit Function
admUsuarioExcluir_err:
    admUsuarioExcluir = False
    MsgBox Err.Description
    Resume admUsuarioExcluir_Fim
End Function


Private Function admUsuariosPermissoes(BaseDeDados As String, strUsuario As String, strPERMISSAO As String, strCATEGORIA As String) As Boolean: admUsuariosPermissoes = True
On Error GoTo admUsuariosPermissoes_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuariosPermissoes As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuariosPermissoes = dbOrcamento.QueryDefs("admUsuariosPermissoes")

With qdfadmUsuariosPermissoes

    .Parameters("NM_USUARIO") = strUsuario
    .Parameters("NM_PERMISSAO") = strPERMISSAO
    .Parameters("NM_CATEGORIA") = strCATEGORIA
    
    .Execute
    
End With

qdfadmUsuariosPermissoes.Close
dbOrcamento.Close

admUsuariosPermissoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuariosPermissoes = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    admUsuariosPermissoes = False
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function

Private Function admUsuariosPermissoesExcluir(BaseDeDados As String, strUsuario As String, strPERMISSAO As String, strCATEGORIA As String) As Boolean: admUsuariosPermissoesExcluir = True
On Error GoTo admUsuariosPermissoesExcluir_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuariosPermissoesExcluir As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuariosPermissoesExcluir = dbOrcamento.QueryDefs("admUsuariosPermissoesExcluir")

With qdfadmUsuariosPermissoesExcluir

    .Parameters("NM_USUARIO") = strUsuario
    .Parameters("NM_PERMISSAO") = strPERMISSAO
    .Parameters("NM_CATEGORIA") = strCATEGORIA
    
    .Execute
    
End With

qdfadmUsuariosPermissoesExcluir.Close
dbOrcamento.Close

admUsuariosPermissoesExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuariosPermissoesExcluir = Nothing
    
    Exit Function
admUsuariosPermissoesExcluir_err:
    admUsuariosPermissoesExcluir = False
    MsgBox Err.Description
    Resume admUsuariosPermissoesExcluir_Fim
End Function


Private Function admUsuarioNovoDepartamentos(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoDepartamentos = True
On Error GoTo admUsuarioNovoDepartamentos_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovoDepartamentos As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovoDepartamentos = dbOrcamento.QueryDefs("admUsuarioNovoDepartamentos")

With qdfadmUsuarioNovoDepartamentos

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoDepartamentos.Close
dbOrcamento.Close

admUsuarioNovoDepartamentos_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoDepartamentos = Nothing
    
    Exit Function
admUsuarioNovoDepartamentos_err:
    admUsuarioNovoDepartamentos = False
    MsgBox Err.Description
    Resume admUsuarioNovoDepartamentos_Fim
End Function


Private Function admUsuarioNovoFuncoes(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoFuncoes = True
On Error GoTo admUsuarioNovoFuncoes_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovoFuncoes As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovoFuncoes = dbOrcamento.QueryDefs("admUsuarioNovoFuncoes")

With qdfadmUsuarioNovoFuncoes

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoFuncoes.Close
dbOrcamento.Close

admUsuarioNovoFuncoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoFuncoes = Nothing
    
    Exit Function
admUsuarioNovoFuncoes_err:
    admUsuarioNovoFuncoes = False
    MsgBox Err.Description
    Resume admUsuarioNovoFuncoes_Fim
End Function


Private Function admUsuarioNovoNotificacoes(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoNotificacoes = True
On Error GoTo admUsuarioNovoNotificacoes_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovoNotificacoes As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovoNotificacoes = dbOrcamento.QueryDefs("admUsuarioNovoNotificacoes")

With qdfadmUsuarioNovoNotificacoes

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoNotificacoes.Close
dbOrcamento.Close

admUsuarioNovoNotificacoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoNotificacoes = Nothing
    
    Exit Function
admUsuarioNovoNotificacoes_err:
    admUsuarioNovoNotificacoes = False
    MsgBox Err.Description
    Resume admUsuarioNovoNotificacoes_Fim
End Function


Private Function admUsuarioNovoStatus(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoStatus = True
On Error GoTo admUsuarioNovoStatus_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovoStatus As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovoStatus = dbOrcamento.QueryDefs("admUsuarioNovoStatus")

With qdfadmUsuarioNovoStatus

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoStatus.Close
dbOrcamento.Close

admUsuarioNovoStatus_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoStatus = Nothing
    
    Exit Function
admUsuarioNovoStatus_err:
    admUsuarioNovoStatus = False
    MsgBox Err.Description
    Resume admUsuarioNovoStatus_Fim
End Function


Private Function admUsuarioNovoUsuarios(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoUsuarios = True
On Error GoTo admUsuarioNovoUsuarios_err
Dim dbOrcamento As DAO.Database
Dim qdfadmUsuarioNovoUsuarios As DAO.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
Set qdfadmUsuarioNovoUsuarios = dbOrcamento.QueryDefs("admUsuarioNovoUsuarios")

With qdfadmUsuarioNovoUsuarios

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoUsuarios.Close
dbOrcamento.Close

admUsuarioNovoUsuarios_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoUsuarios = Nothing
    
    Exit Function
admUsuarioNovoUsuarios_err:
    admUsuarioNovoUsuarios = False
    MsgBox Err.Description
    Resume admUsuarioNovoUsuarios_Fim
End Function
