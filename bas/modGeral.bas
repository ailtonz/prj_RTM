Attribute VB_Name = "modGeral"
''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
''#################### [ FUNÇÕES E PROCEDIMENTOS ] ####################
''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx


Public Function Saida(strConteudo As String, strArquivo As String)
    Open CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1
End Function

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function LocalizacaoDaPlanilha() As String
    LocalizacaoDaPlanilha = ActiveWorkbook.Path
End Function

Public Function CaminhoCompleto() As String
    CaminhoCompleto = ActiveWorkbook.FullName
End Function

Public Function BloqueioDeGuia(strSenha As String)
    ActiveSheet.Protect strSenha
End Function

Public Function DesbloqueioDeGuia(strSenha As String)
    ActiveSheet.Unprotect strSenha
End Function

Public Function IntervaloEditacaoCriar(Titulo As String, Selecao As String)
    ActiveSheet.Protection.AllowEditRanges.Add Title:=Titulo, Range:=Range(Selecao)
    DesmarcaTexto Selecao
End Function

Function LimparTemplate(Selecao As String, conteudo As Variant)
    Range(Selecao).Select
    Selection.Value = conteudo
End Function

Public Function OcultarLinhas(LinhaInicio As Integer, LinhaFinal As Integer, ocultar As Boolean)
    Rows(CStr(LinhaInicio) & ":" & CStr(LinhaFinal)).Select
    Selection.EntireRow.Hidden = ocultar
End Function

Public Function SelecionarGuiaAtual()
    Sheets(ActiveSheet.Name).Select
End Function

Public Function PesquisaNomeGuia(sGuia As String) As Boolean
    For S = 1 To Sheets.Count
        If Sheets(S).Name = sGuia Then
            PesquisaNomeGuia = True
        End If
    Next
End Function

Public Function IntervaloEditacaoRemover(IntervaloDeEdicao As String, MarcarSelecao As String)
    Dim AER As AllowEditRange
    
    For Each AER In ActiveSheet.Protection.AllowEditRanges
        If AER.Title = IntervaloDeEdicao Then
            AER.Delete
            MarcaTexto MarcarSelecao
        End If
    Next AER

End Function

Public Function IntervaloEditacaoRemoverTodos()
    Dim AER As AllowEditRange
    Dim x As Integer
    x = ActiveSheet.Protection.AllowEditRanges.Count
    
    'DesbloqueioDeGuia SenhaBloqueio
    
    For Each AER In ActiveSheet.Protection.AllowEditRanges
        If x > 0 Then
            AER.Delete
        End If
    Next AER
    
    'BloqueioDeGuia SenhaBloqueio
    
    Set AER = Nothing

End Function

Public Function MarcaTexto(Selecao As String)
    
    Range(Selecao).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Function

Public Function DesmarcaTexto(Selecao As String)
    
    Range(Selecao).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Function

Public Function SelecionarBanco() As String
    Dim fd As Office.FileDialog
    Dim strArq As String
    
    On Error GoTo SelecionarBanco_err
    
    'Diálogo de selecionar arquivo - Office
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
    fd.Title = "Localize a fonte de dados"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
    End If
        
    If strArq <> "" Then SelecionarBanco = strArq

SelecionarBanco_Fim:
    Exit Function

SelecionarBanco_err:
    MsgBox Err.Description
    Resume SelecionarBanco_Fim

End Function

Public Function CarregarComboBox(BaseDeDados As String, cbo As ComboBox, strCampo As String, strSQL As String) As Boolean: CarregarComboBox = True
On Error GoTo CarregarComboBox_err
Dim dbOrcamento As DAO.Database
Dim rstCarregarComboBox As DAO.Recordset
Dim retVal As Variant

retVal = Dir(BaseDeDados)

If retVal = "" Then

    CarregarComboBox = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
    Set rstCarregarComboBox = dbOrcamento.OpenRecordset(strSQL)
    
    cbo.Clear
    
    While Not rstCarregarComboBox.EOF
        cbo.AddItem rstCarregarComboBox.Fields(strCampo)
        rstCarregarComboBox.MoveNext
    Wend
        
    rstCarregarComboBox.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarComboBox = Nothing
    
End If

CarregarComboBox_Fim:
  
    Exit Function
CarregarComboBox_err:
    CarregarComboBox = False
    MsgBox Err.Description
    Resume CarregarComboBox_Fim
End Function

Public Function CarregarListBox(BaseDeDados As String, frm As UserForm, NomeLista As String, strCampo As String, strSQL As String)
On Error GoTo CarregarListBox_err

Dim dbOrcamento         As DAO.Database
Dim rstCarregarListBox   As DAO.Recordset
Dim retVal              As Variant

Dim ctrl                As Control

retVal = Dir(BaseDeDados)

If retVal = "" Then

    CarregarListBox = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
    Set rstCarregarListBox = dbOrcamento.OpenRecordset(strSQL)
    
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "ListBox" Then
            If ctrl.Name = NomeLista Then
                ctrl.Clear
                While Not rstCarregarListBox.EOF
                    ctrl.AddItem rstCarregarListBox.Fields(strCampo)
                    rstCarregarListBox.MoveNext
                Wend
            End If
        End If
    Next
    
    rstCarregarListBox.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarListBox = Nothing
    
End If

CarregarListBox_Fim:
  
    Exit Function
CarregarListBox_err:
    CarregarListBox = False
    MsgBox Err.Description
    Resume CarregarListBox_Fim
End Function

Public Function DesbloqueioDeFuncoes(BaseDeDados As String, frm As UserForm, strSQL As String, strCampo As String)
On Error GoTo DesbloqueioDeFuncoes_err

Dim dbOrcamento         As DAO.Database
Dim rstDesbloqueioDeFuncoes   As DAO.Recordset
Dim retVal              As Variant
Dim ctrl                As Control

retVal = Dir(BaseDeDados)

If retVal = "" Then

    DesbloqueioDeFuncoes = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados)
    Set rstDesbloqueioDeFuncoes = dbOrcamento.OpenRecordset(strSQL)
        
    While Not rstDesbloqueioDeFuncoes.EOF
        For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            If Right(ctrl.Name, Len(ctrl.Name) - 3) = rstDesbloqueioDeFuncoes.Fields(strCampo) Then
                ctrl.Enabled = True
            End If
            End If
        Next
        rstDesbloqueioDeFuncoes.MoveNext
    Wend
    
    rstDesbloqueioDeFuncoes.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstDesbloqueioDeFuncoes = Nothing
    
End If

DesbloqueioDeFuncoes_Fim:
  
    Exit Function
DesbloqueioDeFuncoes_err:
    DesbloqueioDeFuncoes = False
    MsgBox Err.Description
    Resume DesbloqueioDeFuncoes_Fim
End Function

Sub CadastroDeLista(CampoDaLista As String, LetraColunaDaLista As String, InicioDaLista As Integer, TerminioDaLista As Integer)
' OBJETIVO: Posiciona uma lista de valores em um campo determinado.

    Range(CampoDaLista).Select
    
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$" & LetraColunaDaLista & "$" & InicioDaLista & ":$" & LetraColunaDaLista & "$" & TerminioDaLista & ""
        
'        Saida "=$" & LetraColunaDaLista & "$" & InicioDaLista & ":$" & LetraColunaDaLista & "$" & TerminioDaLista & "", "CadastroDeLista.LOG"
        
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub EnviarEmail(strEMAIL As String, strAssunto As String)
On Error GoTo EnviarEmail_err
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = strEMAIL
        .CC = ""
        .BCC = ""
        .Subject = strAssunto 'ActiveSheet.Name
'        .Body = "Hello World!"
'        .Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        '.Attachments.Add ("C:\test.txt")
        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        .Send
    End With
    On Error GoTo 0

EnviarEmail_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing
  
    Exit Sub
EnviarEmail_err:
    MsgBox Err.Description
    Resume EnviarEmail_Fim
    
End Sub

Public Function MarcarObrigatorio(Selecao As String)

    Range(Selecao).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Function

