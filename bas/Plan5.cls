VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdPesquisa_Click()
Dim strMSG As String
Dim strTitulo As String
Dim strBanco As String
Dim retVal As Variant

retVal = Dir(Range("B1"))

    If retVal = "" Then
        
        strMSG = "ATEN��O: Banco de dados n�o encontrado. " & Chr(10) & Chr(13) & Chr(13) & "Voc� deseja informar onde est� o banco de dados?"
        strTitulo = "Banco de dados n�o encontrado"
        
        retVal = MsgBox(strMSG, vbQuestion + vbYesNo, strTitulo)
        
        If retVal = vbYes Then
            strBanco = SelecionarBanco
            
            If strBanco <> "" Then
                DesbloqueioDeGuia SenhaBloqueio
                Range("B1").Value = strBanco
                BloqueioDeGuia SenhaBloqueio
                
                strMSG = "Localiza��o do bando de dados est� atualizada!"
                strTitulo = "Localiza��o Banco de dados"
                
                MsgBox strMSG, vbQuestion + vbOKOnly, strTitulo
                
            End If
            
        End If
    
    Else
    
        DesbloqueioDeGuia SenhaBloqueio
        
        admOrcamentoLimpar
        frmPesquisar.Show
        
        BloqueioDeGuia SenhaBloqueio

    End If
End Sub
