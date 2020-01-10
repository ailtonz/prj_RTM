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
        
        strMSG = "ATENÇÃO: Banco de dados não encontrado. " & Chr(10) & Chr(13) & Chr(13) & "Você deseja informar onde está o banco de dados?"
        strTitulo = "Banco de dados não encontrado"
        
        retVal = MsgBox(strMSG, vbQuestion + vbYesNo, strTitulo)
        
        If retVal = vbYes Then
            strBanco = SelecionarBanco
            
            If strBanco <> "" Then
                DesbloqueioDeGuia SenhaBloqueio
                Range("B1").Value = strBanco
                BloqueioDeGuia SenhaBloqueio
                
                strMSG = "Localização do bando de dados está atualizada!"
                strTitulo = "Localização Banco de dados"
                
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
