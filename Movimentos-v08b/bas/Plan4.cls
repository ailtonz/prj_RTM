VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdSemanal_Click()
    Semanal
End Sub

Private Sub Semanal()
Attribute Semanal.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Origem As String                'ORIGEM - (FONTE DE DADOS)
Dim Destino As String               'DESTINO - (FONTE DE DADOS)

Dim ColunaOrigem As String         'COLUNA - ORIGEM
Dim ColunaDestino As String        'COLUNA - DESTINO

Dim LinhaOrigem As String           'LINHA - ORIGEM
Dim LinhaDestino As String          'LINHA - DESTINO

Dim ProxLinhaOrigem As String      'PROX. LINHA - ORIGEM
Dim ProxLinhaDestino As String     'PROX. LINHA - DESTINO

Dim ColunaComMarca As String           'COLUNA COM A MARCA - ORIGEM
Dim InicioPesquisa As String     'INICIO DA PESQUISA - ORIGEM

Dim Marca As String

Dim Contador As Long                    'Contador
Dim x As Long

Dim valOrigem(255) As String
Dim valDestino(255) As String


'Carrega variaveis
Origem = Range("B3")
Destino = Range("B4")
LinhaOrigem = Range("E3")
LinhaDestino = Range("E4")
ColunaOrigem = Range("F3")
ColunaDestino = Range("F4")
ProxLinhaDestino = Range("E5")
InicioPesquisa = Range("E6")

x = 1

For Contador = CInt(InicioPesquisa) To Columns.Count
    valOrigem(x) = Cells(LinhaOrigem, Contador)
    valDestino(x) = Cells(LinhaDestino, Contador)
    x = x + 1
Next


Sheets(Origem).Select

ColunaComMarca = Sheets(Origem).Range("E2")
Marca = Sheets(Origem).Range("E3")
InicioPesquisa = Sheets(Origem).Range("E4")
x = 1

For Contador = CInt(InicioPesquisa) To Rows.Count
    If Sheets(Origem).Cells(Contador, CInt(ColunaComMarca)) = Marca Then
        For x = 1 To 255
            If valOrigem(x) <> "" Then
                Range(valDestino(x) & ProxLinhaDestino) = Sheets(Origem).Range(valOrigem(x) & CStr(Contador))
                Range("A" & ProxLinhaDestino) = Now()
            Else
                Sheets(Origem).Cells(Contador, CInt(ColunaComMarca)).Select
                Selection.ClearContents
                Sheets(Origem).Range("A1").Select
                ProxLinhaDestino = ProxLinhaDestino + 1
                Exit For
            End If
        Next x
    End If
Next Contador

Range("E5") = ProxLinhaDestino

Sheets(Destino).Select
Range("A1").Select

MsgBox "A planilha " & Destino & " foi atualizada com sucesso!", vbInformation + vbOKOnly, "Atualizar Semanal"


End Sub

