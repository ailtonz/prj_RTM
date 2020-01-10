Attribute VB_Name = "basMovimentos"

Sub Movimentos()
Attribute Movimentos.VB_ProcData.VB_Invoke_Func = " \n14"

Dim NewSheet As Variant
Dim Auxiliar As String
Dim InicioPesquisa As String
Dim ColunaCaminho As String
Dim ColunaStatus As String

Dim Origem As String
Dim Destino As String

Sheets("Movimentos").Select

Origem = Range("B2")
Destino = Range("B3")

Auxiliar = Range("H2")
InicioPesquisa = Range("H3")
ColunaCaminho = Range("H4")
ColunaStatus = Range("H5")

Sheets(Auxiliar).Select

For x = InicioPesquisa To Rows.Count

    If Sheets(Auxiliar).Range(ColunaCaminho & x) <> "" Then
        
        If Dir(Sheets(Auxiliar).Range(ColunaCaminho & x)) <> "" Then
            
            Sheets(Auxiliar).Range(ColunaStatus & x) = ""
        
            Dim qtdLinhas As Long
            
            Sheets(Destino).Select
            Range("A" & Range("E5")).Select
            Workbooks.Open FileName:=Sheets(Auxiliar).Range(ColunaCaminho & x)
            
            Sheets(Origem).Select
            
            If Range("B4") > 9 Then
            
                qtdLinhas = Range("B4")
                Range("D9:R" & Range("B4") - 1).Select
                Selection.Copy
                
                ActiveWindow.ActivateNext
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                
                ActiveWindow.ActivateNext
                Application.CutCopyMode = False
                Selection.ClearContents
                Range("B4") = 9
                Range("D9").Select
                ActiveWorkbook.Save
                ActiveWindow.Close
                
                Range("E5") = Range("E5") + (qtdLinhas - 9)
                
                Sheets(Auxiliar).Range(ColunaStatus & x) = "ok"
            
            End If
        
        End If

    Else

        Exit For

    End If

Next


Sheets(Destino).Select

End Sub
