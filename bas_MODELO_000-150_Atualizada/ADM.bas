Attribute VB_Name = "ADM"
Dim ErroInterno As Boolean

Sub NovaGuia()
Attribute NovaGuia.VB_Description = "Macro gravada em 27/08/2008 por admin"
Attribute NovaGuia.VB_ProcData.VB_Invoke_Func = " \n14"
Dim sGuia As String
ErroInterno = False

    sGuia = ProxNomeDaGuia("ADM")
    
    If ErroInterno = False Then
    
        If PesquisaNomeGuia(sGuia) = False Then
            Sheets("000-00").Copy After:=Sheets(1)
            Sheets("000-00 (2)").Select
            Sheets("000-00 (2)").Name = sGuia
            Sheets(sGuia).Select
            ActiveWorkbook.Sheets(sGuia).Tab.ColorIndex = 6
            Range("B4").Select
        Else
            MsgBox "ATENÇÃO: Tente Novamente!", vbInformation + vbOKOnly, "Nova Guia"
        End If
    End If
End Sub

Function PesquisaNomeGuia(sGuia As String) As Boolean

For s = 1 To Sheets.Count
    If Sheets(s).Name = sGuia Then
        PesquisaNomeGuia = True
    End If
Next
End Function

Function ProxNomeDaGuia(sPlanilha As String) As String
On Error GoTo Err_ProxNomeDaGuia

Dim linha As Long
Dim NovoNome As String
Dim LocMarca As String

Sheets(sPlanilha).Activate

If Cells(9, 1) = "" Then
    Cells(9, 1) = "000" & "-" & Right(Year(Date), 2)
End If

For linha = 9 To Rows.Count
    If Cells(linha, 1) = "" Then
        'Monta o novo nome
        NovoNome = Format(CLng(Left(Cells(linha - 1, 1), 3)) + 1, "000") & "-" & Right(Year(Date), 2)
        'Cadastra o nome da nova guia
        Range("A" & CStr(linha)).Select
        ActiveCell.FormulaR1C1 = NovoNome
        'Cadastra a versão da guia
        Range("B" & CStr(linha)) = 0
        'Seleciona a primeira celula da lista
        Range("A9").Select
        'Disponibiliza o "Novo Nome" da guia
        ProxNomeDaGuia = NovoNome
        'Termina o loop
        Exit For
    End If
Next


Exit_ProxNomeDaGuia:
    Exit Function
    
    
Err_ProxNomeDaGuia:
    MsgBox "ATENÇÃO: Erro ao gerar nome da proxima guia. Favor checar o Controle de Guias em ADM", vbInformation + vbOKOnly, "Função: Proximo Nome da guia"
    ErroInterno = True
    Resume Exit_ProxNomeDaGuia
    
End Function

Public Function Historico()

Dim Origem(255) As String
Dim Destino(255) As String
Dim Valores(255) As Variant
Dim sNomeGuia As String

Dim i As Long

x = 1

''#######################################################################
'' Seleciona guia de Administração para coletar dados de (Origem/Destino)
''#######################################################################

Sheets("ADM").Select

'Coleta Origem/Destino
For coluna = 4 To Columns.Count
    If Cells(6, coluna) <> "" And Cells(7, coluna) <> "" Then
        Origem(x) = Cells(6, coluna)
        Destino(x) = Cells(7, coluna)
    End If
    x = x + 1
Next

For s = 1 To Sheets.Count

    If Sheets(s).Name <> "ADM" Or Sheets(s).Name <> "000-00" Then

        ''##########################
        '' Seleciona guia de Origem
        ''##########################
        
        'Seleciona a guia com a marcação
        Sheets(s).Select
        
        If Sheets(s).Cells(3, 9) = "x" Then
        
            'Guarda o nome da guia
            sNomeGuia = Sheets(s).Name
            
            ''#######################
            '' Copia dados da Origem
            ''#######################
            For x = 1 To 255
                If Origem(x) <> "" Then
                    Valores(x) = Range(Origem(x))
                End If
            Next
            
            'Apaga Marcação
            Range("I3") = ""
            
            'Seleciona a guia de Administração
            Sheets("ADM").Select
            
            ''######################
            '' Colar dados copiados
            ''######################
            For x = 1 To 255
                If Destino(x) <> "" Then
                    'Dados copiados
                    Range(Destino(x) & Range("b4")) = Valores(x)
                    
                    'Vendedor
                    If Range("C2") <> "" Then
                        Range(Range("C2") & Range("b4")) = Range("B2") & " " & sNomeGuia
                    End If
                End If
            Next
            
            ''###################
            ''Controle de Versões
            ''###################
            For linha = 9 To Rows.Count
                If Cells(linha, 1) = sNomeGuia Then
                    'VERSÃO
                    Range("B" & CStr(linha)) = Range("B" & CStr(linha)) + 1
                    'Nº MUDANÇA
                    Range(Range("b3") & Range("b4")) = Range("B" & CStr(linha))
                End If
            Next
            
            'PROX. LINHA
            Range("b4") = Range("b4") + 1
            
        End If
    End If
Next

Sheets("ADM").Select

End Function

