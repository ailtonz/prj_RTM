Option Compare Database

Private Sub cmdPadrao_Click()
    Me.Linha = Me.Linha_Padrao
    Me.Coluna = Me.Coluna_Padrao
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
End Sub
