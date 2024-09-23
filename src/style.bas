Attribute VB_Name = "style"
Sub stylefrmMain(frm As UserForm)
    '#ced4da
    frm.BackColor = converterRGBForVBA("#F7F8F9")
    
    frm.Controls("frameMain").BackColor = vbWhite
    frm.Controls("lblStatusProcess").BackStyle = fmBackStyleTransparent
    
    For i = 1 To 5
        With frm.Controls("Label" & i)
            .BackStyle = fmBackStyleTransparent
            .Font.Bold = True
        End With
    Next i
    
    With frm.Controls("Label2")
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
    End With
    
    With frm.Controls("txtPathExcelShipmentLine")
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbGrayText
    End With
    
    With frm.Controls("txtPathExcelStock")
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbGrayText
    End With
    
    With frm.Controls("btnGenerateExcel")
        .BackColor = vbWhite
    End With
    
End Sub
