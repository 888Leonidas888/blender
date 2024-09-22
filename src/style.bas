Attribute VB_Name = "style"
Sub stylefrmMain(frm As UserForm)
    '#ced4da
    frm.BackColor = converterRGBForVBA("#CED4DA")
    
    frm.Controls("frameMain").BackColor = converterRGBForVBA("#A5B0BB")
    frm.Controls("lblStatusProcess").BackStyle = fmBackStyleTransparent
    
    With frm.Controls("Label1")
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
    End With
    
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
        .BackColor = converterRGBForVBA("#CED4DA")
    End With
    
    
    
End Sub
