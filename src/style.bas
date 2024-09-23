Attribute VB_Name = "style"
Sub stylefrmMain(frm As UserForm)
    '#ced4da
    frm.BackColor = converterRGBForVBA("#F7F8F9")
    
    frm.Controls("frameMain").BackColor = vbWhite
 
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
    
    With frm.Controls("lblContactame")
        .Font.Underline = True
        .ControlTipText = "Desarrollado por: Jhony Escriba " & vbCrLf & " Correo :jhonny14_1@hotmail.com"
        .BackStyle = fmBackStyleTransparent
    End With
    
End Sub
