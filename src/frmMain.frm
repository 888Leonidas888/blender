VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "UserForm1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExplorer1_Click()
    txtPathExcelStock.value = selectFileXlsx()
End Sub

Private Sub btnExplorer2_Click()
    txtPathExcelShipmentLine.value = selectFileXlsx()
End Sub

Private Sub btnGenerateExcel_Click()

    Dim fso As New Scripting.FileSystemObject
    Dim ddbbStock As String
    Dim ddbbShitmenpLine As String
    
    On Error GoTo Catch
    
    ddbbStock = txtPathExcelStock.value
    ddbbShitmenpLine = txtPathExcelShipmentLine.value
    
    If fso.FileExists(ddbbStock) And fso.FileExists(ddbbShitmenpLine) Then
    
        lblStatusProcess.Caption = "Espere estoy generando la lista"
        Application.Wait Now() + TimeValue("00:00:01")
        Call Index.main(ddbbStock, ddbbShitmenpLine)
        lblStatusProcess.Caption = "Lista generada"
        
    Else
        MsgBox "Asegure que las rutas sean correctas", vbExclamation, "Archivo no encontrado."
    End If
    
    Exit Sub
Catch:

    lblStatusProcess.Caption = "Opps, hubo un error"
    MsgBox Err.description, vbCritical, Err.Number
    
End Sub

Private Sub UserForm_Initialize()

    Me.Caption = "Generar lista de licuado"
    Me.frameMain.Caption = Empty
    Me.lblStatusProcess.Caption = Now()
    Me.btnExplorer1.ControlTipText = "Buscar archivo"
    Me.btnExplorer2.ControlTipText = "Buscar archivo"
    
    Call style.stylefrmMain(Me)
    Call utils.addNewSheet
    
End Sub
