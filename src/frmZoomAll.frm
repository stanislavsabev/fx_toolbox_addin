VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZoomAll 
   Caption         =   "ZoomAll"
   ClientHeight    =   1440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2784
   OleObjectBlob   =   "frmZoomAll.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZoomAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    StartInCenter Me
    txtZoomVal = ActiveWindow.Zoom
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnMinus_Click()
    On Error GoTo Failed:
    If Validate(CLng(txtZoomVal.Value)) Then
        txtZoomVal = WorksheetFunction.Max(CLng(txtZoomVal) - 5, 10)
    End If
Failed:
End Sub

Private Sub btnPlus_Click()
    On Error GoTo Failed:
    If Validate(CLng(txtZoomVal.Value)) Then
        txtZoomVal = WorksheetFunction.Min(CLng(txtZoomVal) + 5, 400)
    End If
Failed:
End Sub

Private Sub btnZoomAll_Click()
    On Error GoTo Failed
    If Validate(CLng(txtZoomVal.Value)) Then
        mWindowManager.ZoomWindow CLng(txtZoomVal.Value)
        Unload Me
    End If
Failed:
End Sub
