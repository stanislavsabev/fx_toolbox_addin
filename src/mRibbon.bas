Attribute VB_Name = "mRibbon"
Option Explicit
Private Const ModuleName = "mRibbon"

'Callback for customUI.onLoad
Sub ribbonLoaded(ribbon As IRibbonUI)
'    Call fx.Raise_NotImplementedError
End Sub

'Callback for btnCrop onAction
Sub Fx_Callback(control As IRibbonControl)
    Call Application.Run(control.Id)
End Sub
