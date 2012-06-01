Attribute VB_Name = "ERRORModule"
Option Explicit

Function msgERROR(strMessege As String, strERRcode As String)
    MsgBox (strMessege & vbCrLf & strERRcode)
End Function
