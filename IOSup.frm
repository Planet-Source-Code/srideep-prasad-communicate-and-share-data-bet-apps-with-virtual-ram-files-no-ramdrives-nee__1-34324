VERSION 5.00
Begin VB.Form IOSup 
   Caption         =   "Virtual File I/O Supervisor Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "IOSup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "IOSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ParentObj As vFile

Private Sub Form_Resize()
    ParentObj.RaiseChangeEvent GetProp(Me.hwnd, "Offset"), GetProp(Me.hwnd, "Size")
End Sub
Sub SetParams(pObj As vFile)
    Set ParentObj = pObj
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ParentObj = Nothing
End Sub
