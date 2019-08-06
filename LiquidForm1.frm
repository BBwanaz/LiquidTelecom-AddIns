VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LiquidForm1 
   Caption         =   "Liquid Quotation Form"
   ClientHeight    =   1220
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3190
   OleObjectBlob   =   "LiquidForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LiquidForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_BeforeDragOver(ByVal Cancel As msforms.ReturnBoolean, ByVal Data As msforms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As msforms.fmDragState, ByVal Effect As msforms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub CommandButton1_Click()
 Dim ctrl As Control


For Each ctrl In LiquidForm1.Controls

If TypeOf ctrl Is msforms.OptionButton Then

If ctrl.Value = True Then

userSelect = Trim(ctrl.Caption)

End If

End If

Next ctrl
     

    Unload LiquidForm1
End Sub



Private Sub UserForm_Click()

End Sub
