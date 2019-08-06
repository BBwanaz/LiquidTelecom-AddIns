VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LiquidForm 
   Caption         =   "Liquid Quotation Form"
   ClientHeight    =   5430
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8140
   OleObjectBlob   =   "LiquidForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LiquidForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Author: Brandon Bwanakocha
' Purpose: This function controlls what happenes when the "Submit" Button is clicked
' Parameters: Void
' Return Type: Void

Private Sub Label2_Click()
Dim str As String   ' String for the Faults cell
Dim str_1 As String  ' String for the Repairs cell
Dim BoxName As String
Dim PartNumbers As String
Dim count As Integer
Dim cCont As Control
Dim Price_Charged As Double
Dim i As Integer
Dim strings() As String
Dim PlaceHolder



i = 0
 
' Get the strings that are contained int combo boxes and text boxes

strings = getStr()
str = strings(0)
str_1 = strings(1)

PlaceHolder = FindSecondColumn()
PlaceHolder = GoDownRows()


' If we do not have a terminal type then we do nothing but just display a reminder
If ActiveCell.Offset(0, -1) = "" Then
MsgBox "Please Enter Terminal Type", vbOKOnly + vbExclamation, "Liquid Form"
 
Else
 ' Clear contents of every object on the Form
 
   ActiveCell.Value = TextBox1.Value
   TextBox1.Value = ""
   TextBox2.Value = ""
   TextBox3.Value = ""
   ComboBox1.Value = ""
   ComboBox2.Value = ""
   ComboBox3.Value = ""
   ComboBox4.Value = ""
   ComboBox5.Value = ""
   ComboBox6.Value = ""
   ComboBox7.Value = ""
   ComboBox8.Value = ""
   ComboBox9.Value = ""
   ComboBox10.Value = ""
    
 
ActiveCell.Offset(0, 1).Value = str ' Update the value of the selected cell
ActiveCell.Offset(0, 2).Value = str_1 ' Move to the next cell

' Fetch Part Numbers
PartNumbers = LParts(ActiveCell.Offset(0, 2))

ActiveCell.Offset(0, 3).Value = CStr(PartNumbers)


If CStr(PartNumbers) = "" Then
ActiveCell.Offset(0, 3).Value = "-"
End If

' Calculate Total Price
Price_Charged = LPrice(ActiveCell.Offset(0, 3))


' Check to see if Terminal Is Beyond Repair

If Price_Charged > 0.75 * BPrice Then
MsgBox "Beyond Economic Repair", vbCritical, "Liquid Form"
ActiveCell.Offset(0, 3).Value = "BER"
ActiveCell.Offset(0, 4).Value = 0

Else
ActiveCell.Offset(0, 4).Value = Price_Charged

End If

PlaceHolder = GoDownRows()
End If
End Sub

'Author: Brandon Bwanakocha
'Purpose: This function changes the color of the Submit button when it is clicked



Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label2.BackColor = &H8000000F
Label2.BorderColor = RGB(165, 218, 240)
End Sub

'Author Brandon Bwanakocha
'Purpose: This function changes the color of the Submit Button when the Mouse hovers on it
'Parameters: Default VBA Parameters

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label2.BackColor = RGB(231, 245, 251)
Label2.BorderColor = RGB(134, 191, 160)
End Sub

' Author: Brandon Bwanakocha
' Purpose: Restores the color of the Submit Button when mouse moves up after being clicked
' Parameters: Defauld VBA Parameters
' Return Type: Void

Private Sub Label2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label2.BackColor = &H80000005
End Sub

' Author: Brandon Bwanakocha
' Purpose: This function closes our application when the "Cancel" button is pressed
' Parameters: Void
' Return Type: Void

Private Sub Label3_Click()
Label3.BackColor = &H8000000F
Label3.BorderColor = RGB(165, 218, 240)
Unload Me

End Sub

'Author Brandon Bwanakocha
'Purpose: This function changes the color of the Cancel Button when the Mouse hovers on it
'Parameters: Default VBA Parameters

Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label3.BackColor = RGB(231, 245, 251)
Label3.BorderColor = RGB(134, 191, 160)
End Sub



' Author: Brandon Bwanakocha
' Purpose: Restors the color of the Cancel Button and the Submit Button after mouse hovers anywhere else on the form
' Parameters: Default VBA Parameters
' Return Type: Void

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label2.BackColor = &H80000005
Label2.BorderColor = &HA9A9A9
Label3.BackColor = &H80000005
Label3.BorderColor = &HA9A9A9
End Sub
