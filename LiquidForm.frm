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
Dim mySelect As Range
Dim mySelect_1 As Range
Dim S900Sel As Range
Dim S900Sel_1 As Range
Dim S920Sel_1 As Range
Dim S920Sel As Range
Dim count As Integer
Dim cCont As Control
Dim trig As Boolean
Dim trig_1 As Boolean
Dim Price_Charged As Double
Dim i As Integer
Dim strings() As String

Dim box2str As String
Dim box3str As String


' Select ranges for S900 VLookup
Set S900Sel = Worksheets("Parts").Range("A4:B30") ' Part Number Selection
Set S900Sel_1 = Worksheets("Parts").Range("B4:C30") ' Price Selection

' Select ranges for S920 Vlookup
Set S920Sel = Worksheets("Parts").Range("A32:B59") ' Part Number Selection
Set S920Sel_1 = Worksheets("Parts").Range("B32:C59") ' Price Selection


str = ""
str_1 = ""
box2str = ""
box3str = ""
trig = True
trig_1 = True
i = 0
 
' Get the strings that are contained int combo boxes and text boxes

strings = getStr(str, str_1, trig, trig_1)
str = strings(0)
str_1 = strings(1)

' Go to the second column of the table
On Error Resume Next
If Not (ActiveSheet.ListObjects(1).ListColumns = 2) Then
On Error Resume Next
ActiveSheet.ListObjects(1).DataBodyRange(1, 2).Select
End If


' Find the next empty cell in the second column of the table
Do While Not (ActiveCell.Offset(i, 0).Value = "")
 i = i + 1
Loop
ActiveCell.Offset(i, 0).Select


' Check whether current row is outside the table and add a row if so

If Intersect(ActiveCell, ActiveSheet.ListObjects(1).DataBodyRange) Is Nothing Then
      ActiveSheet.ListObjects(1).ListRows.Add
      ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(-1, -1).Value
       
End If

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


' Check if S900 and see
If ActiveCell.Offset(0, -1).Value = "S900" Then
 Set mySelect = S900Sel
 Set mySelect_1 = S900Sel_1
End If

' Check if S920
If ActiveCell.Offset(0, -1).Value = "S920" Then
Set mySelect = S920Sel
Set mySelect_1 = S920Sel_1
End If

' Fetch Part Numbers
PartNumbers = LParts(ActiveCell.Offset(0, 2), mySelect)

ActiveCell.Offset(0, 3).Value = CStr(PartNumbers)

If CStr(PartNumbers) = "" Then
ActiveCell.Offset(0, 3).Value = "-"
End If

' Calculate Total Price
Price_Charged = LPrice(ActiveCell.Offset(0, 3), mySelect_1)
ActiveCell.Offset(0, 4).Value = Price_Charged

Do
ActiveCell.Offset(1, 0).Select ' Go to the next empty cell on second column in whatever object we have
Loop While Not (ActiveCell.Value = "")
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
