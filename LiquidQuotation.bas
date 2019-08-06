Attribute VB_Name = "LiquidQuotation"
Public BPrice As Double
Public PSize As Integer ' Part Number Array Size
Public userSelect As String
Public labour As Double
Public labFlag As Boolean
Public laberr As String


Function LabPrice(ByVal Repair As String)

Dim ws As Worksheet
Dim lRow As Long


Set ws = Worksheets("Countries")

lRow = ws.Range("J3").End(xlDown).Row
On Error GoTo errhandler
LabPrice = Application.WorksheetFunction.VLookup(Repair, ws.Range("J3:K" & lRow), 2, 0)

If Application.WorksheetFunction.VLookup(Repair, ws.Range("J3:K" & lRow), 2, 0) = "" Then
GoTo handler
End If




Exit Function

errhandler:
If labFlag = False Then
laberr = laberr & Repair & ", "
End If
Exit Function

handler:

If labFlag = False Then
laberr = laberr & Repair & " "
End If

End Function




' Author: Brandon Bwanakocha
' Purpose: Calculates Selling Price from given Cost price
' Parameters: CPrice - Cost Price of Terminal
' Return Type: Double

Function SPrice(ByVal CPrice)

Dim Shipping As Double
Dim Duty As Double
Dim PSD As Double ' Price with Shipping and Duty included
Dim Country As String
Dim mySelect As Range
Dim Margin As Double
Dim lRow As Long


lRow = Worksheets("Countries").Range("A3").End(xlDown).Row

Set mySelect = Worksheets("Countries").Range("A3:E" & lRow)

Country = Worksheets("Quote").Range("E5")
Shipping = Application.WorksheetFunction.VLookup(Country, mySelect, 3, 0)
Duty = Application.WorksheetFunction.VLookup(Country, mySelect, 4, 0)
Margin = Application.WorksheetFunction.VLookup(Country, mySelect, 5, 0)

Shipping = (Shipping) * CPrice
Duty = (Duty) * (Shipping + CPrice)
PSD = CPrice + Shipping + Duty

SPrice = PSD / (1 - Margin)


End Function

' Author: Brandon Bwanakocha
' Purpose: Fetches the buying price for the given terminal type from the "Countries' Column
' Parameters: Void
'Return Type : Void

Function BuyingPrice()

Dim ModelNo As String
Dim lRow As Long
Dim ws As Worksheet

Set ws = Worksheets("Countries")

lRow = ws.Range("M3").End(xlDown).Row

ModelNo = ActiveCell.Offset(0, -1).Value
On Error GoTo errhandler
BPrice = Application.WorksheetFunction.VLookup(ModelNo, ws.Range("M3:N" & lRow), 2, 0)

If Application.WorksheetFunction.VLookup(ModelNo, ws.Range("M3:N" & lRow), 2, 0) = "" Then
BPrice = 450
End If

Exit Function

errhandler:
BPrice = 450
End Function

' Author: Brandon Bwanakocha
' Purpose: Looks for multiple matches of a string in the second column of the master sheet
' Parameters: str - string we are looking for
' Return type: str - a string with row indexes of occurances of matches

Function look(str As String)
Dim rng As Range
Dim temp As String
Dim temp1 As String
Dim rows() As String
Dim size As Integer
Dim i As Integer
Dim PNumbers As String
Dim parts() As String
Dim s As String
Dim mystart As String
Dim myend As String
Dim lRow As Long
Dim ws As Worksheet

labour = 0
labFlag = False


Set ws = Worksheets("Master")
mystart = "B5"

PSize = 0

temp = ""
temp1 = ""

    lRow = ws.Range("B6").End(xlDown).Row
 
                    
    myend = "B" & lRow


For Each cell In Worksheets("Master").Range(mystart & ":" & myend)
If InStr(cell.Value, " - ") Then

parts = Split(Trim(cell.Value), " - ")
If UBound(parts()) > 1 Then
If LCase(parts(1)) = LCase(ActiveCell.Offset(0, -1).Value) Then
If LCase(parts(2)) = LCase(str) Then

temp = temp & CStr(cell.Row) & " "

labour = LabPrice(parts(2))
labFlag = True
End If
End If
End If
End If

cont:

   Next cell

Do
  temp1 = temp
  temp = Replace(temp, "  ", " ") 'remove multiple white spaces
Loop Until temp1 = temp

rows = Split(Trim(temp), " ")
size = UBound(rows()) + 1


For i = 0 To size - 1
PNumbers = PNumbers & Application.WorksheetFunction.Index(Worksheets("Master").Range("A1:A" & lRow), CInt(rows(i))) & " "
Next i

PSize = size
look = PNumbers
End Function


' Author: Brandon Bwanakocha
' Purpose: Prepares our popup window which we use to select in case of multiple occurances of the same description
' Parameters: str - string containing all part numbers that correspond to the same description
'             cap - Whatever is going to display on the label
' Return type: void

Function popupConfig(ByVal str As String, ByVal cap As String)

    Dim OptionList(0 To 50) As String
    Dim btn As CommandButton
    Set btn = LiquidForm1.CommandButton1
    Dim opt As Control
    Dim s As Variant
    Dim i As Integer
    Dim PlaceHolder
    Dim PartNumbers() As String

    
    
    PartNumbers = Split(Trim(str), " ")
    i = 0

    For i = 0 To PSize - 1
    OptionList(i) = Trim(PartNumbers(i))
    
    Next i
    
    i = 0

    For Each s In OptionList
   If i < PSize Then
        Set opt = LiquidForm1.Controls.Add("Forms.OptionButton.1", "radioBtn" & i, True)
        
        opt.Caption = s
        opt.FontSize = 8
        opt.Height = 20
        opt.Top = opt.Height * i + 20
        opt.GroupName = "Options"
        opt.Left = 10
        LiquidForm1.Height = opt.Height * (i + 2) + 20
        opt.SpecialEffect = 0
        
        
        End If

        i = i + 1
       
    Next
    
        LiquidForm1.Width = opt.Width
        
        LiquidForm1.Label1.BackColor = RGB(231, 245, 251)
        LiquidForm1.Label1.BorderColor = RGB(134, 191, 160)
        LiquidForm1.Label1.Caption = Trim(cap)
        

    btn.Caption = "Submit"
    
    btn.Top = LiquidForm1.Height - btn.Height + (0.5 * opt.Height)
    btn.Left = (LiquidForm1.Width * 0.5) - (btn.Width * 0.5) - 3

   

    LiquidForm1.Height = LiquidForm1.Height + btn.Height + (1.2 * opt.Height)
  

End Function

'Author: Brandon Bwanakocha
'Purpose: Adds an extension to the terminal type name
'Parameters: Void
'Return type: String

Function TerminalType()
Dim str As String
Dim ModelNo As String

ModelNo = ActiveCell.Offset(0, -1).Value

Select Case ModelNo

Case "S900"
str = "PAX - " & ModelNo & " - "



Case "S920"
str = "PAX - " & ModelNo & " - "


Case "S300"
str = "PAX - " & ModelNo & " - "

Case "D200"
str = "PAX - " & ModelNo & " - "

Case "D180"
str = "PAX - " & ModelNo & " - "

End Select

TerminalType = str
End Function



 'Author: Brandon Bwanakocha
 'Purpose: Finds Second Column on table
 'Parameters: void
 'Return Type: Void
 
 Function FindSecondColumn()
 
 On Error Resume Next
 If Not (ActiveSheet.ListObjects(1).ListColumns = 2) Then
 On Error Resume Next
 ActiveSheet.ListObjects(1).DataBodyRange(1, 2).Select
 End If
 
 End Function



'Author: Brandon Bwanakocha
'Purpose: Runs down the current Column until it finds empty Cell
'Parameters: Void
'Return type: Void

Function GoDownRows()
Do While Not (ActiveCell.Offset(i, 0).Value = "")
 i = i + 1
Loop
ActiveCell.Offset(i, 0).Select

' Check whether current row is outside the table and add a row if so
On Error Resume Next
If Intersect(ActiveCell, ActiveSheet.ListObjects(1).DataBodyRange) Is Nothing Then
      ActiveSheet.ListObjects(1).ListRows.Add
      ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(-1, -1).Value
       
End If

End Function




' Author: Brandon Bwanakocha
' Title: Liquid Parts - LParts When Referencing
' Purpose: Looks Up Part Numbers depending on Name of part (Yikes)
' Parameters: CellRef - The cell containing repair parts
'           : mySelect - Selected range of cells with Key words and Numbers



Function LParts(CellRef As Range)
Dim PartName As String
Dim Result() As String
Dim str As String
Dim errstr As String
Dim size As Integer
Dim count As Integer
Dim temp As String
Dim PlaceHolder



PartName = CellRef

Result = Split(Trim(PartName), ",") ' Split the part name on commas
size = UBound(Result()) + 1

str = ""
errstr = ""
laberr = ""

For count = 0 To size - 1

On Error GoTo errhandler ' In case of error, catch the error!
temp = look(Trim(Result(count)))
If temp = "" Then
GoTo handler
End If

If PSize > 1 Then

PlaceHolder = popupConfig(Trim(temp), Trim(Result(count)))
LiquidForm1.Show
str = str & Trim(CStr(userSelect)) & Chr(10)
Else
 
 str = str & temp & Chr(10)
End If


PSize = 0
counter:
        Next count

If Right(str, 1) = vbLf Then str = Left(str, Len(str) - 1) ' Remove the last new line character





LParts = Trim(str)


If Not (errstr = "") Then
MsgBox "No Part Number found for: " & errstr, vbExclamation, "Liquid Form"
End If

If Not (laberr = "") Then
MsgBox "No Labour price found for: " & laberr, vbCritical
End If


Exit Function
' Error Handler which exercutes when we encounter an error with our lookup methods
handler:

        ' Try using a different extension like "S900/S920" amd of that doesn't work then quit lmao
        
        str = str + "NPN" + Chr(10)
        errstr = errstr + Trim(Result(count)) + ", "
        GoTo counter ' Go back to the loop
errhandler:
        str = str + "NPN" + Chr(10)
        errstr = errstr + Trim(Result(count)) + ", "
        Resume counter ' Go back to the loop



End Function

' Author: Brandon Bwanakocha
' Title: Liquid Price, LPrice (When referencing)
' Date: 27/06/2019
' Purpose: Calculates the charge for a repair based on extra materials added
' Parameters: CellRef - Cell containing Part Numbers
              ' mySelect: Table Array containing prices per part Number
              

Function LPrice(CellRef As Range)
Dim Result() As String
Dim PartNumber As String

Dim count As Integer
Dim size As Integer
Dim Charge As Double
Dim lookupval As Long
Dim temp As String
Dim errstr As String
Dim lRow As Long


lRow = Worksheets("Master").Range("B5").End(xlDown).Row

errstr = ""


PartNumber = CellRef

' Remove all the extra blank spaces that may be entered by user
PartNumber = Replace(PartNumber, vbLf, " ")

PartNumber = Replace(PartNumber, "NPN", "")
PartNumber = Trim(PartNumber)



Do
  temp = PartNumber
  PartNumber = Replace(PartNumber, "  ", " ") 'remove multiple white spaces
Loop Until temp = PartNumber


Result = Split(Trim(PartNumber), " ") ' Trim to remove begining and rear spaces
size = UBound(Result()) + 1
Charge = 0

'Iterate through single cell checking all part numbers
For count = 0 To size - 1
On Error GoTo handler
Charge = Charge + SPrice(Application.WorksheetFunction.VLookup(Trim(Result(count)), Worksheets("Master").Range("A5:I" & lRow), 9, 0)) + labour  ' Lookup spare part price and add that to total charge on customer
If Application.WorksheetFunction.VLookup(Trim(Result(count)), Worksheets("Master").Range("A5:I" & lRow), 9, 0) = "" Then
errstr = Result(count)
End If

counter:
      Next count


If size < 3 Then
  LPrice = 12.5 + Charge ' If we used less than 3 extra parts then charge Labour
Else
 LPrice = Charge  ' Charge for just the extra parts if we used more than three spare parts
End If

If Not (errstr = "") Then
MsgBox "No price found for: " & errstr, vbInformation, "Liquid Form"
End If

Exit Function
handler:
        errstr = errstr + Result(count) + ","
       ' Charge = IIf(size > 3, Charge, Charge + 12.5)
        Resume counter

End Function

'Author: Unknown
'Source: Stack Overflow
'Purpose: Checks if User form is still open

'Private Function IsLoaded(ByVal formName As String) As Boolean
 '   Dim frm As Object
  '  For Each frm In VBA.UserForms
   '     If frm.Name = formName Then
    '        IsLoaded = True
     '       Exit Function
      '  End If
    'Nex 't frm
    'IsLoaded = False
'End Function

' Author: Brandon Bwanakocha
' Title: getStr
'Purpose: To get the strings that are contained in the text boxes and combo boxes
'Parameters: str - string
'            str_1 -  string
'            trig - Boolean boolean flag
'            trig_1 - Boolean flag

Function getStr()
Dim str As String
Dim str_1 As String
Dim trig As Boolean
Dim trig_1 As Boolean
Dim boxstr As String
Dim box2str As String
Dim box3str As String


str = ""
str_1 = ""
trig = True
trig_1 = True
box2str = ""
box3str = ""


Dim Result(2) As String

    For Each cCont In LiquidForm.Controls

        If TypeName(cCont) = "ComboBox" Then
           If cCont.Text = "" Then
           ' Do nothing if the ComboBox contains nothing
           Else

            If trig = False Then ' Make sure you ommit the first comma
            
            Select Case cCont.Name
            Case "ComboBox1"
            str = str + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox2"
            str = str + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox3"
            str = str + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox4"
            str = str + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox5"
            str = str + ", " + Trim(CStr(cCont.Value))
            
            End Select
            
            
            Else
            
            Select Case cCont.Name
            Case "ComboBox1"
            str = str + Trim(CStr(cCont.Value))
            
            Case "ComboBox2"
            str = str + Trim(CStr(cCont.Value))
            
            Case "ComboBox3"
            str = str + Trim(CStr(cCont.Value))
            
            Case "ComboBox4"
            str = str + Trim(CStr(cCont.Value))
            
            Case "ComboBox5"
            str = str + Trim(CStr(cCont.Value))
            
            End Select
            trig = False
            End If
            
            If trig_1 = False Then
            
            Select Case cCont.Name
            Case "ComboBox6"
            str_1 = str_1 + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox7"
            str_1 = str_1 + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox8"
            str_1 = str_1 + ", " + Trim(CStr(cCont.Value))
           
            Case "ComboBox9"
            str_1 = str_1 + ", " + Trim(CStr(cCont.Value))
            
            Case "ComboBox10"
            str_1 = str_1 + ", " + Trim(CStr(cCont.Value))
            
            End Select
            
            Else
            
            Select Case cCont.Name
            Case "ComboBox6"
            trig_1 = False
            str_1 = str_1 + Trim(CStr(cCont.Value))
            
            Case "ComboBox7"
            trig_1 = False
            str_1 = str_1 + Trim(CStr(cCont.Value))
            
            trig_1 = False
            Case "ComboBox8"
            str_1 = str_1 + Trim(CStr(cCont.Value))
            
            trig_1 = False
            Case "ComboBox9"
            trig_1 = False
            str_1 = str_1 + Trim(CStr(cCont.Value))
            
            Case "ComboBox10"
            trig_1 = False
            str_1 = str_1 + Trim(CStr(cCont.Value))
            
            End Select
            
            End If
             
        End If
        End If
        

     Next cCont
    
    ' Fetch contents from the Other text boxes
    If Not (LiquidForm.TextBox2.Value = "") Then
     Select Case trig
     Case False
     box2str = ", " + LiquidForm.TextBox2.Value
     Case True
     box2str = LiquidForm.TextBox2.Value
     End Select
     
     End If
     
     If Not (LiquidForm.TextBox3.Value = "") Then
    Select Case trig_1
     Case False
     box3str = ", " + LiquidForm.TextBox3.Value
     Case True
     box3str = LiquidForm.TextBox3.Value
     End Select
     
     End If
     
     
 ' -------------------------------------------------------------------------------
 str = str + Trim(box2str)
 str_1 = str_1 + Trim(box3str)
 
 Result(0) = str
 Result(1) = str_1

getStr = Result
End Function



' Author: Brandon Bwanakocha
' Purpose: This function initializes  the combo boxes when the quotation Macro is summoned
' Parameters: Void
' Return Type: Void


Sub Quotation()

Dim OpenForms
Dim PlaceHolder

BPrice = 450

' Add items to the combo boxes which are the drop down menus.

With LiquidForm.ComboBox1
.AddItem "Battery cover damaged"
.AddItem "Battery cover missing"
.AddItem "Battery faulty"
.AddItem "Battery missing"
.AddItem "Battery swollen"
.AddItem "Charging port damaged"
.AddItem "Charging port faulty"
.AddItem "Chip card reader damaged"
.AddItem "Chip card reader faulty"
.AddItem "CMOS battery low"
.AddItem "Bottom Case damaged"
.AddItem "Top Case damaged"
.AddItem "I/O board damaged"
.AddItem "I/O board faulty"
.AddItem "Keypad display damaged"
.AddItem "Keypad display unresponsive"
.AddItem "Keypad main damaged"
.AddItem "Keypad main unresponsive"
.AddItem "LCD damaged"
.AddItem "LCD faulty"
.AddItem "LCD no display"
.AddItem "Lens damaged"
.AddItem "Mag card reader damaged"
.AddItem "Mag card reader faulty"
.AddItem "Mainboard damaged"
.AddItem "Mainboard faulty"
.AddItem "Network failure"
.AddItem "NFC board damaged"
.AddItem "NFC board faulty"
.AddItem "NFC light on "
.AddItem "No fault found"
.AddItem "Not charging"
.AddItem "Not powering on"
.AddItem "Power adapter damaged"
.AddItem "Power adapter missing"
.AddItem "Printer cover damaged"
.AddItem "Printer cover missing"
.AddItem "Printer handle damaged"
.AddItem "Printer handle missing"
.AddItem "Printer roller damaged"
.AddItem "Printer roller missing"
.AddItem "Printer unit damaged"
.AddItem "Printer unit faulty"
.AddItem "SIM card holder damaged"
.AddItem "SIM card holder missing"
.AddItem "Software call customer service"
.AddItem "Software fault"
.AddItem "Software no go variable"
.AddItem "Substance damaged "
.AddItem "Tamper hard"
.AddItem "Tamper soft"

End With

With LiquidForm.ComboBox2
.AddItem "Battery cover damaged"
.AddItem "Battery cover missing"
.AddItem "Battery faulty"
.AddItem "Battery missing"
.AddItem "Battery swollen"
.AddItem "Charging port damaged"
.AddItem "Charging port faulty"
.AddItem "Chip card reader damaged"
.AddItem "Chip card reader faulty"
.AddItem "CMOS battery low"
.AddItem "Bottom Case damaged"
.AddItem "Mainboard faulty"
.AddItem "Network failure"
.AddItem "NFC board damaged"
.AddItem "NFC board faulty"
.AddItem "NFC light on "
.AddItem "No fault found"
.AddItem "Not charging"
.AddItem "Not powering on"
.AddItem "Power adapter damaged"
.AddItem "Power adapter missing"
.AddItem "Printer cover damaged"
.AddItem "Printer cover missing"
.AddItem "Printer handle damaged"
.AddItem "Printer handle missing"
.AddItem "Printer roller damaged"
.AddItem "Printer roller missing"
.AddItem "Printer unit damaged"
.AddItem "Printer unit faulty"
.AddItem "SIM card holder damaged"
.AddItem "SIM card holder missing"
.AddItem "Software call customer service"
.AddItem "Software fault"
.AddItem "Software no go variable"
.AddItem "Substance damaged "
.AddItem "Tamper hard"
.AddItem "Tamper soft"

End With

With LiquidForm.ComboBox3
.AddItem "Battery cover damaged"
.AddItem "Battery cover missing"
.AddItem "Battery faulty"
.AddItem "Battery missing"
.AddItem "Battery swollen"
.AddItem "Charging port damaged"
.AddItem "Charging port faulty"
.AddItem "Chip card reader damaged"
.AddItem "Chip card reader faulty"
.AddItem "CMOS battery low"
.AddItem "Bottom Case damaged"
.AddItem "Top Case damaged"
.AddItem "I/O board damaged"
.AddItem "I/O board faulty"
.AddItem "Keypad display damaged"
.AddItem "Keypad display unresponsive"
.AddItem "Keypad main damaged"
.AddItem "Keypad main unresponsive"
.AddItem "LCD damaged"
.AddItem "LCD faulty"
.AddItem "LCD no display"
.AddItem "Lens damaged"
.AddItem "Mag card reader damaged"
.AddItem "Mag card reader faulty"
.AddItem "Mainboard damaged"
.AddItem "Mainboard faulty"
.AddItem "Network failure"
.AddItem "NFC board damaged"
.AddItem "NFC board faulty"
.AddItem "NFC light on "
.AddItem "No fault found"
.AddItem "Not charging"
.AddItem "Not powering on"
.AddItem "Power adapter damaged"
.AddItem "Power adapter missing"
.AddItem "Printer cover damaged"
.AddItem "Printer cover missing"
.AddItem "Printer handle damaged"
.AddItem "Printer handle missing"
.AddItem "Printer roller damaged"
.AddItem "Printer roller missing"
.AddItem "Printer unit damaged"
.AddItem "Printer unit faulty"
.AddItem "SIM card holder damaged"
.AddItem "SIM card holder missing"
.AddItem "Software call customer service"
.AddItem "Software fault"
.AddItem "Software no go variable"
.AddItem "Substance damaged "
.AddItem "Tamper hard"
.AddItem "Tamper soft"

End With

With LiquidForm.ComboBox4
.AddItem "Battery cover damaged"
.AddItem "Battery cover missing"
.AddItem "Battery faulty"
.AddItem "Battery missing"
.AddItem "Battery swollen"
.AddItem "Charging port damaged"
.AddItem "Charging port faulty"
.AddItem "Chip card reader damaged"
.AddItem "Chip card reader faulty"
.AddItem "CMOS battery low"
.AddItem "Bottom Case damaged"
.AddItem "Top Case damaged"
.AddItem "I/O board damaged"
.AddItem "I/O board faulty"
.AddItem "Keypad display damaged"
.AddItem "Keypad display unresponsive"
.AddItem "Keypad main damaged"
.AddItem "Keypad main unresponsive"
.AddItem "LCD damaged"
.AddItem "LCD faulty"
.AddItem "LCD no display"
.AddItem "Lens damaged"
.AddItem "Mag card reader damaged"
.AddItem "Mag card reader faulty"
.AddItem "Mainboard damaged"
.AddItem "Mainboard faulty"
.AddItem "Network failure"
.AddItem "NFC board damaged"
.AddItem "NFC board faulty"
.AddItem "NFC light on "
.AddItem "No fault found"
.AddItem "Not charging"
.AddItem "Not powering on"
.AddItem "Power adapter damaged"
.AddItem "Power adapter missing"
.AddItem "Printer cover damaged"
.AddItem "Printer cover missing"
.AddItem "Printer handle damaged"
.AddItem "Printer handle missing"
.AddItem "Printer roller damaged"
.AddItem "Printer roller missing"
.AddItem "Printer unit damaged"
.AddItem "Printer unit faulty"
.AddItem "SIM card holder damaged"
.AddItem "SIM card holder missing"
.AddItem "Software call customer service"
.AddItem "Software fault"
.AddItem "Software no go variable"
.AddItem "Substance damaged "
.AddItem "Tamper hard"
.AddItem "Tamper soft"

End With

With LiquidForm.ComboBox5
.AddItem "Battery cover damaged"
.AddItem "Battery cover missing"
.AddItem "Battery faulty"
.AddItem "Battery missing"
.AddItem "Battery swollen"
.AddItem "Charging port damaged"
.AddItem "Charging port faulty"
.AddItem "Chip card reader damaged"
.AddItem "Chip card reader faulty"
.AddItem "CMOS battery low"
.AddItem "Bottom Case damaged"
.AddItem "Top Case damaged"
.AddItem "I/O board damaged"
.AddItem "I/O board faulty"
.AddItem "Keypad display damaged"
.AddItem "Keypad display unresponsive"
.AddItem "Keypad main damaged"
.AddItem "Keypad main unresponsive"
.AddItem "LCD damaged"
.AddItem "LCD faulty"
.AddItem "LCD no display"
.AddItem "Lens damaged"
.AddItem "Mag card reader damaged"
.AddItem "Mag card reader faulty"
.AddItem "Mainboard damaged"
.AddItem "Mainboard faulty"
.AddItem "Network failure"
.AddItem "NFC board damaged"
.AddItem "NFC board faulty"
.AddItem "NFC light on "
.AddItem "No fault found"
.AddItem "Not charging"
.AddItem "Not powering on"
.AddItem "Power adapter damaged"
.AddItem "Power adapter missing"
.AddItem "Printer cover damaged"
.AddItem "Printer cover missing"
.AddItem "Printer handle damaged"
.AddItem "Printer handle missing"
.AddItem "Printer roller damaged"
.AddItem "Printer roller missing"
.AddItem "Printer unit damaged"
.AddItem "Printer unit faulty"
.AddItem "SIM card holder damaged"
.AddItem "SIM card holder missing"
.AddItem "Software call customer service"
.AddItem "Software fault"
.AddItem "Software no go variable"
.AddItem "Substance damaged "
.AddItem "Tamper hard"
.AddItem "Tamper soft"

End With

With LiquidForm.ComboBox6
.AddItem "Antenna"
.AddItem "Battery coin"
.AddItem "Battery connector"
.AddItem "Battery cover"
.AddItem "Battery rechargeable"
.AddItem "Board I/O"
.AddItem "Board main"
.AddItem "Board NFC"
.AddItem "Bottom Case"
.AddItem "Bottom Cover Rubber Feet"
.AddItem "Carbon"
.AddItem "Card reader chip"
.AddItem "Charge Connecter"
.AddItem "Dome Pack"
.AddItem "ESD Shield"
.AddItem "Flat"
.AddItem "Keymesh"
.AddItem "Keypad display"
.AddItem "Keypad main"
.AddItem "Labour 15 mins"
.AddItem "Labour 30 mins"
.AddItem "Labour 45 mins"
.AddItem "Labour reload software"
.AddItem "Labour repair"
.AddItem "Labour strip and clean"
.AddItem "Labour tamper reset"
.AddItem "LCD"
.AddItem "Lens standard"
.AddItem "Lens touch"
.AddItem "MAG Bracket"
.AddItem "Magnetic Card Reader"
.AddItem "Metal paper cutter"
.AddItem "Phillips"
.AddItem "Power adapter"
.AddItem "Power cable"
.AddItem "Power connector"
.AddItem "Printer assembly"
.AddItem "Printer bracket"
.AddItem "Printer Cover Screw Set"
.AddItem "Printer cover"
.AddItem "Printer handle"
.AddItem "Printer roller"
.AddItem "Printer unit"
.AddItem "Printer"
.AddItem "Rechargable Battery"
.AddItem "RF Antenna"
.AddItem "Rubber foot"
.AddItem "Rubber switch"
.AddItem "SAM Board"
.AddItem "SAM Cover"
.AddItem "Top Case"
.AddItem "Touch Lens"
.AddItem "USB port"
.AddItem "USB to Micro Charge"
.AddItem "Wireless Antenna"
.AddItem "Zebra block"


End With

With LiquidForm.ComboBox7
.AddItem "Antenna"
.AddItem "Battery coin"
.AddItem "Battery connector"
.AddItem "Battery cover"
.AddItem "Battery rechargeable"
.AddItem "Board I/O"
.AddItem "Board main"
.AddItem "Board NFC"
.AddItem "Bottom Case"
.AddItem "Bottom Cover Rubber Feet"
.AddItem "Carbon"
.AddItem "Card reader chip"
.AddItem "Charge Connecter"
.AddItem "Dome Pack"
.AddItem "ESD Shield"
.AddItem "Flat"
.AddItem "Keymesh"
.AddItem "Keypad display"
.AddItem "Keypad main"
.AddItem "Labour 15 mins"
.AddItem "Labour 30 mins"
.AddItem "Labour 45 mins"
.AddItem "Labour reload software"
.AddItem "Labour repair"
.AddItem "Labour strip and clean"
.AddItem "Labour tamper reset"
.AddItem "LCD"
.AddItem "Lens standard"
.AddItem "Lens touch"
.AddItem "MAG Bracket"
.AddItem "Magnetic Card Reader"
.AddItem "Metal paper cutter"
.AddItem "Phillips"
.AddItem "Power adapter"
.AddItem "Power cable"
.AddItem "Power connector"
.AddItem "Printer assembly"
.AddItem "Printer bracket"
.AddItem "Printer Cover Screw Set"
.AddItem "Printer cover"
.AddItem "Printer handle"
.AddItem "Printer roller"
.AddItem "Printer unit"
.AddItem "Printer"
.AddItem "Rechargable Battery"
.AddItem "RF Antenna"
.AddItem "Rubber foot"
.AddItem "Rubber switch"
.AddItem "SAM Board"
.AddItem "SAM Cover"
.AddItem "Top Case"
.AddItem "Touch Lens"
.AddItem "USB port"
.AddItem "USB to Micro Charge"
.AddItem "Wireless Antenna"
.AddItem "Zebra block"

End With

With LiquidForm.ComboBox8
.AddItem "Antenna"
.AddItem "Battery coin"
.AddItem "Battery connector"
.AddItem "Battery cover"
.AddItem "Battery rechargeable"
.AddItem "Board I/O"
.AddItem "Board main"
.AddItem "Board NFC"
.AddItem "Bottom Case"
.AddItem "Bottom Cover Rubber Feet"
.AddItem "Carbon"
.AddItem "Card reader chip"
.AddItem "Charge Connecter"
.AddItem "Dome Pack"
.AddItem "ESD Shield"
.AddItem "Flat"
.AddItem "Keymesh"
.AddItem "Keypad display"
.AddItem "Keypad main"
.AddItem "Labour 15 mins"
.AddItem "Labour 30 mins"
.AddItem "Labour 45 mins"
.AddItem "Labour reload software"
.AddItem "Labour repair"
.AddItem "Labour strip and clean"
.AddItem "Labour tamper reset"
.AddItem "LCD"
.AddItem "Lens standard"
.AddItem "Lens touch"
.AddItem "MAG Bracket"
.AddItem "Magnetic Card Reader"
.AddItem "Metal paper cutter"
.AddItem "Phillips"
.AddItem "Power adapter"
.AddItem "Power cable"
.AddItem "Power connector"
.AddItem "Printer assembly"
.AddItem "Printer bracket"
.AddItem "Printer Cover Screw Set"
.AddItem "Printer cover"
.AddItem "Printer handle"
.AddItem "Printer roller"
.AddItem "Printer unit"
.AddItem "Printer"
.AddItem "Rechargable Battery"
.AddItem "RF Antenna"
.AddItem "Rubber foot"
.AddItem "Rubber switch"
.AddItem "SAM Board"
.AddItem "SAM Cover"
.AddItem "Top Case"
.AddItem "Touch Lens"
.AddItem "USB port"
.AddItem "USB to Micro Charge"
.AddItem "Wireless Antenna"
.AddItem "Zebra block"


End With

With LiquidForm.ComboBox9
.AddItem "Antenna"
.AddItem "Battery coin"
.AddItem "Battery connector"
.AddItem "Battery cover"
.AddItem "Battery rechargeable"
.AddItem "Board I/O"
.AddItem "Board main"
.AddItem "Board NFC"
.AddItem "Bottom Case"
.AddItem "Bottom Cover Rubber Feet"
.AddItem "Carbon"
.AddItem "Card reader chip"
.AddItem "Charge Connecter"
.AddItem "Dome Pack"
.AddItem "ESD Shield"
.AddItem "Flat"
.AddItem "Keymesh"
.AddItem "Keypad display"
.AddItem "Keypad main"
.AddItem "Labour 15 mins"
.AddItem "Labour 30 mins"
.AddItem "Labour 45 mins"
.AddItem "Labour reload software"
.AddItem "Labour repair"
.AddItem "Labour strip and clean"
.AddItem "Labour tamper reset"
.AddItem "LCD"
.AddItem "Lens standard"
.AddItem "Lens touch"
.AddItem "MAG Bracket"
.AddItem "Magnetic Card Reader"
.AddItem "Metal paper cutter"
.AddItem "Phillips"
.AddItem "Power adapter"
.AddItem "Power cable"
.AddItem "Power connector"
.AddItem "Printer assembly"
.AddItem "Printer bracket"
.AddItem "Printer Cover Screw Set"
.AddItem "Printer cover"
.AddItem "Printer handle"
.AddItem "Printer roller"
.AddItem "Printer unit"
.AddItem "Printer"
.AddItem "Rechargable Battery"
.AddItem "RF Antenna"
.AddItem "Rubber foot"
.AddItem "Rubber switch"
.AddItem "SAM Board"
.AddItem "SAM Cover"
.AddItem "Top Case"
.AddItem "Touch Lens"
.AddItem "USB port"
.AddItem "USB to Micro Charge"
.AddItem "Wireless Antenna"
.AddItem "Zebra block"

End With

With LiquidForm.ComboBox10
.AddItem "Antenna"
.AddItem "Battery coin"
.AddItem "Battery connector"
.AddItem "Battery cover"
.AddItem "Battery rechargeable"
.AddItem "Board I/O"
.AddItem "Board main"
.AddItem "Board NFC"
.AddItem "Bottom Case"
.AddItem "Bottom Cover Rubber Feet"
.AddItem "Carbon"
.AddItem "Card reader chip"
.AddItem "Charge Connecter"
.AddItem "Dome Pack"
.AddItem "ESD Shield"
.AddItem "Flat"
.AddItem "Keymesh"
.AddItem "Keypad display"
.AddItem "Keypad main"
.AddItem "Labour 15 mins"
.AddItem "Labour 30 mins"
.AddItem "Labour 45 mins"
.AddItem "Labour reload software"
.AddItem "Labour repair"
.AddItem "Labour strip and clean"
.AddItem "Labour tamper reset"
.AddItem "LCD"
.AddItem "Lens standard"
.AddItem "Lens touch"
.AddItem "MAG Bracket"
.AddItem "Magnetic Card Reader"
.AddItem "Metal paper cutter"
.AddItem "Phillips"
.AddItem "Power adapter"
.AddItem "Power cable"
.AddItem "Power connector"
.AddItem "Printer assembly"
.AddItem "Printer bracket"
.AddItem "Printer Cover Screw Set"
.AddItem "Printer cover"
.AddItem "Printer handle"
.AddItem "Printer roller"
.AddItem "Printer unit"
.AddItem "Printer"
.AddItem "Rechargable Battery"
.AddItem "RF Antenna"
.AddItem "Rubber foot"
.AddItem "Rubber switch"
.AddItem "SAM Board"
.AddItem "SAM Cover"
.AddItem "Top Case"
.AddItem "Touch Lens"
.AddItem "USB port"
.AddItem "USB to Micro Charge"
.AddItem "Wireless Antenna"
.AddItem "Zebra block"

End With

'LiquidForm1.Show vbModeless

'Do While IsLoaded("LiquidForm1")
 '   OpenForms = DoEvents() 'Hand control to the Operating System
'Loop
'------------------------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------------------------------------------------

LiquidForm.Show vbModeless

PlaceHolder = FindSecondColumn()
PlaceHolder = GoDownRows()

PaceHolder = BuyingPrice()
End Sub





