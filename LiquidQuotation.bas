Attribute VB_Name = "LiquidQuotation"
' Author: Brandon Bwanakocha
' Title: Liquid Parts - LParts When Referencing
' Purpose: Looks Up Part Numbers depending on Name of part (Yikes)
' Parameters: CellRef - The cell containing repair parts
'           : mySelect - Selected range of cells with Key words and Numbers



Function LParts(CellRef As Range, mySelect As Range)
Dim PartName As String
Dim Result() As String
Dim str As String
Dim Size As Integer
Dim count As Integer



PartName = CellRef

Result = Split(Trim(PartName), ",") ' Split the part name on commas
Size = UBound(Result()) + 1
str = ""
For count = 0 To Size - 1
On Error Resume Next
str = str & " " & (Application.WorksheetFunction.VLookup(Trim(Result(count)), mySelect, 2, 0))
Next count



LParts = Trim(str)

End Function

' Author: Brandon Bwanakocha
' Title: getStr
'Purpose: To get the strings that are contained in the text boxes and combo boxes
'Parameters: str - string
'            str_1 -  string
'            trig - Boolean boolean flag
'            trig_1 - Boolean flag

Function getStr(str As String, str_1 As String, trig As Boolean, trig_1 As Boolean)

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
' Title: Liquid Price, LPrice (When referencing)
' Date: 27/06/2019
' Purpose: Calculates the charge for a repair based on extra materials added
' Parameters: CellRef - Cell containing Part Numbers
              ' mySelect: Table Array containing prices per part Number
              

Function LPrice(CellRef As Range, mySelect As Range)
Dim Result() As String
Dim PartNumber As String

Dim count As Integer
Dim Size As Integer
Dim Charge As Double
Dim lookupval As Long
Dim temp As String


PartNumber = CellRef

' Remove all the extra blank spaces that may be entered by user

Do
  temp = PartNumber
  PartNumber = Replace(PartNumber, "  ", " ") 'remove multiple white spaces
Loop Until temp = PartNumber


Result = Split(Trim(PartNumber), " ") ' Trim to remove begining and rear spaces
Size = UBound(Result()) + 1
Charge = 0

'Iterate through single cell checking all part numbers
For count = 0 To Size - 1
On Error Resume Next
Charge = Charge + Application.WorksheetFunction.VLookup(Trim(Result(count)), mySelect, 2, 0) ' Lookup spare part price and add that to total charge on customer
Next count


If Size < 3 Then
  LPrice = 12.5 + Charge ' If we used less than 3 extra parts then charge Labour
Else
 LPrice = Charge ' Charge for just the extra parts if we used more than three spare parts
End If


End Function

' Author: Brandon Bwanakocha
' Purpose: This function initializes  the combo boxes when the quotation Macro is summoned
' Parameters: Void
' Return Type: Void


Sub Quotation()

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


LiquidForm.Show vbModeless

On Error Resume Next ' In case of error skip this line
ActiveSheet.ListObjects(1).DataBodyRange(1, 2).Select ' Select the second columb in whatever table exists on the sheet

Do While Not (ActiveCell.Value = "") ' Go down until you find an empty cell
 ActiveCell.Offset(1, 0).Select
Loop
End Sub
