# LiquidTelecom-AddIns


The features include:
 
    - Locating the first empty cell on the Serial Numbers Column. (this feature also works with a partially competed document.)
    - Fetching Serial Numbers, Calculating the price to be charged on the customer while putting into consideration the condition by which  we charge for labour.
    - Locating the first empty cell when SUBMIT is pressed
    - Creating a new Row if SUBMIT is pressed when the last row in the Table is occupied.
    
NOTE: This addIn has to be installed into your excel and you also need to create Net Tab where you will add the "Quotation" Macro on the ribbons. This feature also needs you to have your "Price Lookup" Sheet in the same sheet as the quotation otherwise edit the code as required.
# Installing the AddIn

Download the xlam file and note your download file location. 
Open Excel: -> File -> Options - > AddIns
On the dragdown at the bottom and select "Exel Addins" and click "GO"
Browse your computer for the xlam (Recall the location) and open it. 
Add a checkmark to the addin.

Now the addin is in your excel and you can use it with any document. However, you need an icon on the tool bar that you click when you want to use the Quotation Form Application. In order to do that, follow the following tutorial on how to add new tabs on your excel: https://www.homeandlearn.org/customize_the_excel_ribbon.html . Our User Form is called "Quotation" so instead of "CallUserForm" look for "Quotation"

# Editing the AddIn

Please see this tutorial on how to add the "Developer" Tab on your Excel and Access the files. https://youtu.be/JLQ8OuW0FlY. After adding the Developer Tab, click on "Visual Basic" and And a new window will pop up. On the left of the window, click on "VBA Project (LiquidAddin.xlam) And this will reveal contents of the addin. There are two important files in the addin. The first one you can find it in the part where it says "Forms". This is basically our User Interface which can be edited by drag and drop. Double clicking on that interface will reveal the code for functionality of the components on the form such as text boxes. The second important file is in Modules and this contains most of our useful functions. Below is an overview of these functions:

# Functions in the Module LiquidQuotation
LParts
- This function splits the contents of the parameter "CellRef", and uses VLookup to look up the part numbers and then uses a loop to store these part numbers in a single string.

LPrice
- This function splits the contents of the parameter "CellRef" and uses VLookup to look up the price that matches each part number. It then Uses a loop to add these together and takes into consideration the condition in which we don't charge Labour to the customer.

getstr
- This fuction loops through the drop down combo boxes and the text boxes and gets their contents and then stores them into two strings (One for faults and the other one for repairs) separating each value by comma.

Quotation
- This is defined by the key word Sub because it's a "Macro" Function. This basically initializes our Quotation Form whenever we call it. This is where we added stuff onto the combo boxes. It also locates the first empty cell on the second column of the table.

# Functions in the LiquidForm module

These are the functions that determine what happens when we press the submit button, Cancel Button, and any other button.

Here are useful tutorials for those who are not familiar with VBA but would like to help develop the Form Further:

https://trumpexcel.com/excel-add-in/ : How to create a user form.

https://www.thespreadsheetguru.com/blog/build-modern-vba-userforms: Making it look modern.






