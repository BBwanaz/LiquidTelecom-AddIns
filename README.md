# LiquidTelecom-AddIns


>The features include:
 
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


# Troubleshooting
If you have your price in the price column in Master, and the app still tells you that No price found, try changing the formating of the first column into text by doing this.

Use the Text to columns option to convert exponential notation to text (Click on column header, click Text to data, then Delimited, Next, untick all delimiter boxes, Select Text in Column data format and then Finish) to have the values converted to text and immediately displayed as 15 digit number.

https://trumpexcel.com/excel-add-in/ : How to create a user form.

https://www.thespreadsheetguru.com/blog/build-modern-vba-userforms: Making it look modern.






