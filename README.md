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
