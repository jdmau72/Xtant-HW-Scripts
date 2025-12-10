# Xtant Medical HW Scripts
A collection of ExcelScripts used to streamline many functions of the Hardware team. Scripts available as both .osts and .ts files - the former used in Excel, the latter for ease of reading and editing. .ts files can quickly be "compiled" to .osts files and vice versa with the provided bash scripts. 

## HOW TO INSTALL AND RUN
(!) If you have not used any ExcelScripts before do the following:
1. On any open Excel file, navigate to the "Automate" pane at the top
2. Click "New Script". A panel will open on the right side of the screen. 
3. Click the "${\textsf{\color{green}+}{\color{white}New Script"}}$ . This will create the necessary folder to add and use these scripts. 
4. Click the back arrow to go back to ${\textsf{\color{green}"<- All Scripts"}}$
5. At the bottom of the tab, click ${\textsf{\color{green}"View more scripts ->"}}$

(!) Otherwise, do the following:
1. Navigate to "Automate" / "All Scripts"
2. At the bottom of the Code Editor, click ${\textsf{\color{green}"View more scripts ->"}}$
3. Download the scripts you need from this repository
4. Move the .osts files into the Office Scripts folder (OneDrive/My Files/Documents/Office Scripts)

Now you have the scripts available to run in any Excel file you wish. Hope they help!

## Available Scripts
### AutoFormat_JDM.osts
Sets up the necessary columns and generates conditional formatting for ${\textsf{\color{red}FAIL}}$, ${\textsf{\color{grey}MIA}}$, and ${\textsf{\color{green}EXTRA}}$ statuses, as well as formatting for ${\textsf{\color{yellow}Lot Adjustments}}$. 

Creates borders between different item numbers to group together lots of any given item.

### GenerateCSV_ADJ.osts
An ExcelScript that parses a HW return sheet to create a list of adjustments. It then generates a csv that can be uploaded to NetSuite to process quickly process multiple adjustments for a tray. 

### OrganizeLooseReturnSheet.osts
Script takes all items of a certain category (${\textsf{\color{red}FAIL}}$, ${\textsf{\color{grey}MIA}}$, and ${\textsf{\color{green}EXTRA}}$, etc) and copies them to a separate sheet. Helps declutter and organize the return sheet.