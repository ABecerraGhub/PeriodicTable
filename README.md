'////////////////////////////////////  ((((((()))))))  \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
MODEL - ELECTRIC FIELD FOR ELEMENTS & COMPOUNDS

Modeling Physics and Mathematics in Microsoft Excel
© 2025, Ariel R. Becerra B., Alvaro Herrera, Martha Molina|

This package provides a computational model for calculating and visualizing electric fields for elements of the periodic table and their combinations. It is implemented as a VBA macro-enabled workbook for Microsoft Excel (tested on Excel 2013-2019 for Windows and Mac).

Installation

You can install the model using one of the following two methods.

 Option 1: Ready-to-Use Workbook (Recommended)

	1. Download the pre-configured Excel workbook ScienSolarPeriodicTable_E_Field.xlsm.
	2. Open the downloaded file. Enable macros if prompted.	
	3. The model is ready to use.

 Option 2: Manual Integration from Source Files
This option requires integrating the source code into a blank Excel workbook. This option is recommended if option 1 was not satisfactory.

1.  Download the three source code files:
    *   `1_6MODULE_1.txt` (Main Engine)
    *   `1_6MODULE_2.txt` (Main Engine 2)
    *   `1_6MODULE_3.txt` (Periodic Table Elements Model)

2.  Create a New Workbook:
    *   Open Microsoft Excel and create a new, blank workbook.

3.  Open the VBA Editor:
    *   Windows: Press `Alt + F11`
    *   macOS: Press `Fn + Option + F11`

4.  Import the Modules:
    *   In the VBA Editor, right-click on `VBAProject (YourWorkbookName.xlsx)` in the Project Explorer pane.
    *   Select Insert > Module. Repeat this step to insert three separate modules.
    *   For each new module, copy and paste the entire contents of one corresponding text file into it.
        *   Paste `1_6MODULE_1.txt` into `Module1`
        *   Paste `1_6MODULE_2.txt` into `Module2`
        *   Paste `1_6MODULE_3.txt` into `Module3`

5.  Close the VBA Editor and return to your Excel workbook.

 Getting Started

1.  Create the Interface Button:
    *   Ensure the Developer tab is visible in Excel's ribbon. (If not, enable it in `File > Options > Customize Ribbon`).
    *   Go to the Developer tab, click Insert, and choose the Button (Form Control).
    *   Click anywhere on a worksheet to place the button. The "Assign Macro" dialog will appear.
    *   Select the macro named `NewSheet` and click OK.

2.  Start a New Project:
    *   Click the button you just created. This will generate a new sheet with the ScienSolar interface.

3.  Load a Sample Project:
    *   On the new interface sheet, select the project number 3 (Periodic Table Elements) from the provided list.
    *   Click the +Vector button to load it.

4.  Save Your Workbook:
    *   Save the file as a Macro-Enabled Workbook (`*.xlsm`) to preserve all functionality.

 Support & Resources

*   For ScienSolar documentation, tutorials, and updates, please visit the official website:
    [www.sciensolar.com](http://www.sciensolar.com)
