# Knight's Tour Solution in Excel VBA

This is my demo of solving the Knights Tour problem using Warnsdorf Rule.

You can view a video the solution here: https://player.vimeo.com/video/233579477

The code files have been separated from the main workbook, so they can be versioned controlled in Github.

To view a demo of the solution open up the file demo.xlsb and press the button 'Start'.


## Preliminary notes

This project uses a custom Chess Font to display a Knight Chess piece. The demo workbook when opened, will automatically install the font on your machine and then
uninstall the font when closed down.

If you have any reservations about this, disable macros when you open the workbook or you can use the code files below and create your own version.

Without the font installed the program should still run, however it won't look like any of the screenshots/video that you might have seen of the application in action.

Always refer to the version controlled code files to get the latest code changes.


## Setup with the latest version of the code 

To setup with the latest version of the code, do the following:

- Open the file UI.xlsb
- Import the class files: clsFontInstall, clsKnightTour
- Import the module file: mdMain
- Ensure the font file Alpha.ttf is in the same location as the UI.xlsb file.
- Add the code in the ThisWorkbook.cls file to the ThisWorkbook code module.
- Assign the macro called 'StartKnightsTour' to the Start button on the main worksheet.
- Run the 'StartKnightsTour' macro to see the code in action.


## Issues

You may have to close and reopen the macro workbook for the Chess Font to register with Excel.

## Acknowledgements

The Chess Font was created by Eric Bentzen and has its own copyright notice which is included with this project.

## Further info

- https://en.wikipedia.org/wiki/Knight%27s_tour
- https://datapluscode.com/general/solving-the-knights-tour-using-excel-vba/
