# Instructions to Reproduce Error:
Clone the repo to your local directory

Use `npm install` to install all necessary packages

Copy the MLServerExcelAddinManifest.xml file to a shared directory and use that to sideload the add-in in Excel

From the console, start the add-in with `npm start` and sideload it in Excel

Click 'Bind to A1'

Then click 'Add handler to A1'

If you do some sort of change in cell A1, it will trigger Hello World to be printed in it.  However, this handler method changeEvent(eventArgs: any) which is located in src/app/app.component.ts cannot call any methods within the app component.  It can only call native Office JS functions.  Is this expected behavior?  What should I do if I want to get around this?