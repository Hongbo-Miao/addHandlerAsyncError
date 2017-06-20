# Instructions to Reproduce Error:
Clone the repo to your local directory

Use `npm install` to install all necessary packages

Copy the MLServerExcelAddinManifest.xml file to a shared directory and use that to sideload the add-in in Excel

From the console, start the add-in with `npm start` and sideload it in Excel

Click 'Bind to A1'

Then click 'Add handler to A1'

No error will appear as if nothing happened.

If you click 'Add handler to document overall'

You will see that 'addHandlerAsync' has invalid parameters for this one.

All Excel logic is in the ./services/excel.service.ts directory