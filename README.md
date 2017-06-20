# Instructions to Reproduce Error:
Clone the repo to your local directory

Use `npm install` to install all necessary packages

Copy the MLServerExcelAddinManifest.xml file to a shared directory and use that to sideload the add-in in Excel

From the console, start the add-in with `npm start` and sideload it in Excel

Click 'Bind to A1'

Then click 'Add handler to A1'

It will supposedly successfully create the handler.  However, if data is changed in cell A1, the event never triggers the handler function.

If you click 'trigger communication from service', you will see the communication of 'data change' that is supposed to be sent when data is changed in the binding.