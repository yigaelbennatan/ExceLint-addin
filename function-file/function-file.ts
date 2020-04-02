/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    // Add the event handler
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  };

  /**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
  async function onChange(event) {
    return Excel.run(function(context) {
      return context.sync().then(function() {
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
      });
    });
  }
  
  

  // Add any ui-less function here
})();
