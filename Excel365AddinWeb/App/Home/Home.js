/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#send-data-from-selection').click(sendDataFromSelection)
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value === '')
                    {
                        app.showNotification('There was no selected text');
                    }
                    else { app.showNotification('The selected text is:', '"' + result.value + '"'); }
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function sendDataFromSelection()
    {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
           function (result) {
               if (result.status === Office.AsyncResultStatus.Succeeded) {
                   if (result.value === '') {
                       app.showNotification('There was no selected text');
                   }
                   else { app.showNotification('The selected text is:', '"' + result.value + '"'); }
               } else {
                   app.showNotification('Error:', result.error.message);
               }
           });
    }
})();