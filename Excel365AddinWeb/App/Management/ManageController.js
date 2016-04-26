(function () {
    Office.initialize = function (reason) {

    };
});

(function () {
    
    var managerService = function ($q) {
        //Reads data from current document selection and displays a notification
      
        this.sendDataFromSelection = function () {
            var deferred = $q.defer();
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
               function (result) {
                   if (result.status === Office.AsyncResultStatus.Succeeded) {
                       deferred.resolve(result.value.split('\n'));
                   }
                    else { deferred.reject([]); }
               });
            return deferred.promise;
        }

        this.updateRange = function ()
        {
            var deferred = $q.defer();

            // Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
            Excel.run(function (ctx) {

                // Create a proxy object for the sheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                // Values to be updated
                var values = [
                             ["Type", "Estimate"],
                             ["Transportation", 1670],
                             ["Food", 800],
                             ["Fuel", 1111]
                ];
                // Create a proxy object for the range
                var range = sheet.getRange("A1:B4");

                // Assign array value to the proxy object's values property.
                range.values = values;

                // Queue a command to load the text property for the proxy range object.    
                range.load('text');

                // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context 
                return ctx.sync().then(function () {
                    console.log("Done");
                    deferred.resolve();

                });
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

            return deferred.promise;

        }
    }

    var manageCtrl = function ($scope, managerService) {
        var vm = this;
        vm.Hello = 'Person';
        vm.showMessage = false;
        vm.message = {header:'', body:''};
        vm.test = function() { vm.showMessage = true; vm.message.body ="HELLOOOOOO" };
        vm.results = [];
        //$scope.$watchCollection('vm.results', function (newValue, oldValue) {
        //    vm.results = angular.copy(newValue);
        //});

        function sendDataFromSelection() {
            managerService.sendDataFromSelection().then(
               function (result) {
                   vm.showMessage = true;

                   if (result.length === 0) {
                       vm.message.body = 'There was no selected text';
                   }
                   else {
                       vm.message = { header: 'The selected text was sent to the server:', body: result.join(' ') };
                      vm.results = result;
                      //vm.results = angular.copy(result);

                       //var element = angular.element($('#resultsRow'));
                       //element.scope().$apply();
                       //$scope.$apply();
                   }
               },
               function (error) {
                   vm.message = { header: 'Error:', body: error.message };
               }
               );
        }


        vm.updateRange = function ()
        {
            managerService.updateRange();
        }

        vm.sendDataFromSelection = sendDataFromSelection;
    }

    angular.module('appMain').service('managerService', managerService);
    angular.module('appMain').controller('manageCtrl', manageCtrl);

})();