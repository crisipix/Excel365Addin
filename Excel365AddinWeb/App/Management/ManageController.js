(function () {
    Office.initialize = function (reason) {

    };
});

(function () {
    
    var managerService = function ($q) {
        //Reads data from current document selection and displays a notification
        this.getDataFromSelection = function () {
            var deferred = $q.defer();
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
               function (result) {
                   if (result.status === Office.AsyncResultStatus.Succeeded) {
                       deferred.resolve(result.value.split('\n'));
                   }
                   else {
                       deferred.reject([]);
                   }
               });
            return deferred.promise;
        }

        this.sendDataFromSelection = function () {
            var deferred = $q.defer();
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
               function (result) {
                   if (result.status === Office.AsyncResultStatus.Succeeded) {
                       deferred.resolve(result.value.split('\n'));
                   }
                    else {
                       deferred.reject([]);
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
        $scope.$watchCollection('results', function (newValue, oldValue) { vm.results = newValue; });

        function getDataFromSelection() {
            managerService.getDataFromSelection().then(
                function (result) {
                    vm.showMessage = true;

                    if (result.length === 0) {
                        vm.message.body = 'There was no selected text';
                    }
                    else {
                        vm.message = { header: 'The selected text was sent to the server:', body: result.join(' ')  };
                        vm.results = result;
                    }
                },
                function (error) {
                    vm.message = { header: 'Error:', body: error.message };
                }
                );
           

        }

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

                      // $scope.$apply();
                   }
               },
               function (error) {
                   vm.message = { header: 'Error:', body: error.message };
               }
               );
        }

        vm.getDataFromSelection = getDataFromSelection;
        vm.sendDataFromSelection = sendDataFromSelection;
        

    }

    angular.module('appMain').service('managerService', managerService);
    angular.module('appMain').controller('manageCtrl', manageCtrl);

})();