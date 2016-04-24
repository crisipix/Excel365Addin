(function () {

    var manageCtrl = function ($scope) {
        var vm = this;
        vm.Hello = 'Person';
        vm.showMessage = false;
        vm.message = {header:'', body:''};
        vm.test = function() { vm.showMessage = true; vm.message.body ="HELLOOOOOO" };
        vm.results = [];

         //Reads data from current document selection and displays a notification
        function getDataFromSelection() {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        vm.showMessage = true;
                        if (result.value === '') {
                            vm.message.body = 'There was no selected text';
                        }
                        else {
                            vm.message = { header: 'The selected text is:', body: result.value };
                            vm.results = result.value.split('\n');
                        }
                    } else {
                        vm.message = { header: 'Error:', body: result.error.message };
                    }
                    $scope.$apply();
                }
            );
        }

        function sendDataFromSelection() {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
               function (result) {
                   if (result.status === Office.AsyncResultStatus.Succeeded) {
                       vm.showMessage = true;

                       if (result.value === '') {
                           vm.message.body = 'There was no selected text';
                       }
                       else {
                           vm.message = { header: 'The selected text was sent to the server:', body: result.value };
                           vm.results = result.value.split('\n');


                       }
                   } else {
                       vm.message ={header : 'Error:', body : result.error.message};
                   }
                   $scope.$apply();
               });
        }

        vm.getDataFromSelection = getDataFromSelection;
        vm.sendDataFromSelection = sendDataFromSelection;
        

    }

    angular.module('appMain').controller('manageCtrl', manageCtrl);

})();