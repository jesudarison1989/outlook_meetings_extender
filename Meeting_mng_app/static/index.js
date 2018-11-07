var app = angular.module('exhangeApp', []);
app.controller('myCtrl',['$scope', '$http', '$location', function ($scope, $http, $location) {
        $scope.isSucess = false;
        $scope.getCredential = false;
        $scope.isHome = true;
        $scope.userPermission = {};

        //Approve or reject call
        $scope.requestApprovaed = function (decision) {
                if(decision == "ChangeRoom"){
                        decision = $scope.selctedRoom;
                }
                $http({
                        method: 'POST',
                        url: 'http://localhost:5000/requestResp',
                        headers: {
                                'Content-Type': 'application/json'
                        },
                        data: JSON.stringify(decision)
                }).then(function (resp) {
                        if(resp.status == '200'){
                                $scope.getCredential = false;
                                $scope.isSucess = true;
                        }
                },function (e) {
                        console.log(e);
                });
        };
        
        //select room
        $scope.changeRoom = function(roomDetails){
                $scope.selctedRoom = roomDetails;
        };

        //submit Selected room
        $scope.submitChangeSuggesetion = function(action){
                $scope.isHome = false;
                $scope.getCredential = true
                $scope.userAction = action;
        };

        $scope.submitCred = function(){
                data ={
                        status: $scope.userAction,
                        userId: $scope.userPermission.userId,
                        password: $scope.userPermission.userPwd
                };
                $scope.requestApprovaed(data);
        };

        //Sugested rooms
        $scope.meetingSuggestion = function () {
                $scope.metingId = "dsdsdsdsdsdsd"
                $http({
                        method: 'POST',
                        url: 'http://localhost:5000/meetingSuggestion',
                        headers: {
                                'Content-Type': 'application/json'
                        },
                        data: JSON.stringify($scope.metingId)
                }).then(function (resp) {
                        $scope.meetingSug = resp.data;
                },function (e) {
                        console.log(e);
                });
        };

        $scope.meetingSuggestion();

}]);