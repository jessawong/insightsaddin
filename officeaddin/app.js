angular.module("insightsOfficeApp").config(["$stateProvider", "$urlRouterProvider", function($stateProvider, $urlRouterProvider){

    $urlRouterProvider.when("", "/");

    $stateProvider
        .state("main", {
            url: "/",
            templateUrl: "components/main.html",
            controller: "mainController",
            controllerAs: "ctrl"
        });
}]);