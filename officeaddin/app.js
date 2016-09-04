var app = angular.module("insightsOfficeApp");

app.config(["$stateProvider", "$urlRouterProvider", function ($stateProvider, $urlRouterProvider) {

    $urlRouterProvider.otherwise("/404");
    $urlRouterProvider.when("", "/");

    $stateProvider
        .state("main", {
            url: "/",
            templateUrl: "app.html",
            controller: "FormController",
            controllerAs: "formCtrl"
        })
        .state("404", {
            url: "/404",
            templateUrl: "app/shared/404.html"
        });
}]);