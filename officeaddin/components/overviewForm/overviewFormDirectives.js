(function () {

  "use strict";

  var app = angular.module("insightsOfficeApp");

  app.directive("overviewForm", function() {
        return {
            templateUrl: "components/overviewForm/overviewFormView.html",
            controller: "FormController",
            controllerAs: "formCtrl"
        };
    });   
})();
