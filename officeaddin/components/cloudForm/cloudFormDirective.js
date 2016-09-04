(function () {

  "use strict";

  var app = angular.module("insightsOfficeApp");

  app.directive("cloudForm", function() {
        return {
            templateUrl: "components/cloudForm/cloudFormView.html"
        };
    });   
})();
