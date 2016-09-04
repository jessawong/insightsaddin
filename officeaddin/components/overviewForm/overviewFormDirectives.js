(function () {

  "use strict";

  var app = angular.module("insightsOfficeApp", []);

  app.directive("overviewForm", function() {
        return {
            templateUrl: "../../overviewFormView.html"
        };
    });   

})();
