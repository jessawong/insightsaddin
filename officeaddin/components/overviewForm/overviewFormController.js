(function () {

  "use strict";

  var info =
    {
      "te": "",
      "pbe": "",
      "isv": "",
      "si": "",
      "isvWebsite": "",
      "crmLink": "",
      "stage": "",
      "city": "",
      "state": "",
      "date": "",
      "time": "",
      "meeting": "Skype",
      "engagement": ""
    };

  var cloudInfo =
    {
      "industry": "",
      "industryWorkload": "",
      "status": "",
      "provider": "",
      "consumption": "",
      "workloads": ""
    };

  Office.initialize = function (args) {
    jQuery(document).ready(function () {
      info.te = Office.context.mailbox.userProfile.displayName;
    });
  };

  var app = angular.module("insightsOfficeApp");

  app.controller("FormController", formController);

  function formController($scope) {
    info["Technology"] = "Intelligent Cloud";
    $scope.technology = [{ "type": "Intelligent Cloud", "url": "../../images/cloud.png", "alias": "usdxrisvintelligentcloudteam@service.microsoft.com", "fadeUrl": "../../images/cloudOverlay.png", "intel": true }];
    $scope.engagement = ["Briefing", "Envisioning", "ADS", "Hackfest/PoC", "Other"];;
    $scope.cloudInfor = { "status": "", "provider": "", "consumption": "", "workloads": "" };
    $scope.cloudStatus = ["New", "Experimenting", "Hybrid", "Running"];
    $scope.cloudProv = ["None", "Azure", "AWS", "Google", "Other"];
    $scope.consumption = ["<25k", "25k-99k", "100k-499k", "500k+"];
    $scope.workloads = { "Advanced Analytics": false, "Big Data & SQL": false, "Compute": false, "Identity & Access Mgt": false, "IoT": false, "Media Solutions": false, "Mobility": false, "Networking": false, "OSS Platforms": false, "PaaS Services": false, "Storage & DR": false };
    $scope.information = info;
    $scope.timeOptions = ["30 min", "60 min", "90 min", "120 min", "2+ hours"];
    $scope.skype = false;
    $scope.intelCloud = true;
    $scope.showArrow = true;
    $scope.showCont = false;
    $scope.ads = false;
    $scope.showSubmit = false;
    $scope.showBack = false;
    $scope.stages = ["0%", "10%", "20%", "40%", "60%", "80%", "95%", "100%"];
    $scope.pbes = ["Alexandra Detweiler", "Beverly Ann Smith", "Bill Lyle", "David Cazel", "Frances Calandra", "Harsha Vishwanathan", "Hong Choing", "Jon Box", "Kevin Boyle", "Micheal Liwanag", "Paul Debaun", "Tina Prause", "Tony Piltzecker", "Wes Yanage", "Will Tschumy", "Sam Chenaur"];
    $scope.industries = ["CDN", "Cross Industry", "Dev Tools", "Education", "Finance Services", "Health Care", "Insurance", "Manufacturing", "Mining", "Oil and Gas", "Public Sector", "Retail"];
    $scope.showMain = true;
    $scope.goBack = function () {
      $scope.showCont = false;
      $scope.showMain = true;
      $scope.showArrow = true;
      $scope.showBack = false;
      $scope.showSubmit = false;
    };

    $scope.setWork = function (option) {
      option.value = !option.value;
    };

    $scope.cont = function () {
      $scope.showCont = true;
      $scope.showBack = true;
      $scope.showSubmit = true;
      $scope.showMain = false;
      $scope.showArrow = false;
    };

    $scope.addRequest = function () {
      var cloudDetails = "";
      if ($scope.intelCloud) {
        cloudDetails =
          "<br/><h4>Industry/Vertical: </h4>" + $scope.cloudInfor.industry +
          "<br/><h4>Cloud Status: </h4>" + this.cloudInfor.status +
          "<br/><h4>Cloud Provider: </h4>" + $scope.cloudInfor.provider +
          "<br/><h4>Annual Consumption: </h4>" + this.cloudInfor.consumption;
        var workloadString = "[";
        for (var key in $scope.workloads) {
          if ($scope.workloads[key]) {
            workloadString += key + ",";
          }
        }
        workloadString = workloadString.substr(0, workloadString.length - 1) + "]";
      }

      Office.context.mailbox.item.subject.setAsync("Intelligent Cloud Partner Meeting Verification");
      Office.context.mailbox.item.body.setAsync(
        "<h4>TE: </h4>" + info.te +
        "<h4>SI: </h4>" + $scope.information.si +
        "<h4>ISV: </h4>" + $scope.information.isv +
        "<h4>ISV Website: </h4>" + $scope.information.isvWebsite +
        "<br/><h4>CRM Link: </h4>" + $scope.information.crmLink +
        "<br/><h4>Stage:</h4>" + $scope.information.stage +
        "<br/><h4>PBE:</h4>" + $scope.information.pbe +
        "<br/><h4>Engagement Requested: </h4>" + $scope.information.engagement +
        "<br/><h4>Requested Date for Engagement:</h4>" + $scope.information.date +
        "<br/><h4>Duration of meeting:</h4>" + $scope.information.time +
        "<br/><h4>Location:</h4>" + $scope.information.city + ", " + $scope.information.state +
        "<br/><h4>Meeting:</h4>" + $scope.information.meeting + cloudDetails +
        "<br/><h4>Workloads:</h4>" + workloadString, { coercionType: "html" }
      );

      var payload =
        {
          "te": info.te,
          "pbe": $scope.information.pbe,
          "isv": $scope.information.isv,
          "si": $scope.information.si,
          "isvWebsite": $scope.information.isvWebsite,
          "crmLink": $scope.information.crmLink,
          "stage": $scope.information.stage,
          "city": $scope.information.city,
          "state": $scope.information.state,
          "workLoads":
          {
            "Advanced Analytics": $scope.workloads["Advanced Analytics"],
            "Big Data & SQL": $scope.workloads["Big Data & SQL"],
            "Compute": $scope.workloads["Compute"],
            "Identity & Access Mgt": $scope.workloads["Identity & Access Mgt"],
            "IoT": $scope.workloads["IoT"],
            "Media Solutions": $scope.workloads["Media Solutions"],
            "Mobility": $scope.workloads["Mobility"],
            "Networking": $scope.workloads["Networking"],
            "OSS Platforms": $scope.workloads["OSS Platforms"],
            "PaaS Services": $scope.workloads["PaaS Services"],
            "Storage & Disaster Recovery": $scope.workloads["Storage & Disaster Recovery"]
          },
          "engagement":
          {
            "date": $scope.information.date,
            "notes": "notes",
            "meeting": $scope.information.meeting,
            "type": $scope.information.engagement,
          },
          "industry": $scope.cloudInfor.industry,
          "industryWorkload": $scope.cloudInfor.industryWorkload,
          "cloudStatus": $scope.cloudInfor.status,
          "cloudProvider": $scope.cloudInfor.provider,
          "consumption": $scope.cloudInfor.consumption
        };

      var namespace = "insightsaddin-eh";
      var hubname = "insights-eh";
      var key_name = "send";

      var ehClient = new EventHubClient(
        {
          "name": hubname,
          "namespace": namespace,
          "sasKey": "HsoQY4aVbX2cOEhg5hqwfzQDSLLIKyvvIrNt/u2jU+k=",
          "sasKeyName": "send",
          "timeOut": 100,
        });

      var msg = new EventData(payload);

      ehClient.sendMessage(msg, function (messagingResult) {
        if (messagingResult.result == "Success") {
          showStatus("#submission");
        } else {
          showStatus("#error");
          console.log(messagingResult);
        }
        console.log("Complete! This payload is now in the database: " + JSON.stringify(payload));
      });

      $scope.showBack = false;
      $scope.skype = false;
    };
    $scope.canSkype = function () {
      $scope.skype = !$scope.skype;
    };
  };

  function showStatus(domId) {
    $("#form").hide();
    $(domId).show();
  }

  function hideStatus() {
    $("#submission").hide();
  }

  function showForm() {
    $("#form").show();
  }

})();
