(function () {

  "use strict";

  var app = angular.module('readHome', []);

  var info = { 'te': '', 'pbe': '', 'si': '', 'isv': '', 'website': '', 'date': '', 'time': '', 'reason': '', 'meeting': 'Skype', 'city': '', 'state': '', 'engagement': '', 'crm': '', 'stage': ''};
  var cloudInfo = { 'status': '', 'provider': '', 'consumption': '', 'workloads': '' };

  Office.initialize = function (reason) {
    jQuery(document).ready(function () {
      info.te = Office.context.mailbox.userProfile.displayName;
    });
  };

  app.controller('FormController', function ($scope) {
    info["Technology"] = "Intelligent Cloud";
    $scope.technology = [{ 'type': 'Intelligent Cloud', 'url': '../../images/cloud.png', 'alias': 'usdxrisvintelligentcloudteam@service.microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true }];
    $scope.engagement = ['Briefing', 'Envisioning', 'ADS', 'Hackfest/PoC', 'Other'];;
    $scope.cloudInfor = { 'status': '', 'provider': '', 'consumption': '', 'workloads': '' };
    $scope.cloudStatus = ['New', 'Experimenting', 'Hybrid', 'Running'];
    $scope.cloudProv = ['None', 'Azure', 'AWS', 'Google', 'Other'];
    $scope.consumption = ['<25k', '25k-99k', '100k-499k', '500k+'];
    $scope.workloads = { 'Identity & Access Mgt': false, 'Compute': false, 'Networking': false, 'Storage & DR': false, 'Big Data & SQL': false, 'IoT': false, 'Advanced Analytics': false, 'PaaS Services': false, 'OSS Platforms': false, 'Mobility': false, 'Media Solutions': false };
    $scope.information = info;
    $scope.timeOptions = ['30 min', '60 min', '90 min', '120 min', '2+ hours'];
    $scope.skype = false;
    $scope.intelCloud = true;
    $scope.showArrow = true;
    $scope.showCont = false;
    $scope.ads = false;
    $scope.showSubmit = false;
    $scope.showBack = false;
    $scope.stages = ['0%', '10%', '20%', '40%', '60%', '80%', '95%', '100%'];
    $scope.pbes = ['30 min', '60 min', '90 min', '120 min', '2+ hours'];
    //$scope.pbes = ["Alexandra Detweiler", "Beverly Ann Smith", "Bill Lyle", "David Cazel", "Frances Calandra", "Harsha Vishwanathan", "Hong Choing", "Jon Box", "Kevin Boyle", "Micheal Liwanag", "Paul Debaun", "Tina Prause", "Tony Piltzecker", "Wes Yanage", "Will Tschumy", "Sam Chenaur"];
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
      var cloudDetails = '';
      if ($scope.intelCloud) {
        cloudDetails = "<br/><h4>Industry/Vertical: </h4>" + $scope.cloudInfor.industry + 
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
      
      Office.context.mailbox.item.subject.setAsync("Intelligent Cloud TE Request");
      Office.context.mailbox.item.body.setAsync(
        "<h4>PBE: </h4>" + $scope.information.pbe +
        "<h4>SI: </h4>" + $scope.information.si +
        "<h4>ISV: </h4>" + $scope.information.isv +
        "<h4>Product's Website: </h4>" + $scope.information.website +
        "<br/><h4>CRM Link: </h4>" + $scope.information.crm +
        "<br/><h4>Stage:</h4>" + $scope.information.stage +
        "<br/><h4>Engagement Requested: </h4>" + $scope.information.engagement +
        "<br/><h4>Requested Date for Engagement:</h4>" + $scope.information.date +
        "<br/><h4>Reason:</h4>" + $scope.information.reason +
        "<br/><h4>Duration of meeting:</h4>" + $scope.information.time +
        "<br/><h4>Location:</h4>" + $scope.information.location +
        "<br/><h4>Meeting:</h4>" + $scope.information.meeting + cloudDetails +
        "<br/><h4>Workloads:</h4>" + workloadString, { coercionType: "html" }
      );

      // TODO: Add drop down for industry verticals
      // manufacturing, public sector, mining, dev tools, cross industry, health care, insurance, oil and gas, education, CDN, retail, finance services
      var payload = {
        "TE": info.te,
        "PBE": $scope.information.pbe,
        "SI": $scope.information.si,
        "ISV": $scope.information.isv,
        "Website": $scope.information.website,
        "Crm": $scope.information.crm,
        "Stage": $scope.information.stage,
        "EngagementType": $scope.information.engagement,
        "Date": $scope.information.date,
        "Reason": $scope.information.reason,
        "City": $scope.information.city,
        "State": $scope.information.state,
        "Meeting": $scope.information.meeting,
        "Industry": $scope.cloudInfor.industry,
        "cloudStatus": $scope.cloudInfor.status,
        "cloudProvider": $scope.cloudInfor.provider,
        "consumption": $scope.cloudInfor.consumption,
        "workLoads": workloadString
      };

      var namespace = "insightsaddin-eh";
      var hubname = "insights-eh";
      var key_name = "send";

      var ehClient = new EventHubClient(
        {
          'name': hubname,
          'namespace': namespace,
          'sasKey': "HsoQY4aVbX2cOEhg5hqwfzQDSLLIKyvvIrNt/u2jU+k=",
          'sasKeyName': "send",
          'timeOut': 100,
        });

      var msg = new EventData(payload);

      ehClient.sendMessage(msg, function (messagingResult) {
        if (messagingResult.result == "Success") {
          showStatus("#submission");
        } else {
          showStatus("#error");
        }
        console.log(JSON.stringify(payload));
      });

      $scope.showBack = false;
      $scope.workloads = workLoads;
      $scope.skype = false;
    };
    $scope.canSkype = function () {
      $scope.skype = !$scope.skype;
    };

  });

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
