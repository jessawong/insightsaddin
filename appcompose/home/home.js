(function(){
  'use strict';

  var app = angular.module('readHome', []);
  //var selectTech = [{'type':'Intelligent Cloud', 'url':'../../images/cloud.png', 'alias':'usdxrisvintelligentcloudteam@service.microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true}];
  var info = {'pbe': '', 'website': '', 'date':'', 'time':'', 'reason':'', 'meeting':'Skype', 'location':'', 'engagement':'', 'crm': '', 'stage': ''};
  var cloudInfo = {'status': '', 'provider':'', 'consumption':'', 'workloads':'', 'vteam':''};
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      
      info.pbe = Office.context.mailbox.userProfile.displayName;
      //$("#pbeName").text(info.pbe);
      //Office.context.mailbox.item.subject.setAsync("Intelligent Cloud TE Request");
      
     // Office.context.mailbox.item.to.setAsync([selectTech[0].alias]);
     /* $(".valign").hover(function() {
        $(this).stop().animate({"opacity":"1.0"}, "fast");
      },
      function() {
        $(this).stop().animate({"opacity":"0.0"}, "fast");
      });*/
    });
  };

    app.controller('FormController', function($scope) {
        info["Technology"] = "Intelligent Cloud";
        $scope.technology = [{'type':'Intelligent Cloud', 'url':'../../images/cloud.png', 'alias':'usdxrisvintelligentcloudteam@service.microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true}];
        $scope.engagement = ['Briefing', 'Envisioning', 'ADS', 'Hackfest/PoC', 'Other'];;
        $scope.cloudInfor = {'status': '', 'provider':'', 'consumption':'', 'workloads':'', 'vteam':''};
        $scope.cloudStatus = ['New', 'Experimenting', 'Hybrid', 'Running'];
        $scope.cloudProv = ['None','Azure', 'AWS', 'Google', 'Other'];
        $scope.consumption = ['<25k', '25k-99k', '100k-499k', '500k+'];
        $scope.primeWorkloads = ['Modern Datacenter(IT Pro)', 'Data Platform & Analytics', 'Modern Apps (Cloud Dev)'];
        $scope.secondWorkloads = {'Identity & Access Mgt': false, 'Compute':false, 'Networking':false, 'Storage & DR':false, 'Big Data & SQL':false,'IoT':false,'Advanced Analytics':false, 'PaaS Services':false,'OSS Platforms':false, 'Mobility':false, 'Media Solutions':false};
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
        $scope.showMain = true;
        $scope.goBack = function() {
          $scope.showCont = false;
          $scope.showMain = true;
          $scope.showArrow = true;
          $scope.showBack = false;
          $scope.showSubmit = false;
        };

        $scope.setWork = function(option) {
          option.value = !option.value;
        };

        $scope.cont = function() {
          $scope.showCont = true;
          $scope.showBack = true;
          $scope.showSubmit = true;
          $scope.showMain = false;
          $scope.showArrow = false;
        };

        $scope.addRequest = function() {
          console.log($scope.information);
          var cloudDetails = '';
          if ($scope.intelCloud) {
            cloudDetails = "<br/><h4>Industry/Vertical: </h4>" + $scope.cloudInfor.industry + "<br/><h4>Cloud Status: </h4>" + this.cloudInfor.status
                            + "<br/><h4>Cloud Provider: </h4>" + $scope.cloudInfor.provider + "<br/><h4>Annual Consumption: </h4>" + this.cloudInfor.consumption;
            var workloadString = "[";
            for (var key in $scope.secondWorkloads) {
              if ($scope.secondWorkloads[key]) {
                workloadString += key + ",";
              }
            }
            workloadString = workloadString.substr(0, workloadString.length - 1) + "]";
          }
          Office.context.mailbox.item.subject.setAsync("[" + $scope.cloudInfor.vteam + "] Intelligent Cloud TE Request");
          Office.context.mailbox.item.body.setAsync("<h4>Product's Website: </h4>" + $scope.information.website + "<br/><h4>CRM Link: </h4>" + $scope.information.crm
                                                      + "<br/><h4>Stage:</h4>" + $scope.information.stage + "<br/><h4>Engagement Requested: </h4>" 
                                                      + $scope.information.engagement + "<br/><h4>Requested Date for Engagement:</h4>" + $scope.information.date 
                                                      + "<br/><h4>Reason:</h4>" + $scope.information.reason + "<br/><h4>Duration of meeting:</h4>" 
                                                      + $scope.information.time + "<br/><h4>Location:</h4>" + $scope.information.location + "<br/><h4>Meeting:</h4>" 
                                                      + $scope.information.meeting + cloudDetails + "<br/><h4>Secondary Workloads:</h4>" + workloadString, {coercionType: "html"});
          //REDUNDANCY EVERYWHERE. NEED TO FIX.                                            
          var payload = { 
                                  "Pbe": info.pbe,
                                  "Website": information.website, 
                                  "Crm": information.crm, 
                                  "Stage": information.stage, 
                                  "EngagementType": information.engagement,
                                  "Date": information.date,
                                  "Reason": information.reason,
                                  "Location": information.location,
                                  "Meeting": information.meeting,
                                  "Industry": cloudInfor.industry,
                                  "cloudStatus": cloudInfor.status,
                                  "cloudProvider": cloudInfor.provider,
                                  "consumption": cloudInfor.consumption,
                                  "primaryWorkLoad": cloudInfor.vteam,
                                  "secWorkLoads": workloadString
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
        $scope.canSkype = function() {
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
