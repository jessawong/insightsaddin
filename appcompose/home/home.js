(function(){
  'use strict';

  var app = angular.module('readHome', []);
  var selectTech = [{'type':'Intelligent Cloud', 'url':'../../images/cloud.png', 'alias':'usdxrisvintelligentcloudteam@service.microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true}];
  var engageType = ['Briefing', 'Envisioning', 'ADS', 'Hackfest/PoC', 'Other'];
  var info = {'pbe':'', 'website': '', 'date':'', 'time':'', 'reason':'', 'meeting':'Skype', 'location':'', 'engagement':'', 'crm': '', 'stage': ''};
  var cloudInfo = {'status': '', 'provider':'', 'consumption':'', 'workloads':''};
  var crmStage = ['0%', '10%', '20%', '40%', '60%', '80%', '95%', '100%'];
  var time = ['30 min', '60 min', '90 min', '120 min', '2+ hours'];
  var status = ['New', 'Experimenting', 'Hybrid', 'Running'];
  var provider = ['None','Azure', 'AWS', 'Google', 'Other'];
  var consumptionLevel = ['<25k', '25k-99k', '100k-499k', '500k+'];
  var workLoads = {'Compute': false, 'Web & Mobile': false, 'Data & Storage': false,
                    'Analytics': false, 'Internet of Things': false,
                    'Networking': false, 'Media & CDN': false, 'Hybrid Integration': false,
                    'Identity & Access Management': false,
                    'Dev Services': false, 'Management & Security': false};


    // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      /*(Check if the browser supports the date input type
      if (!Modernizr.inputtypes.date){
        // Add the jQuery UI DatePicker to all
        // input tags that have their type attributes
        // set to 'date'
        console.log("here");
        $('input[type=date]').datepicker({
            // specify the same format as the spec
            dateFormat: 'mm-dd-yy'
        });
      }*/
      info.pbe = Office.context.mailbox.userProfile.emailAddress;
      Office.context.mailbox.item.subject.setAsync("TE Request");
      $(".valign").hover(function() {
        $(this).stop().animate({"opacity":"1.0"}, "fast");
      },
      function() {
        $(this).stop().animate({"opacity":"0.0"}, "fast");
      });
    });
  };

    app.controller('FormController', function($scope, $http) {
        this.technology = selectTech;
        this.engagement = engageType;
        this.cloudInfor = cloudInfo;
        this.showTech = true;
        this.cloudStatus = status;
        this.cloudProv = provider;
        this.consumption = consumptionLevel;
        this.workloads = workLoads;
        this.information = info;
        this.timeOptions = time;
        this.skype = false;
        this.intelCloud = false;
        this.showArrow = false;
        this.showCont = false;
        this.ads = false;
        this.showSubmit = false;
        this.showBack = false;
        this.stage = crmStage;
        this.showMain = false;
        this.goBack = function() {
          if (this.showCont) {
            this.showCont = false;
            this.showMain = true;
            this.showArrow = true;
          } else {
            this.skype = false;
            this.showTech = true;
            this.intelCloud = false;
            this.showCont = false;
            this.showMain = false;
            this.showArray = false;
            this.showBack = false;
            this.showArrow = false;
            hideStatus();
            this.information["Technology"] = "";
            reset();
          }
          this.showSubmit = false;
        };
        this.setTech = function(option) {
          console.log(this.technology[option]);
          this.information["Technology"] = this.technology[option].type;
          if (this.technology[option].type === "Intelligent Cloud") {
            this.intelCloud = true;
            this.showArrow = true;
            this.showSubmit = false;
          } else {
            this.showSubmit = true;
          }
          Office.context.mailbox.item.to.setAsync([this.technology[option].alias]);
          this.showTech = false;
          this.showMain = true;
          this.showBack = true;
        };

        this.setWork = function(option) {
          option.value = !option.value;
        };

        this.cont = function() {
          this.showCont = true;
          this.showSubmit = true;
          this.showMain = false;
          this.showArrow = false;
        };

        this.addRequest = function() {
          console.log(this.information);
          var cloudDetails = '';
          if (this.intelCloud) {
            cloudDetails = "<br/><h4>Industry/Vertical: </h4>" + this.cloudInfor.industry + "<br/><h4>Cloud Status: </h4>" + this.cloudInfor.status
                            + "<br/><h4>Cloud Provider: </h4>" + this.cloudInfor.provider + "<br/><h4>Annual Consumption: </h4>" + this.cloudInfor.consumption
                            + "<br/><h4>Workloads: </h4>";
            for (var key in this.workloads) {
              if (this.workloads[key]) {
                cloudDetails += key + ",";
              }
            }
          }
          Office.context.mailbox.item.body.setAsync("<h4>Product's Website: </h4>" + this.information.website + "<br/><h4>CRM Link: </h4>" + this.information.crm
                                                      + "<br/><h4>Stage:</h4>" + this.information.stage + "<br/><h4>Engagement Requested: </h4>" 
                                                      + this.information.engagement + "<br/><h4>Requested Date for Engagement:</h4>" + this.information.date 
                                                      + "<br/><h4>Reason:</h4>" + this.information.reason + "<br/><h4>Duration of meeting:</h4>" 
                                                      + this.information.time + "<br/><h4>Location:</h4>" + this.information.location + "<br/><h4>Meeting:</h4>" 
                                                      + this.information.meeting + cloudDetails, {coercionType: "html"});
                                                      
          var payload = { 
                                  "Pbe": "test",
                                  "Website": this.information.website, 
                                  "Crm": this.information.crm, 
                                  "Stage":this.information.stage, 
                                  "EngagementType":this.information.engagement,
                                  "Date":this.information.date,
                                  "Reason":this.information.reason,
                                  "Location": this.information.location,
                                  "Meeting": this.information.meeting,
                                  "Industry":this.cloudInfor.industry,
                                  "cloudStatus":this.cloudInfor.status,
                                  "cloudProvider":this.cloudInfor.provider,
                                  "consumption":this.cloudInfor.consumption,
                                  "WorkLoads": cloudDetails
                                };
                                
          var namespace = "insightsaddin-eh";
          var hubname = "insights-eh";
          var key_name = "send";
        
          // Token expires in 24 hours 
          var my_uri = 'https://' + namespace + '.servicebus.windows.net' + '/' + hubname + '/messages'; 
          

//          var expiry = Math.floor(new Date().getTime()/1000+3600*24); 
      
      
  //        var string_to_sign = encodeURIComponent(my_uri) + '\n' + expiry; 
    //      var hash = CryptoJS.HmacSHA256(string_to_sign, "HsoQY4aVbX2cOEhg5hqwfzQDSLLIKyvvIrNt/u2jU+k=");
      //    var hashInBase64 = CryptoJS.enc.Base64.stringify(hash);
          
          var token = 'SharedAccessSignature sr=insightsaddin-eh.servicebus.windows.net%2finsights-eh&sig=DRb9KkE%2fxSieq95%2f0MGVSdFGibeXzEUhLctLAGEFtg%3d&se=1470433911&skn=send';
                      
          var options = { 
            hostname: namespace + '.servicebus.windows.net', 
            path: '/' + hubname, 
            method: 'POST',
            headers: { 
              'Authorization': token, 
              'Content-Length': payload.length, 
              'Content-Type': 'application/atom+xml;type=entry;charset=utf-8',
              'Access-Control-Allow-Origin':'https://insightsaddinstaging.azurewebsites.net/' 
            } 
          };
          
          $http.post(my_uri, payload, options).then(function onSuccess(response) { 
                                                      console.log(response)
                                                    }, function onError(response) {
                                                      console.log(response)
                                                    });

          reset();
          showStatus();
          this.information = info;
          this.showBack = false;
          this.workloads = workLoads;
          this.skype = false;
        };
        this.canSkype = function() {
          this.skype = !this.skype;
        };

        this.goHome = function() {
          this.showTech = true;
          this.intelCloud = false;
          this.showCont = false;
          this.showMain = false;
          this.showArrow = false;
          showForm();
          hideStatus();
        };
    });

    function showStatus() {
      $("#form").hide();
      $("#submission").show();
    }

    function hideStatus() {
      $("#submission").hide();
    }

    function showForm() {
      $("#form").show();
    }

    function reset() {
      for (var key in info) {
        if (key != "pbe") {
          if (key === "meeting") {
            info[key] = "Skype";
          } else {
            info[key] = "";
          }
        }
      }
      for (var key in workLoads) {
        workLoads[key] = false;
      }
    }

})();
