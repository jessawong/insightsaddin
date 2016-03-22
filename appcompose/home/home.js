(function(){
  'use strict';

  var app = angular.module('readHome', []);
  var selectTech = [{'type': 'Productivity','url':'../../images/office.png', 'alias':'US_DX_RISV_PROD_TEAM@microsoft.com', 'fadeUrl': '../../images/officeOverlay.png', 'prod': true}, 
                    {'type':'Modern Apps','url':'../../images/windows.png', 'alias': 'US_DX_RISV_APPS@microsoft.com', 'fadeUrl': '../../images/modernOverlay.png', 'modern': true}, 
                    {'type':'Intelligent Cloud', 'url':'../../images/cloud.png', 'alias':'US_DX_RISV_CLOUD@microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true}];
  var engageType = ['Briefing', 'Envisioning', 'ADS', 'Hackfest/PoC', 'Delivery', 'Other'];
  var info = {'pbe':'', 'website': '', 'date':'', 'time':'', 'reason':'', 'meeting':'Skype', 'location':'', 'engagement':'', 'crm': '', 'stage': ''};
  var cloudInfo = {'status': '', 'provider':'', 'consumption':'', 'workloads':''};
  var crmStage = ['0%', '10%', '20%', '40%', '60%', '80%', '95%', '100%'];
  var time = ['30 min', '60 min', '90 min', '120 min', '2+ hours'];
  var status = ['New', 'Experimenting', 'Hybrid', 'Running'];
  var provider = ['Azure', 'AWS', 'Google', 'Other'];
  var consumptionLevel = ['<25k', '25k-99k', '100k-499k', '500k+'];
  var workLoads = {'Compute': false, 'Web & Mobile': false, 'Data & Storage': false,
                    'Analytics': false, 'Internet of Things': false,
                    'Networking': false, 'Media & CDN': false, 'Hybrid Integration': false,
                    'Identity & Access Management': false,
                    'Dev Services': false, 'Management & Security': false};


    // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      info.pbe = Office.context.mailbox.userProfile.emailAddress;
      Office.context.mailbox.item.subject.setAsync("TE Request");
      $(".valign").hover(function() {
        $(this).stop().animate({"opacity":"0.5"}, "fast");
      },
      function() {
        $(this).stop().animate({"opacity":"1"}, "fast");
      });
    });
  };

    app.controller('FormController', function($scope) {
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
            hideStatus();
            this.information["Technology"] = "";
            reset();
          }
          this.showSubmit = false;
        };
        this.setTech = function(option) {
          console.log(this.workloads);
          this.information["Technology"] = option.type;
          Office.context.mailbox.item.to.addAsync([option.alias]);
          this.showTech = false;
          this.showMain = true;
          this.showBack = true;
          this.showSubmit = true;
          if (option.type === "Intelligent Cloud") {
            this.intelCloud = true;
            this.showArrow = true;
            this.showSubmit = false;
          }
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
          var cloudDetails = '';
          if (this.intelCloud) {
            cloudDetails = "<br/><h4>Industry/Verticle: </h4>" + this.cloudInfor.industry + "<br/><h4>Cloud Status: </h4>" + this.cloudInfor.status
                            + "<br/><h4>Cloud Provider: </h4>" + this.cloudInfor.provider + "<br/><h4>Annual Consumption: </h4>" + this.cloudInfor.consumption
                            + "<br/><h4>Workloads: </h4>";
            for (var key in this.workloads) {
              if (this.workloads[key]) {
                cloudDetails += key + "<br/>";
              }
            }
          }
          Office.context.mailbox.item.body.setAsync("<h4>Product's Website: </h4>" + this.information.website + "<br/><h4>CRM Link: </h4>" + this.information.crm
                                                      + "<br/><h4>Stage:</h4>" + this.information.stage + "<br/><h4>Engagement Requested: </h4>" 
                                                      + this.information.engagement + "<br/><h4>Requested Date for Engagement:</h4>" + this.information.date 
                                                      + "<br/><h4>Reason:</h4>" + this.information.reason + "<br/><h4>Duration of meeting:</h4>" 
                                                      + this.information.time + "<br/><h4>Location:</h4>" + this.information.location + "<br/><h4>Meeting:</h4>" 
                                                      + this.information.meeting + cloudDetails, {coercionType: "html"});
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
