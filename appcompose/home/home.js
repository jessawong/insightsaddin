(function(){
  'use strict';

  var app = angular.module('readHome', []);
  var selectTech = [{'type': 'Productivity','url':'../../images/office.png', 'alias':'US_DX_RISV_PROD_TEAM@microsoft.com', 'fadeUrl': '../../images/officeOverlay.png', 'prod': true}, 
                    {'type':'Modern Apps','url':'../../images/windows.png', 'alias': 'US_DX_RISV_APPS@microsoft.com', 'fadeUrl': '../../images/modernOverlay.png', 'modern': true}, 
                    {'type':'Intelligent Cloud', 'url':'../../images/cloud.png', 'alias':'US_DX_RISV_CLOUD@microsoft.com', 'fadeUrl': '../../images/cloudOverlay.png', 'intel': true}];
  var engageType = ['Briefing', 'Envisioning', 'ADS', 'Hackfest/PoC', 'Delivery', 'Other'];
  var info = {'pbe':'', 'website': '', 'date':'', 'time':'', 'reason':'', 'meeting':'Skype', 'location':'', 'engagement':'', 'crm': '', 'stage': ''};
  var crmStage = ['0%', '10%', '20%', '40%', '60%', '80%', '95%', '100%'];
  var time = ['30 min', '60 min', '90 min', '120 min', '2+ hours'];


    // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      info.pbe = Office.context.mailbox.userProfile.emailAddress;
      Office.context.mailbox.item.subject.setAsync("TE Request");
      $(".valign").hover(function() {
        $(this).stop().animate({"opacity":"0.5"}, "slow");
      },
      function() {
        $(this).stop().animate({"opacity":"1"}, "slow");
      });
    });
  };

    app.controller('FormController', function() {
        this.technology = selectTech;
        this.engagement = engageType;
        this.showTech = true;
        this.information = info;
        this.timeOptions = time;
        this.skype = false;
        this.ads = false;
        this.stage = crmStage;
        this.goBack = function() {
          this.skype = false;
          this.showTech = true;
          hideStatus();
          this.information["Technology"] = "";
        };
        this.setTech = function(option) {
          console.log(this.information);
          this.information["Technology"] = option.type;
          Office.context.mailbox.item.to.setAsync([option.alias]);
          this.showTech = false;
        };

        this.addRequest = function() {
          Office.context.mailbox.item.body.setAsync("<h4>Product's Website: </h4>" + this.information.website + "<br/><h4>CRM Link: </h4>" + this.information.crm
                                                      + "<br/><h4>Engagement Requested: </h4>" + this.information.engagement + "<br/><h4>Requested Date for Engagement:</h4>" 
                                                      + this.information.date + "<br/><h4>Reason:</h4>" + this.information.reason + "<br/><h4>Duration of meeting:</h4>" 
                                                      + this.information.time + "<br/><h4>Location:</h4>" + this.information.location + "<br/><h4>Meeting:</h4>" 
                                                      + this.information.meeting, {coercionType: "html"});
          reset();
          showStatus();
          this.information = info;
          this.skype = false;
        };
        this.canSkype = function() {
          this.skype = !this.skype;
        };
        this.setEngage = function(engageChoice) {
          this.information["engagement"] = engageChoice;
        };

        this.setTime = function(option) {
          this.information["time"] = option;
        }

        this.goHome = function() {
          this.showTech = true;
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
    }

})();
