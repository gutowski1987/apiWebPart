import Axios from "axios";
import styles from './DisplayPageWebPart.module.scss';

window.url = window.location.href;
window.arrUrl = window.url.split('/');
window.starts;
window.arrUrl.forEach(function(e) {
    if(e.startsWith('PB')) {
        window.starts = e;
    }
});
window.projectData = window.starts;

window.today = new Date();
var dd = String(today.getDate()).padStart(2, '0');
var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = today.getFullYear();
window.today = dd + '/' + mm + '/' + yyyy;

$(window).ready(
    projectDetails(),
    projectObjective(),
    executiveStatus(),
    ragStatus(),
    keyMilestones(),
    completedThisPeriod(),
    focusForNextPeriod(),
    projectRisks(),
    projectIssues()
);
function projectDetails() {
    $.ajax({
      type: 'GET',
      headers: {
        'accept': 'application/json;odata=verbose'
      },
      url: "/sites/ictprojects/_api/web/lists(guid\'e2e2998e-1f81-4c67-82c6-3fd32a22ca13\')/items?Title,SponsorId,Project_x0020_ID,Project_x0020_Status,Gate_x0020_Stage,ICT_x002d_LeadId,ICT_x0020_Workstream,ICT_x0020_Watchlist,Modified&$filter=Project_x0020_ID eq \'" + projectData + "\'",
      success: 
    function(data){
          $(data.d.results).each(function(){
            $('#projectDetails').append("<h2>Project Summary - " + today + "</h2><h2>" + this.Project_x0020_ID + " - " + this.Title + "</h2><h3>Project Details</h3><table class=\"" + styles.WithoutBorder + "\"><thead><tr><th>Project Sponsor</th><th>Project Status</th><th>Gate Stage</th><th>ICT Lead</th><th>Workstream</th><th>Watch List</th><th>Last Updated</th></tr></thead><tbody><tr><td id=\"sponsor\"></td><td>" + ifNull(this.Project_x0020_Status) + "</td><td>" + ifNull(this.Gate_x0020_Stage) + "</td><td id=\"ictlead\"></td><td>" + ifNull(this.ICT_x0020_Workstream) + "</td><td>" + ifNull(this.ICT_x0020_Watchlist) + "</td><td>" + changeDate(this.Modified) + "</td></tr></tbody></table>");
        callPerson(this.SponsorId, 'sponsor');
        callPerson(this.ICT_x002d_LeadId, 'ictlead');
        });
      }
    });
}
function projectObjective() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'e2e2998e-1f81-4c67-82c6-3fd32a22ca13\')/items?Success&$filter=Project_x0020_ID eq \'" + projectData + "\'",
              success: 
              function(data){
                    $(data.d.results).each(function(){
                      $('#projectObjective').append("<h3>Project Objective</h3><div class=\"" + styles.displayRichText + "\">" + ifNull(this.Success) + "</div>");
                  });
                }
            });
}
function executiveStatus() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'81395a13-25a3-487b-8d50-306aa4ddc78d\')/items?Exec_x0020_Status&$filter=Project_x0020_ID eq \'" + projectData + "\'&$top=1&$orderby=Created desc",
              success: 
            function(data){
                  $(data.d.results).each(function(){
                    $('#executiveStatus').append("<h3>Executive Status</h3><div class=\"" + styles.displayRichText + "\">" + ifNull(this.Exec_x0020_Status) + "</div>");
                });
              }
            });
}
function ragStatus() {
  $.ajax({
      type: 'GET',
      headers: {
          'accept': 'application/json;odata=verbose'
      },
      url: "/sites/ictprojects/_api/web/lists(guid\'494f3987-42aa-403a-bed6-d33eff0743b5\')/items?ProjectCostCode,ProjectCostIndicator,ProjectScopeCode,ProjectScopeIndicator,ProjectResourceCode,ProjectResourceIndicator,ProjectScheduleCode,ProjectScheduleInducator&$filter=Project_x0020_ID eq \'" + projectData + "\'&$orderby=Created desc&$top=1",
      success: 
      function(data){
          $(data.d.results).each(function(){
              icon(this.ProjectCostIndicator, '#iconCost');
              icon(this.ProjectResourceIndicator, '#iconResource');
              icon(this.ProjectScopeIndicator, '#iconScope');
              icon(this.ProjectScheduleIndicator, '#iconSchedule');
              color(this.ProjectCostCode, '#colorCost');
              color(this.ProjectResourceCode, '#colorResource');
              color(this.ProjectScopeCode, '#colorScope');
              color(this.ProjectScheduleCode, '#colorSchedule');
          });
      }
  });
}
function keyMilestones() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'c55aed24-afe4-4c59-afae-a12e8d55bf2d\')/items?txtMilestoneType,Date1,Comment1,txtGate&$filter=Project_x0020_ID eq \'" + projectData + "\'&$top=12&$orderby=MOrder asc",
              success: 
              function(data){
                  $(data.d.results).each(function(){
                    $('#keyMilestones').append("<tr><td>" + ifNull(this.txtGate) + " - " + ifNull(this.txtMilestoneType) + "</td><td><div>" + ifNull(this.Comment1) + "</div></td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date1) + "</td></tr>");
                });
              }
            });
}
function completedThisPeriod() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'15bed8d1-278b-4427-9881-e54a5bd2bef1\')/items?Task,Date1&$filter=(Project_x0020_ID eq \'" + projectData + "\') and (TaskCategory eq \'Completed This Period\')",
              success: 
            function(data){
                  $(data.d.results).each(function(){
                    $('#completedThisPeriod').append("<tr><td>" + ifNull(this.Task) + "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date1) + "</td></tr>");
                });
              }
            });
}
function focusForNextPeriod() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'15bed8d1-278b-4427-9881-e54a5bd2bef1\')/items?Task,Date1&$filter=(Project_x0020_ID eq \'" + projectData + "\') and (TaskCategory eq \'Focus for Next Period\')",
              success: 
            function(data){
                  $(data.d.results).each(function(){
                    $('#focusForNextPeriod').append("<tr><td>" + ifNull(this.Task) + "</td><td class=\"" + styles.narrow + "\">" + changeDate(this.Date1) + "</td></tr>");
                });
              }
            });
}
function projectRisks() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'fd4111fb-ef23-4638-b77b-fcde30a624cb\')/items?RiskReference,ProjectRiskDescription,GrossImpact0,RiskMitigation,NetImpact,RiskOwner&$filter=Project_x0020_ID eq \'" + projectData + "\'",
              success: 
            function(data){
                  $(data.d.results).each(function(){
                    $('#projectRisks').append("<tr><td>" + ifNull(this.RiskReference) + "</td><td>" + ifNull(this.ProjectRiskDescription) + "</td><td class=\"" + styles.narrow + "\">" + ifNull(this.GrossImpact0) + "</td><td>" + ifNull(this.RiskMitigation) + "</td><td class=\"" + styles.narrow + "\">" + ifNull(this.NetImpact) + "</td><td id=\"riskOwner\"></td></tr>");
                callPerson(this.RiskOwnerId, 'riskOwner');
                });
              }
            });
}
function projectIssues() {
            $.ajax({
              type: 'GET',
              headers: {
                'accept': 'application/json;odata=verbose'
              },
              url: "/sites/ictprojects/_api/web/lists(guid\'93bb9701-c527-483b-af1c-e337ef9cd523\')/items?IssueReference,ProjectIssue,EscalationLevel,ActionsToResolve,ActionsDue,ActionOwnerId&$filter=Project_x0020_ID eq \'" + projectData + "\'",
              success: 
            function(data){
                  $(data.d.results).each(function(){
                    $('#projectIssues').append("<tr><td>" + ifNull(this.IssueReference) + "</td><td>" + ifNull(this.ProjectIssue) + "</td><td>" + ifNull(this.EscalationLevel) + "</td><td>" + ifNull(this.ActionsToResolve) + "</td><td>" + changeDate(this.ActionsDue) + "</td><td id=\"actionOwner\"></td></tr>");
                callPerson(this.ActionOwnerId, 'actionOwner');
                });
              }
            });
}
function color(data, where) { 
        if(data === 'Green') {
            $(where).addClass("ms-bgColor-greenLight");
        } else if(data === 'Amber') {
            $(where).addClass("ms-bgColor-yellow");
        } else if(data === 'Red') {
            $(where).addClass("ms-bgColor-red");
        }
}
function icon(data, where) {
        if(data === '=') {
            $(where).append("<i class=\"ms-Icon ms-Icon--CalculatorEqualTo\" aria-hidden=\"true\"></i>");
        } else if(data === '<') {
            $(where).append("<i class=\"ms-Icon ms-Icon--ArrowTallDownRight\" aria-hidden=\"true\"></i>");
        } else if(data === '>') {
            $(where).append("<i class=\"ms-Icon ms-Icon--ArrowTallUpRight\" aria-hidden=\"true\"></i>");
        }
}
function callPerson(identificator, where) {
        var resultElement = document.getElementById(where);
        resultElement.innerHTML = '';
        if(identificator == null) {
          return
        } else {
          Axios.get("https://intu.sharepoint.com/sites/ictprojects/_api/web/getuserbyid(" + identificator + ")")
          .then(function (response) {
            resultElement.innerHTML = response.data.Title;
          })  
        }
}
function changeDate(input) {
  if(input == null) {
    return " ";
  } else {
    var stringDate = input.slice(0,10).split("-");
    var readyDate = stringDate[2] + "-" + stringDate[1] + "-" + stringDate[0];
    return readyDate;
  }
}
function ifNull(val) {
  if(val == null) {
    return " ";
  } else {
    return val;
  }
}