/// <reference path="../node_modules/vss-web-extension-sdk/typings/vss.d.ts" />
/// <reference path="../node_modules/vss-web-extension-sdk/typings/tfs.d.ts" />
/// <reference path="../node_modules/vss-web-extension-sdk/typings/rmo.d.ts" />
"use strict";

VSS.init({
		  usePlatformScripts: true, 
		  usePlatformStyles: true
		});
	  
		VSS.ready(function() {
		  // Start using VSS
		

		var apis = ["VSS/Controls", "VSS/Service", "TFS/TestManagement/RestClient", "TFS/TestManagement/Contracts", "TFS/TestManagement/Utils", "TFS/WorkItemTracking/RestClient", "TFS/WorkItemTracking/Contracts", "VSS/WebApi/Contracts"];

		VSS.require(apis, (Controls, VSS_Service, TFS_Test_WebApi, TestContracts, Utils, WorkItemRestClient, WitContract, VSS_Common_Contracts) => {
						
			// title
			var patchDocument1 = {
				from: null,
				op: VSS_Common_Contracts.Operation.Add,
				path: "/fields/System.Title",
				value: "Testing Rest Api"
			};
		
			//priority
			var patchDocument2 = {
				from: null,
				op: VSS_Common_Contracts.Operation.Add,
				path: "/fields/Microsoft.VSTS.Common.Priority",
				value: "2"
			};
			
			// create testbasehelper object and then create testbase object to utilize teststep helpers
			var helper = new Utils.TestBaseHelper();
            var tb = helper.create();
			
			// create 2 test steps step1 and step2
            var step1 = tb.createTestStep();
            var step2 = tb.createTestStep();
            step1.setTitle("action1");
            step2.setTitle("title ->  title1");
            step1.setExpectedResult("expect1");
            step2.setExpectedResult("expected2");
            step1.description = "desc1";
            step2.description = "desc2";
			step1.stepType = Utils.TestStepType.Validate;
			
			// attachment objects got from createAttachment call
            var attachment1 = "http://pankagar-hp2:8080/tfs/DefaultCollection/_apis/wit/attachments/955e7a7d-a6c4-4af6-bfe1-e909bd93db63?fileName=Canvas%20Image";
            var attachment2 = "http://pankagar-hp2:8080/tfs/DefaultCollection/_apis/wit/attachments/ddcc418c-dffa-4008-b3db-486395dccc80?fileName=SystemInformation";
            var attachment3 = "http://pankagar-hp2:8080/tfs/DefaultCollection/_apis/wit/attachments/ce7e0a01-46da-40fb-a957-2ece3aae7589?fileName=Nuget%20Package";
            var attachment4 = "http://pankagar-hp2:8080/tfs/DefaultCollection/_apis/wit/attachments/dbfbac3f-3d8a-4dda-b2d9-5990bdb67671?fileName=Xt%20Session%20report";
            var attachment5 = "http://pankagar-hp2:8080/tfs/DefaultCollection/_apis/wit/attachments/4be2ee2e-e3ee-4771-86c9-ef349da60d82?fileName=screen%20shot";
            
			// adding attachment to steps
			step1.attachments.push(step1.createAttachment(attachment1, "canvas image"));
            step1.attachments.push(step1.createAttachment(attachment2, "System info"));
            step2.attachments.push(step2.createAttachment(attachment3, "nuget package"));
            
			// add your steps actions to tesebase object
			tb.actions.push(step1);
            tb.actions.push(step2);
			
            /* getting xml for teststeps
            var xml = tb.generateXmlFromActions();
            */
				
			var json = [patchDocument1, patchDocument2];
			
			// update json based on all actions (including teststeps and teststep attachemnts) 
			json = tb.saveActions(json);
			
			// create workitemtracking client
			var witClient = VSS_Service.getClient(WorkItemRestClient.WorkItemTrackingHttpClient3);
			
			// create Test Case
			witClient.createWorkItem(json, VSS.getWebContext().project.name, "Test Case").then(function(workitem) {
				console.log("create workitem id is " + workitem.id);
				
				// get Test Case using all relations
				witClient.getWorkItem(workitem.id, null, null, WitContract.WorkItemExpand.Relations).then(function(workitem)
				{
					console.log("get workitem id is " + workitem.id);
					
					var xml = workitem.fields["Microsoft.VSTS.TCM.Steps"].toString();
					
					//create tcmattachemntlink object from workitem relation, teststep helper will use this
					var attachmentLinks = [];
					for (var i = 0; i < workitem.relations.length; i++) {
                    var rel = workitem.relations[i];
                    var attachmentLink = {
                        rel: rel.rel,
                        url: rel.url,
                        attributes: rel.attributes,
                        title: ""
                    };
                    attachmentLinks.push(attachmentLink);
					}
					
                var helper2 = new Utils.TestBaseHelper();
                var tb2 = helper2.create();
				
				// load teststep xml and attachemnt links
                tb2.loadActions(xml, attachmentLinks);
				
				// updating 1st test step, removing 2nd test step and adding new 3rd teststep
                var ts1 = tb2.actions[0];
                var ts2 = tb2.actions[1];
                ts1.setTitle("action11");
                ts1.setExpectedResult("expect11");
                ts1.attachments.push(ts1.createAttachment(attachment4, "session report"));
                ts1.attachments.splice(0, 2);
                tb2.actions.splice(1, 1);
                var ts3 = tb2.createTestStep();
                ts3.setExpectedResult("expect3");
                ts3.setTitle("action3");
                ts3.attachments.push(ts3.createAttachment(attachment5, "screenshot"));
                tb2.actions.push(ts3);
				
				// update json based on all new changes ( updated step xml and attachments)
				json= [];
				json = tb2.saveActions(json);
							
					
				// updating work item with new json changes
				witClient.updateWorkItem(json, workitem.id).then(function (workitem) {
                    console.log("update workitem id is " + workitem.id);
                });
				
				/* Note : If you want to remove attachment then create new patchdocument, details are available here : https://www.visualstudio.com/en-us/docs/integrate/api/wit/work-items#remove-an-attachment				
				*/
				
				});
			});
			
	});
});	 
	  

//# sourceMappingURL=app.js.map