using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using Microsoft.TeamFoundation.TestManagement.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {

            // provide collection url of tfs
            var collectionUri = "http://pankagar-hp2:8080/tfs/DefaultCollection";
            VssCredentials connection = new VssCredentials(true);
            var testPlanId = 1;
            var testSuiteId = 2;
            var projectName = "agile";

            // create workitemtracking client
            var _witClient = new WorkItemTrackingHttpClient(new Uri(collectionUri), connection);

            // upload attachment
            FileStream uploadStream = File.Open(@"C: \Users\pankagar\Downloads\Canvas.png", FileMode.Open, FileAccess.Read);
            var attachmentObject = _witClient.CreateAttachmentAsync(uploadStream, "Canvas.png", "Simple").Result;

            // create a patchdocument object
            JsonPatchDocument json = new JsonPatchDocument();
            
            // create a new patch operation for title field
            JsonPatchOperation patchDocument1 = new JsonPatchOperation();
            patchDocument1.Operation = Operation.Add;
            patchDocument1.Path = "/fields/System.Title";
            patchDocument1.Value = "Testing Rest Api";
            json.Add(patchDocument1);

            // create a new patch operation for priority field
            JsonPatchOperation patchDocument2 = new JsonPatchOperation();
            patchDocument2.Operation = Operation.Add;
            patchDocument2.Path = "/fields/Microsoft.VSTS.Common.Priority";
            patchDocument2.Value = "2";
            json.Add(patchDocument2);

                      
            // create testbasehelper object
            TestBaseHelper helper = new TestBaseHelper();
            // create testbase object to utilize teststep helpers
            ITestBase tb = helper.Create();

            // create 2 test steps ts1 and ts2
            ITestStep ts1 = tb.CreateTestStep();
            ITestStep ts2 = tb.CreateTestStep();
            ts1.Title = "title -> title1";
            ts2.Title = "title -> title2";
            ts1.ExpectedResult = "expected1";
            ts2.ExpectedResult = "expected2";
            ts1.Description = "description1";
            ts2.Description = "description2";
            // adding attachment to step1
             ts1.Attachments.Add(ts1.CreateAttachment(attachmentObject.Url, "CanvasImage"));

            // add your steps actions to testbase object
            tb.Actions.Add(ts1);
            tb.Actions.Add(ts2);

            // update json based on all actions (including teststeps and teststep attachemnts) 
            json = tb.SaveActions(json);

            var xml = "";
            /* getting xml for teststeps
            xml = tb.GenerateXmlFromActions();
            */

            // create Test Case work item using all test steps: ts1 and ts2
            var testCaseObject = _witClient.CreateWorkItemAsync(json, projectName, "Test Case").Result;
            int testCaseId = Convert.ToInt32(testCaseObject.Id);
            
            // get Test Case using all relations
            testCaseObject = _witClient.GetWorkItemAsync(testCaseId, null, null, WorkItemExpand.Relations).Result;

            // update Test Case
            if (testCaseObject.Fields.ContainsKey("Microsoft.VSTS.TCM.Steps"))
            {
                
                xml = testCaseObject.Fields["Microsoft.VSTS.TCM.Steps"].ToString();
                tb = helper.Create();

                // create tcmattachemntlink object from workitem relation, teststep helper will use this
                IList<TestAttachmentLink> tcmlinks= new List<TestAttachmentLink>();
                foreach (WorkItemRelation rel in testCaseObject.Relations)
                {
                    TestAttachmentLink tcmlink = new TestAttachmentLink();
                    tcmlink.Url = rel.Url;
                    tcmlink.Attributes = rel.Attributes;
                    tcmlink.Rel = rel.Rel;
                    tcmlinks.Add(tcmlink);
                    
                }

                // load teststep xml and attachemnt links
                tb.LoadActions(xml, tcmlinks);
                
                ITestStep ts;
                //updating 1st test step
                ts = (ITestStep)tb.Actions[0];
                ts.Title = "title -> title11";
                ts.ExpectedResult = "expected11";

                //removing 2ns test step
                tb.Actions.RemoveAt(1);

                //adding new test step
                ITestStep ts3 = tb.CreateTestStep();
                ts3.Title = "title -> title3";
                ts3.ExpectedResult = "expected3";
                tb.Actions.Add(ts3);

                JsonPatchDocument json2 = new JsonPatchDocument();
                // update json based on all new changes ( updated step xml and attachments)
                json2 = tb.SaveActions(json2);

                // update testcase wit using new json
                testCaseObject = _witClient.UpdateWorkItemAsync(json2, testCaseId).Result;

                /* Note : If you want to remove attachment then create new patchOperation, details are available here :
                          https://www.visualstudio.com/en-us/docs/integrate/api/wit/work-items#remove-an-attachment
                */
            }
        }
    }
}
