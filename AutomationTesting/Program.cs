using FdxLeadAddressUpdate;
using FdxLeadAssignmentPlugin;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTesting
{
    class Program
    {
        static string environment = ConfigurationManager.AppSettings["Environment"];
        static string connectionString = ConfigurationManager.ConnectionStrings[environment].ConnectionString;
        static string crmInstanceName = ConfigurationManager.AppSettings["CRMInstanceName-" + environment];
        static Guid contactId = new Guid(ConfigurationManager.AppSettings["Contactid-" + environment]);
        static Guid accountId = new Guid(ConfigurationManager.AppSettings["Accountid-" + environment]);
        static Guid opportunityId = new Guid(ConfigurationManager.AppSettings["Opportunityid-" + environment]);
        static string smartCrmSyncServerName = ConfigurationManager.AppSettings["SmartCrmSyncServerName-" + environment];
        static string recordUrlTemplate = "https://{0}.crm.dynamics.com/main.aspx?etn={1}&extraqs=&id=%7b{2}%7d&newWindow=true&pagetype=entityrecord";
        static string smartCrmSyncServiceUrl = string.Format("http://{0}.1800dentist.com", smartCrmSyncServerName);

        static void Main(string[] args)
        {
            int indexOfUrl = connectionString.IndexOf("Url=") + 4;
            int indexOfUsername = connectionString.IndexOf("Username=") + 9;
            Console.WriteLine("Connecting to {0}", connectionString.Substring(indexOfUrl));
            Console.WriteLine("Using {0}", connectionString.Substring(indexOfUsername).Split(';')[0]);
            Console.Write("Password:");
            string password = GetPassword();
            Console.WriteLine(Environment.NewLine);
            connectionString = string.Format(connectionString, password);
            TestConnection();

            int optionSelected = -1;
            while (optionSelected != 0)
            {
                optionSelected = Menu();
                DateTime startTime = DateTime.Now;
                if (optionSelected == 0)
                {
                    return;
                }
                else if (optionSelected == 1)
                {
                    startTime = DateTime.Now;
                    CreateLeadWithStreetAddress();
                }
                else if (optionSelected == 2)
                {
                    startTime = DateTime.Now;
                    CreateLeadWithoutStreetAddress();
                }
                else if (optionSelected == 3)
                {
                    startTime = DateTime.Now;
                    CreateLeadWithAccount();
                }
                else if (optionSelected == 4)
                {
                    startTime = DateTime.Now;
                    CreateLeadWithNewAccount();
                }
                else if (optionSelected == 5)
                {
                    startTime = DateTime.Now;
                    TestCreateLeadAPI();
                }
                else if (optionSelected == 6)
                {
                    startTime = DateTime.Now;
                    TestCreateLeadAPIWithUrl();
                }
                else if (optionSelected == 7)
                {
                    startTime = DateTime.Now;
                    CreateOpportunity();
                }

                else if (optionSelected == 8)
                {
                    Console.Write("Provide the number of days to back date:");
                    int backDateByDays = int.Parse(Console.ReadLine());
                    startTime = DateTime.Now;
                    Entity opportunity = new Entity("opportunity", opportunityId);
                    opportunity["fdx_lastcustomofferapprovalapproveddate"] = DateTime.UtcNow.AddDays(-1 * backDateByDays);
                    CrmServiceClient conn = new CrmServiceClient(connectionString);
                    conn.OrganizationServiceProxy.Update(opportunity);
                }
                else if (optionSelected == 9)
                {
                    startTime = DateTime.Now;
                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='opportunity'>
                                            <attribute name='name' />
                                            <attribute name='createdon' />
                                            <attribute name='modifiedon' />
                                            <attribute name='customerid' />
                                            <attribute name='fdx_goldmineaccountnumber' />
                                            <attribute name='statuscode' />
                                            <attribute name='opportunityid' />
                                            <order attribute='createdon' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='fdx_goldmineaccountnumber' operator='like' value=' %' />
                                            </filter>
                                          </entity>
                                        </fetch>";
                    CrmServiceClient conn = new CrmServiceClient(connectionString);
                    EntityCollection opportunities = conn.OrganizationServiceProxy.RetrieveMultiple(new FetchExpression(fetchXml));
                }

                DateTime endTime = DateTime.Now;
                Console.WriteLine("Time Taken: " + endTime.Subtract(startTime).TotalSeconds);
            }

            //TestUpdateLeadAPI("CRM2099466ECBF0DU230", "DU2303181340");
            //TestUpdateLeadAPI("CRM3A089765E845Wilbu", "DU2303181340");
        }

        private static int Menu()
        {
            Console.WriteLine("Automated Testing Console");
            Console.WriteLine("1. Create Lead with Street Address");
            Console.WriteLine("2. Create Lead without Street Address");
            Console.WriteLine("3. Create Lead with Existing Account");
            Console.WriteLine("4. Create Lead with New Account");
            Console.WriteLine("5. Prospecting API Test - Create Lead - New Contact");
            Console.WriteLine("6. Prospecting API Test - Create Lead - Existing Contact");
            Console.WriteLine("7. Create Opportunity and Go to Proposal Sent Stage");
            Console.WriteLine("8. Back Date Approval Date by X Days");
            Console.WriteLine("0. Exit");
            Console.Write("Choose an option:");
            string optionString = Console.ReadLine();
            return int.Parse(optionString);
        }

        private static void TestConnection()
        {
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            WhoAmIResponse whoAmIResponse = (WhoAmIResponse)conn.Execute(new WhoAmIRequest());
            Console.WriteLine("Connection Successful!");
            Console.WriteLine(Environment.NewLine);
        }

        private static void CreateLeadWithStreetAddress()
        {
            string lastName = GetContactLastName();
            Entity lead = new Entity("lead");
            lead["firstname"] = "DU";
            lead["lastname"] = lastName;
            lead["telephone2"] = lastName;
            lead["leadsourcecode"] = new OptionSetValue(5); //Other
            lead["address1_line1"] = "6060 Centre Drive";
            lead["address1_line2"] = "7th Floor";
            lead["address1_city"] = "Los Angeles";
            lead["fdx_stateprovince"] = new EntityReference("fdx_state", new Guid("5F144EBB-8B7C-E611-80EB-5065F38B5171")); //California
            lead["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", new Guid("752DCA28-107D-E611-80EA-5065F38BE0E1")); //90045
            //lead["fdx_prospectgroup"] = new EntityReference("fdx_prospectgroup", "fdx_name", "FASTTRACK");
            //lead["fdx_pricelist"] = new EntityReference("pricelevel", "name", "PL_895");
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            Guid leadId = conn.OrganizationServiceProxy.Create(lead);
            LaunchInChrome(leadId, lead.LogicalName);
        }

        private static void CreateLeadWithoutStreetAddress()
        {
            string lastName = GetContactLastName();
            Entity lead = new Entity("lead");
            lead["firstname"] = "DU";
            lead["lastname"] = lastName;
            lead["telephone2"] = lastName;
            lead["leadsourcecode"] = new OptionSetValue(5); //Other
            lead["address1_city"] = "Los Angeles";
            lead["fdx_stateprovince"] = new EntityReference("fdx_state", new Guid("5F144EBB-8B7C-E611-80EB-5065F38B5171")); //California
            lead["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", new Guid("752DCA28-107D-E611-80EA-5065F38BE0E1")); //90045
            //lead["fdx_prospectgroup"] = new EntityReference("fdx_prospectgroup", "fdx_name", "FASTTRACK");
            //lead["fdx_pricelist"] = new EntityReference("pricelevel", "name", "PL_895");
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            Guid leadId = conn.OrganizationServiceProxy.Create(lead);
            LaunchInChrome(leadId, lead.LogicalName);
        }

        private static void CreateLeadWithAccount()
        {
            string lastName = GetContactLastName();
            Entity lead = new Entity("lead");
            lead["firstname"] = "DU";
            lead["lastname"] = lastName;
            lead["telephone2"] = lastName;
            lead["leadsourcecode"] = new OptionSetValue(5); //Other
            lead["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", new Guid("752DCA28-107D-E611-80EA-5065F38BE0E1")); //90045
            lead["parentaccountid"] = new EntityReference("account", accountId); //90045
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            Guid leadId = conn.OrganizationServiceProxy.Create(lead);
            LaunchInChrome(leadId, lead.LogicalName);
        }

        private static void CreateLeadWithNewAccount()
        {
            CrmServiceClient conn = new CrmServiceClient(connectionString);

            string lastName = GetContactLastName();
            Entity account = new Entity("account");
            account["name"] = "DU " + lastName;
            account["telephone1"] = lastName;
            account["address1_line1"] = "6060 Centre Drive";
            account["address1_line2"] = "7th Floor";
            account["address1_city"] = "Los Angeles";
            account["fdx_stateprovinceid"] = new EntityReference("fdx_state", new Guid("5F144EBB-8B7C-E611-80EB-5065F38B5171")); //California
            account["fdx_zippostalcodeid"] = new EntityReference("fdx_zipcode", new Guid("752DCA28-107D-E611-80EA-5065F38BE0E1")); //90045
            Guid accountId = conn.OrganizationServiceProxy.Create(account);

            Entity lead = new Entity("lead");
            lead["firstname"] = "DU";
            lead["lastname"] = lastName;
            lead["telephone2"] = lastName;
            lead["leadsourcecode"] = new OptionSetValue(5); //Other
            lead["address1_line1"] = "6060 Centre Drive";
            lead["address1_line2"] = "7th Floor";
            lead["address1_city"] = "Los Angeles";
            lead["fdx_stateprovince"] = new EntityReference("fdx_state", new Guid("5F144EBB-8B7C-E611-80EB-5065F38B5171")); //California
            lead["fdx_zippostalcode"] = new EntityReference("fdx_zipcode", new Guid("752DCA28-107D-E611-80EA-5065F38BE0E1")); //90045
            lead["parentaccountid"] = new EntityReference("account", accountId);

            Guid leadId = conn.OrganizationServiceProxy.Create(lead);
            LaunchInChrome(leadId, lead.LogicalName);
        }

        private static void TestCreateLeadAPI()
        {
            string phoneNumber = GetContactLastName();
            string contactName = "DU " + phoneNumber;
            bool createNewGMNo = true;
            string url = string.Format("{0}/api/lead/createlead?Zip=90045&Contact={1}&Phone1={2}&City=Los Angeles&State=CA&Address1=6060%20Center%20Drive", smartCrmSyncServiceUrl, contactName, phoneNumber);
            const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
            var uri = new Uri(url);
            var request = WebRequest.Create(uri);
            if (createNewGMNo)
            {
                request.Method = WebRequestMethods.Http.Post;
            }
            else
            {
                request.Method = WebRequestMethods.Http.Put;
            }
            request.ContentType = "application/json";
            request.ContentLength = 0;
            request.Headers.Add("Authorization", token);
            using (var getResponse = request.GetResponse())
            {
                DataContractJsonSerializer serializer =
                                       new DataContractJsonSerializer(typeof(Lead));
                //Console.WriteLine(getResponse.ToString());
                Lead leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
            }
        }

        private static void TestUpdateLeadAPI(string goldMineNumber, string companyName)
        {
            bool createNewGMNo = false;
            string url = string.Format("{0}/api/lead/updatelead?Zip=90045&Phone1=2903181507&City=Los Angeles&State=CA&Address1=6060%20Center%20Drive&AccountNo_in={1}&Company={2}", smartCrmSyncServiceUrl, WebUtility.UrlEncode(goldMineNumber), companyName);
            const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
            var uri = new Uri(url);
            var request = WebRequest.Create(uri);
            if (createNewGMNo)
            {
                request.Method = WebRequestMethods.Http.Post;
            }
            else
            {
                request.Method = WebRequestMethods.Http.Put;
            }
            request.ContentType = "application/json";
            request.ContentLength = 0;
            request.Headers.Add("Authorization", token);
            using (var getResponse = request.GetResponse())
            {
                DataContractJsonSerializer PutSerializer = new DataContractJsonSerializer(typeof(API_PutResponse));

                API_PutResponse leadObj = new API_PutResponse();
                leadObj = (API_PutResponse)PutSerializer.ReadObject(getResponse.GetResponseStream());
            }
        }

        private static void TestCreateLeadAPIWithUrl()
        {
            bool createNewGMNo = true;
            string url = smartCrmSyncServiceUrl + "/api/lead/createlead?Zip=90045&Contact=DU%201804051022&Phone1=1804051022&City=Los%20Angeles&State=CA&Address1=6060%20Center%20Drive";
            const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
            var uri = new Uri(url);
            var request = WebRequest.Create(uri);
            if (createNewGMNo)
            {
                request.Method = WebRequestMethods.Http.Post;
            }
            else
            {
                request.Method = WebRequestMethods.Http.Put;
            }
            request.ContentType = "application/json";
            request.ContentLength = 0;
            request.Headers.Add("Authorization", token);
            using (var getResponse = request.GetResponse())
            {
                DataContractJsonSerializer serializer =
                                       new DataContractJsonSerializer(typeof(Lead));
                //Console.WriteLine(getResponse.ToString());
                Lead leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
            }
        }

        private static void CreateOpportunity()
        {
            Entity opportunity = new Entity("opportunity");

            opportunity["parentcontactid"] = new EntityReference("contact", contactId);
            opportunity["parentaccountid"] = new EntityReference("account", accountId);
            opportunity["description"] = DateTime.Now.ToString();

            opportunity["estimatedclosedate"] = DateTime.UtcNow.AddDays(30);
            opportunity["fdx_nasurveydate"] = DateTime.UtcNow;
            opportunity["fdx_naapainpoint_neednewpatients"] = true;
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            Guid opportunityId = conn.OrganizationServiceProxy.Create(opportunity);
            Console.WriteLine("Opportunity Created!");

            List<Entity> parties = new List<Entity>();
            Entity activityparty = new Entity("activityparty");
            activityparty["partyid"] = new EntityReference("opportunity", opportunityId);
            parties.Add(activityparty);

            Entity appointment = new Entity("appointment");
            appointment["regardingobjectid"] = new EntityReference(opportunity.LogicalName, opportunityId);
            appointment["requiredattendees"] = parties.ToArray();
            appointment["scheduledstart"] = StartTime();
            appointment["scheduledend"] = EndTime();
            appointment["subject"] = "Hello";
            appointment["location"] = "office";
            appointment["fdx_appointmenttype"] = new OptionSetValue(756480001);
            Guid testAppointmentId = conn.OrganizationServiceProxy.Create(appointment);
            Console.WriteLine("Demo Appointment Set!");

            appointment = new Entity("appointment", testAppointmentId);
            SetStateRequest completeAppointment = new SetStateRequest();
            completeAppointment.EntityMoniker = new EntityReference("appointment", testAppointmentId);
            completeAppointment.Status = new OptionSetValue(3);
            completeAppointment.State = new OptionSetValue(1);
            conn.OrganizationServiceProxy.Execute(completeAppointment);
            Console.WriteLine("Demo Appointment Complete");

            LaunchInChrome(opportunityId, opportunity.LogicalName);
        }

        private static string GetContactLastName()
        {
            return string.Format("{0}{1}{2}{3}{4}", (DateTime.Now.Year % 2000).ToString("00"), DateTime.Now.Month.ToString("00"), DateTime.Now.Day.ToString("00"), DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00"));
        }

        private static string RemoveParenthesisFromGuid(Guid id)
        {
            return id.ToString().Replace("{", string.Empty).Replace("}", string.Empty);
        }

        private static void LaunchInChrome(Guid recordId, string entityName)
        {
            Process.Start("chrome.exe", string.Format(recordUrlTemplate, crmInstanceName, entityName, RemoveParenthesisFromGuid(recordId)));
        }

        /// <summary>
        /// Generates tomorrow's UTC time End Time
        /// </summary>
        /// <returns>End Time</returns>
        private static DateTime EndTime()
        {
            DateTime dt1 = new DateTime();
            dt1 = DateTime.UtcNow;
            dt1 = dt1.AddDays(1);
            dt1 = dt1.AddHours(1);
            dt1 = dt1.AddMinutes(1);
            return dt1;
        }

        /// <summary>
        ///  Generates tomorrow's UTC time Start Time
        /// </summary>
        /// <returns>Start Time</returns>
        private static DateTime StartTime()
        {
            DateTime dt = new DateTime();
            dt = DateTime.UtcNow;
            dt = dt.AddDays(1);
            dt = dt.AddMinutes(1);
            return dt;
        }

        public static string GetPassword()
        {
            string pwd = string.Empty;
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd = pwd.Remove(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd += i.KeyChar.ToString();
                    Console.Write("*");
                }
            }
            return pwd;
        }
    }
}
