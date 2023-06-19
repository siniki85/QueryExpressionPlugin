using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueryExpressionAccount
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Query Expression Demo......");
            try
            {
                string conn = getConnectionString();

                Console.WriteLine("Connection Established......");

                CrmServiceClient service = new CrmServiceClient(conn);

                if (service.IsReady)
                {
                    Console.WriteLine("Select Entity to Query:");
                    Console.WriteLine("1. View Account Details For Revenue > 100000 :");
                    Console.WriteLine("2. Create Lead:");
                    Console.WriteLine("3. Update Lead:");
                    Console.WriteLine("4. Delete Lead:");
                    Console.WriteLine("5. Qualify Lead:");
                    Console.WriteLine("6. Get Total Amount in Order & related Quote: ");

                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            getAccountDetails(service);
                            break;

                        case "2":
                            createLead(service);
                            break;

                        case "3":
                            updateLead(service);
                            break;

                        case "4":
                            deleteLead(service);
                            break;

                        case "5":
                            qualifyLead(service);
                            break;

                        case "6":
                            getTotalAmountOrderAndRelatedQuote(service);
                            break;

                        default:
                            Console.WriteLine("Invalid choice.");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Failed to connect with dynamics 365 crm");
                }
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();

            }
            catch (Exception e)
            {
                Console.WriteLine("Error Message : " + e);
            }
        }

        public static string getConnectionString()
        {
            string sUserKey = getAppSettingKey("UserKey");
            string sUserPassword = getAppSettingKey("UserPassword");
            string sEnvironment = getAppSettingKey("Environment");

            return $@" Url = {sEnvironment};AuthType = OAuth;UserName = {sUserKey}; Password = {sUserPassword};AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;LoginPrompt=Auto; RequireNewInstance = True";
        }

        public static string getAppSettingKey(string key)
        {
            return System.Configuration.ConfigurationManager.AppSettings[key].ToString();
        }

        public static void getAccountDetails(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("account");

            ColumnSet columnSet = new ColumnSet("name", "address1_city", "revenue");
            query.ColumnSet = columnSet;

            ConditionExpression condition = new ConditionExpression("revenue", ConditionOperator.GreaterThan, 100000);
            FilterExpression filter = new FilterExpression(LogicalOperator.Or);
            filter.AddCondition(condition);
            query.Criteria = filter;

            OrderExpression order = new OrderExpression("name", OrderType.Ascending);
            query.Orders.Add(order);

            query.PageInfo = new PagingInfo { Count = 100, PageNumber = 1, PagingCookie = null, ReturnTotalRecordCount = true };

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity account in collection.Entities)
            {
                string name = account.GetAttributeValue<string>("name");
                string city = account.GetAttributeValue<string>("address1_city");
                Money revenue = account.GetAttributeValue<Money>("revenue");

                Console.WriteLine("Account Name: " + name + "\tCity: " + city + "\tRevenue: " + revenue.Value);
                Console.WriteLine("-------------*************------------");
            }
        }

        public static void createLead(CrmServiceClient service)
        {
            var lead = new Entity("lead");

            Console.WriteLine("Enter Subject Name: ");
            string subject = Console.ReadLine();
            lead["subject"] = subject;

            Console.WriteLine("Enter First Name: ");
            string fname = Console.ReadLine();
            lead["firstname"] = fname;

            Console.WriteLine("Enter Last Name: ");
            string lname = Console.ReadLine();
            lead["lastname"] = lname;

            Console.WriteLine("Enter Job Title: ");
            string jobTitle = Console.ReadLine();
            lead["jobtitle"] = jobTitle;

            Console.WriteLine("Enter Business Phone: ");
            string businessPhone = Console.ReadLine();
            lead["telephone1"] = businessPhone;

            Console.WriteLine("Enter Mobile No.: ");
            string mobilePhone = Console.ReadLine();
            lead["mobilephone"] = mobilePhone;

            Console.WriteLine("Enter Email: ");
            string email = Console.ReadLine();
            lead["emailaddress1"] = email;

            Console.WriteLine("Enter Company Name: ");
            string company = Console.ReadLine();
            lead["companyname"] = company;

            Console.WriteLine("Enter Website URL: ");
            string websiteurl = Console.ReadLine();
            lead["websiteurl"] = websiteurl;

            Console.WriteLine("Enter Address: ");
            string address1 = Console.ReadLine();
            lead["address1_line1"] = address1;

            Console.WriteLine("Enter City: ");
            string city1 = Console.ReadLine();
            lead["address1_city"] = city1;

            Console.WriteLine("Enter Country: ");
            string country1 = Console.ReadLine();
            lead["address1_country"] = country1;

            Console.WriteLine("Enter Annual Revenue: ");
            decimal estAmount = Convert.ToDecimal(Console.ReadLine());
            lead["revenue"] = new Money(estAmount);

            Console.WriteLine("Enter Lead Source(1,2,3): 1. Advertisement\n 2. Employee Referral\n 3. Seminar ");
            int leadsourcecode = Convert.ToInt32(Console.ReadLine());
            lead["leadsourcecode"] = new OptionSetValue(leadsourcecode);

            Console.WriteLine("Creating Lead.........");

            Guid leadId = service.Create(lead);

            Console.WriteLine("Lead created.........");

            QualifyLeadRequest qualifyRequest = new QualifyLeadRequest
            {
                CreateOpportunity = true,
                CreateContact = true,
                CreateAccount = true,
                LeadId = new EntityReference("lead", leadId),
                Status = new OptionSetValue(3),
            };

            var qualifyResponse = (QualifyLeadResponse)service.Execute(qualifyRequest);

            Console.WriteLine("Lead qualified successfully.");

            ColumnSet leadColumns = new ColumnSet("subject", "firstname", "emailaddress1", "companyname", "address1_country", "revenue");
            Entity retrievedLead = service.Retrieve("lead", leadId, leadColumns);

            Console.WriteLine("Lead Data: ");
            Console.WriteLine("Subject: " + retrievedLead.GetAttributeValue<string>("subject"));
            Console.WriteLine("First Name: " + retrievedLead.GetAttributeValue<string>("firstname"));
            Console.WriteLine("Email: " + retrievedLead.GetAttributeValue<string>("emailaddress1"));
            Console.WriteLine("Company Name: " + retrievedLead.GetAttributeValue<string>("companyname"));
            Console.WriteLine("Website URL: " + retrievedLead.GetAttributeValue<string>("websiteurl"));
            Console.WriteLine("Country: " + retrievedLead.GetAttributeValue<string>("address1_country"));
            Console.WriteLine("Revenue: " + retrievedLead.GetAttributeValue<Money>("revenue").Value);
            Console.ReadLine();
        }

        public static void updateLead(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("lead");

            ColumnSet columnSet = new ColumnSet("leadid", "firstname", "lastname");
            query.ColumnSet = columnSet;

            ConditionExpression condition = new ConditionExpression("revenue", ConditionOperator.GreaterThan, 100000);
            query.Criteria.AddCondition(condition);

            query.PageInfo = new PagingInfo { Count = 20, PageNumber = 1 };

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity lead in collection.Entities)
            {
                Guid leadid = lead.GetAttributeValue<Guid>("leadid");
                string fname = lead.GetAttributeValue<string>("firstname");
                string lname = lead.GetAttributeValue<string>("lastname");

                Console.WriteLine("Lead ID: " + leadid + "\tFull Name: " + fname + " " + lname);
                Console.WriteLine("-----------------------------------------------------------");
            }

            Console.WriteLine("Enter Lead Id: ");
            var updateID = Console.ReadLine();

            var updateGuid = new Guid(updateID);

            Entity retrieveLead = service.Retrieve("lead", updateGuid, new ColumnSet());

            Console.WriteLine("Enter Subject Which You Want To Update: ");
            string subject = Console.ReadLine();
            retrieveLead["subject"] = subject;

            Console.WriteLine("Enter Annual Revenue You Want To Update: ");
            decimal revenue = Convert.ToDecimal(Console.ReadLine());
            retrieveLead["revenue"] = new Money(revenue);

            service.Update(retrieveLead);

            Console.WriteLine("Lead Updated Successfully");
        }

        public static void deleteLead(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("lead");

            ColumnSet columnSet = new ColumnSet("leadid", "firstname", "lastname");
            query.ColumnSet = columnSet;

            ConditionExpression condition = new ConditionExpression("revenue", ConditionOperator.GreaterThan, 50000);
            query.Criteria.AddCondition(condition);

            query.PageInfo = new PagingInfo { Count = 20, PageNumber = 1 };

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity lead in collection.Entities)
            {
                Guid leadid = lead.GetAttributeValue<Guid>("leadid");
                string fname = lead.GetAttributeValue<string>("firstname");
                string lname = lead.GetAttributeValue<string>("lastname");

                Console.WriteLine("Lead ID: " + leadid + "\tFull Name: " + fname + " " + lname);
                Console.WriteLine("-------------*************------------");
            }

            Console.WriteLine("Enter Lead Id: ");
            var deleteID = Console.ReadLine();

            var deleteGuid = new Guid(deleteID);

            service.Delete("lead", deleteGuid);

            Console.WriteLine("Lead Deleted Successfully...........");
        }

        public static void qualifyLead(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("lead");

            ColumnSet columnSet = new ColumnSet("leadid", "statuscode", "subject");
            query.ColumnSet = columnSet;

            ConditionExpression condition = new ConditionExpression("statuscode", ConditionOperator.Equal, 1);
            query.Criteria.AddCondition(condition);

            query.PageInfo = new PagingInfo { Count = 5, PageNumber = 1 };

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity lead1 in collection.Entities)
            {
                Guid leadid = lead1.GetAttributeValue<Guid>("leadid");
                var statuscode = lead1.FormattedValues["statuscode"].ToString();
                string subject = lead1.GetAttributeValue<string>("subject");

                Console.WriteLine("Lead Id: " + leadid + "\n\tStatus: " + statuscode + "\n\tSubject: " + subject);
                Console.WriteLine("---------------------------------------------------");
            }

            Console.WriteLine("Enter Lead Id You Want To Qualify: ");
            var qualifyLeadId = Console.ReadLine();

            Guid qualifyLeadGuid = new Guid(qualifyLeadId);

            QualifyLeadRequest qualifyLeadRequest = new QualifyLeadRequest
            {
                CreateAccount = true,
                CreateContact = true,
                CreateOpportunity = true,
                LeadId = new EntityReference("lead", qualifyLeadGuid),
                Status = new OptionSetValue(3),
            };

            var qualifyResponse = (QualifyLeadResponse)service.Execute(qualifyLeadRequest);

            Console.WriteLine("Lead Qualified Successfully..........");
        }

        public static void getTotalAmountOrderAndRelatedQuote(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("salesorder");

            query.ColumnSet = new ColumnSet("name", "totalamount");

            LinkEntity quoteLink = new LinkEntity("salesorder", "quote", "quoteid", "quoteid", JoinOperator.Inner);
            quoteLink.Columns = new ColumnSet("name", "totalamount");
            quoteLink.EntityAlias = "quote";

            query.LinkEntities.Add(quoteLink);

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity order in collection.Entities)
            {
                Console.WriteLine("Order Name: " + order.GetAttributeValue<string>("name"));
                Console.WriteLine("Order Total Amount: " + order.GetAttributeValue<Money>("totalamount").Value);
                Console.WriteLine("----------------------------------------------");

                Console.WriteLine("Quote Name: " + order.GetAttributeValue<AliasedValue>("quote.name").Value);
                string formattedamt = (((Money)order.GetAttributeValue<AliasedValue>("quote.totalamount").Value).Value).ToString("0.00");
                Console.WriteLine("Quote Total Amount: " + formattedamt);
                Console.WriteLine("----------------------------------------------");
            }
        }
    }
}
