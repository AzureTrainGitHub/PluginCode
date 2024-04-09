using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace KSA.CRM.Plugins
{
    public class CommonDBOperations
    {
        /// <summary>
        /// Assign User
        /// </summary>
        /// <param name="userID"></param>
        /// <param name="targetEntityID"></param>
        /// <param name="targetEntityName"></param>
        /// <param name="organizationService"></param>
        public void AssignUser(Guid userID, Guid targetEntityID, string targetEntityName, IOrganizationService organizationService)
        {
            EntityReference newuserId = new EntityReference() { LogicalName = "systemuser", Id = userID };
            AssignRequest req = new AssignRequest()
            {
                Assignee = newuserId,
                Target = new EntityReference(targetEntityName, targetEntityID)
            };
            organizationService.Execute(req);
        }

        /// <summary>
        /// Create Task On Assignment
        /// </summary>
        /// <param name="userID"></param>
        /// <param name="targetEntityID"></param>
        /// <param name="targetEntityName"></param>
        /// <param name="subject"></param>
        /// <param name="category"></param>
        /// <param name="status"></param>
        /// <param name="organizationService"></param>
        public void CreateTask(Guid userID, Guid targetEntityID, string targetEntityName, string subject, int category, int status, IOrganizationService organizationService)
        {
            Entity task = new Entity(Constants.TaskAttributes.entityname);
            task[Constants.TaskAttributes.Subject] = subject;
            task[Constants.TaskAttributes.regardingobjectid] = new EntityReference(targetEntityName, targetEntityID);
            task[Constants.TaskAttributes.ActivityCategory] = new OptionSetValue(category);
            task[Constants.TaskAttributes.ActivityStatus] = new OptionSetValue(status);
            task[Constants.TaskAttributes.ActualStartDate] = DateTime.Now;
            task[Constants.TaskAttributes.ownerid] = new EntityReference(Constants.Arguments.SystemUser, userID);

            Guid taskid = organizationService.Create(task);
        }

        /// <summary>
        /// Complete Task On Approval/Rejection
        /// </summary>
        /// <param name="Status"></param>
        /// <param name="State"></param>
        /// <param name="task"></param>
        /// <param name="organizationService"></param>
        public void ClompleteTask(int Status, int State, Entity task, IOrganizationService organizationService)
        {
            SetStateRequest setStateRequest = new SetStateRequest();
            setStateRequest.EntityMoniker = new EntityReference(task.LogicalName, task.Id);
            setStateRequest.State = new OptionSetValue(State);
            setStateRequest.Status = new OptionSetValue(Status);
            SetStateResponse setStateResponse = (SetStateResponse)organizationService.Execute(setStateRequest);
        }

        /// <summary>
        /// Get Associated Task records to a record
        /// </summary>
        /// <param name="entityId"></param>
        /// <param name="organizationService"></param>
        public void GetTaskForRecord(Guid entityId, IOrganizationService organizationService)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='task'>");
            strFetchXml.Append("<attribute name='activityid' />");
            strFetchXml.Append("<attribute name='subject' />");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='regardingobjectid' operator='eq'  value='" + entityId + "' />");
            strFetchXml.Append("<condition attribute='statuscode' value='5' operator='ne'/>");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = organizationService.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));

                if (resultXml != null && resultXml.Entities.Count > 0)
                {
                    foreach (Entity tasks in resultXml.Entities)
                    {
                        Entity task = new Entity()
                        {
                            LogicalName = "task",
                            Id = tasks.Id,
                        };

                        ClompleteTask(5, 1, task, organizationService);
                    }
                }
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// Get List of users depending on security role and LOB
        /// </summary>
        /// <param name="securityRole"></param>
        /// <param name="lobType"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public EntityCollection RetrieveUser(string securityRole, string lobType, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='systemuser'>");
            strFetchXml.Append("<attribute name='systemuserid'/>");
            strFetchXml.Append("<attribute name='fullname'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_lineofbusiness' operator='eq' value='" + lobType + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("<link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>");
            strFetchXml.Append("<link-entity name='role' from='roleid' to='roleid' alias='ab'>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='name' operator='eq' value='" + securityRole + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }

        /// <summary>
        /// Get List of users depending on security role and LOB
        /// </summary>
        /// <param name="securityRole"></param>
        /// <param name="lobType"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public EntityCollection RetrieveUser(string securityRole, Guid lobType, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='systemuser'>");
            strFetchXml.Append("<attribute name='systemuserid'/>");
            strFetchXml.Append("<attribute name='fullname'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_lineofbusiness' operator='eq' value='" + lobType + "' />");
            strFetchXml.Append("<condition attribute='rsa_departmentteam' operator='eq' value='1'/>");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("<link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>");
            strFetchXml.Append("<link-entity name='role' from='roleid' to='roleid' alias='ab'>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='name' operator='eq' value='" + securityRole + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }

        /// <summary>
        /// Get List of users depending on security role
        /// </summary>
        /// <param name="securityRole"></param>
        /// <param name="lobType"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public EntityCollection RetrieveUser(string securityRole, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='systemuser'>");
            strFetchXml.Append("<attribute name='systemuserid'/>");
            strFetchXml.Append("<attribute name='fullname'/>");
            strFetchXml.Append("<link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>");
            strFetchXml.Append("<link-entity name='role' from='roleid' to='roleid' alias='ab'>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='name' operator='eq' value='" + securityRole + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }

        /// <summary>
        /// Get List of users depending on Security Role and Branch
        /// </summary>
        /// <param name="securityRole"></param>
        /// <param name="branch"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public EntityCollection RetrieveUser(string securityRole, int branch, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='systemuser'>");
            strFetchXml.Append("<attribute name='systemuserid'/>");
            strFetchXml.Append("<attribute name='fullname'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_branch' operator='eq' value='" + branch.ToString() + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("<link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>");
            strFetchXml.Append("<link-entity name='role' from='roleid' to='roleid' alias='ab'>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='name' operator='eq' value='" + securityRole + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</link-entity>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }

        /// <summary>
        /// Get Records of a particular type assigned to an user
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="targetEntityName"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public EntityCollection RetrieveRecordForUser(Guid userid, string targetEntityName, string targetEntityGuid, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            strFetchXml.Append("<entity name='" + targetEntityName + "'>");
            strFetchXml.Append("<attribute name='" + targetEntityGuid + "'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='ownerid' operator='eq' value='" + Convert.ToString(userid) + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }

        /// <summary>
        /// Send Email With Template
        /// </summary>
        /// <param name="service"></param>
        /// <param name="to"></param>
        /// <param name="entityTypeTo"></param>
        /// <param name="from"></param>
        /// <param name="entityTypeFrom"></param>
        /// <param name="regarding"></param>
        /// <param name="entityTypeRegarding"></param>
        /// <param name="template"></param>
        public bool SendMailFromTemplate(IOrganizationService service, string templateName, Guid to, string sTo, string entityTypeTo, string CCEmailIds, Guid KAMId, Guid regarding, string entityTypeRegarding, int status, int state, string actName, string actDesc, int category, ITracingService tracingservice, IOrganizationServiceFactory servicefactory)
        {

            bool isemailsent = false;

            QueryByAttribute configentity = new QueryByAttribute(Constants.ConfigurationEntity.entityname);
            configentity.ColumnSet = new ColumnSet { AllColumns = true };
            configentity.AddAttributeValue(Constants.ConfigurationEntity.name, Constants.ConfigurationEntity.serviceAccount);
            EntityCollection ServiceAccount = service.RetrieveMultiple(configentity);

            string fromuser = string.Empty;
            foreach (var serv in ServiceAccount.Entities)
            {
                fromuser = serv[Constants.ConfigurationEntity.value] as string;
            }

            QueryByAttribute fromusers = new QueryByAttribute(Constants.Arguments.SystemUser);
            fromusers.ColumnSet = new ColumnSet(Constants.SystemUser.fullname);
            fromusers.AddAttributeValue(Constants.SystemUser.fullname, fromuser);
            EntityCollection emailfromuser = service.RetrieveMultiple(fromusers);
            Guid emailfromuserid = new Guid();
            foreach (var efrom in emailfromuser.Entities)
            {
                emailfromuserid = efrom.Id;

            }

            EntityReference emailfromusers = new EntityReference(Constants.Arguments.SystemUser, emailfromuserid);
            Entity fromparty = new Entity(Constants.Activity.activityparty);
            fromparty.Attributes.Add(Constants.Activity.partyid, emailfromusers);

            Entity toparty = null;
            if (to != Guid.Empty)
            {
                EntityReference emailtousers = new EntityReference(entityTypeTo, to);
                toparty = new Entity(Constants.Activity.activityparty);
                toparty.Attributes.Add(Constants.Activity.partyid, emailtousers);
            }
            else if (sTo != string.Empty)
            {
                toparty = new Entity(Constants.Activity.activityparty);
                toparty.Attributes.Add(Constants.Activity.ccList, sTo);
            }

            EntityCollection frompartycollection = new EntityCollection();
            frompartycollection.EntityName = Constants.Arguments.SystemUser;
            frompartycollection.Entities.Add(fromparty);

            EntityCollection topartycollection = new EntityCollection();
            topartycollection.EntityName = Constants.Arguments.SystemUser;
            topartycollection.Entities.Add(toparty);

            string CCEmailList = CCEmailIds;
            string[] cclist = CCEmailList.Split(';');
            EntityCollection ccListcollection = new EntityCollection();
            foreach (var cc in cclist)
            {
                Entity rsa_EmailCCReciepent = new Entity(Constants.Activity.activityparty);

                rsa_EmailCCReciepent[Constants.Activity.ccList] = cc;
                ccListcollection.Entities.Add(rsa_EmailCCReciepent);
            }
            if (KAMId != Guid.Empty)
            {
                EntityReference KAMUser = new EntityReference(Constants.Arguments.SystemUser, KAMId);
                Entity kamCC = new Entity(Constants.Activity.activityparty);
                kamCC.Attributes.Add(Constants.Activity.partyid, KAMUser);
                ccListcollection.Entities.Add(kamCC);
            }


            Entity email = new Entity(Constants.EmailEntityAttributes.entityname);

            email[Constants.EmailEntityAttributes.ownerid] = new EntityReference(Constants.Arguments.SystemUser, emailfromuserid);
            email[Constants.EmailEntityAttributes.regardingobjectid] = new EntityReference(entityTypeRegarding, regarding);

            if (topartycollection.Entities.Count > 0)
                email[Constants.EmailEntityAttributes.to] = topartycollection;
            else
                email[Constants.EmailEntityAttributes.to] = ccListcollection;

            email[Constants.EmailEntityAttributes.from] = frompartycollection;
            email[Constants.EmailEntityAttributes.cc] = ccListcollection;

            email[Constants.EmailEntityAttributes.category] = new OptionSetValue(category);
            email[Constants.EmailEntityAttributes.state] = new OptionSetValue(state);
            email[Constants.EmailEntityAttributes.status] = new OptionSetValue(status);
            email[Constants.EmailEntityAttributes.type] = new OptionSetValue(1);
            email[Constants.EmailEntityAttributes.name] = actName;
            email[Constants.EmailEntityAttributes.cdescription] = actDesc;

            Entity GlobalEmailTemplate = GetEmailTemplate(templateName, service, tracingservice);
            email[Constants.EmailEntityAttributes.subject] = GetDataFromXML(GlobalEmailTemplate.Attributes[Constants.EmailTemplate.subject].ToString(), Constants.EmailTemplate.match);
            email[Constants.EmailEntityAttributes.description] = GetDataFromXML(GlobalEmailTemplate.Attributes[Constants.EmailTemplate.body].ToString(), Constants.EmailTemplate.match);
            email.Attributes[Constants.EmailEntityAttributes.description] = email.Attributes[Constants.EmailEntityAttributes.description].ToString();

            Guid emailid = service.Create(email);
            if (emailid != null)
            {
                IOrganizationService sraservice = servicefactory.CreateOrganizationService(emailfromuserid);
                SendEmailRequest req = new SendEmailRequest();
                req.EmailId = emailid;
                req.TrackingToken = "";
                req.IssueSend = true;
                SendEmailResponse res = (SendEmailResponse)sraservice.Execute(req);
                if (res.Results.Count > 0)
                {
                    isemailsent = true;
                }
                else
                {
                    isemailsent = false;
                }
            }
            return isemailsent;
        }

        /// <summary>
        /// Create Decline Email
        /// </summary>
        /// <param name="to"></param>
        /// <param name="CCEmailIds"></param>
        /// <param name="service"></param>
        /// <param name="regarding"></param>
        /// <param name="entityTypeRegarding"></param>
        /// <param name="tracingservice"></param>
        /// <returns></returns>
        public bool CreateDeclineEmail(string templateName, Guid to, string sTo, string entityTypeTo, string CCEmailIds, Guid KAMId, IOrganizationService service, Guid regarding, string entityTypeRegarding, int status, int state, string actName, string actDesc, int category, ITracingService tracingservice, string subject, string enquiryId, Guid unPM)
        {
            Guid emailid = Guid.Empty;
            bool isemailcreated = false;
            string rsaLogo = string.Empty;
            QueryByAttribute configentity = new QueryByAttribute(Constants.ConfigurationEntity.entityname);
            configentity.ColumnSet = new ColumnSet { AllColumns = true };
            configentity.AddAttributeValue(Constants.ConfigurationEntity.name, Constants.ConfigurationEntity.serviceAccount);
            EntityCollection ServiceAccount = service.RetrieveMultiple(configentity);

            string fromuser = string.Empty;
            foreach (var serv in ServiceAccount.Entities)
            {
                fromuser = serv[Constants.ConfigurationEntity.value] as string;
            }

            QueryByAttribute fromusers = new QueryByAttribute(Constants.Arguments.SystemUser);
            fromusers.ColumnSet = new ColumnSet(Constants.SystemUser.fullname);
            fromusers.AddAttributeValue(Constants.SystemUser.fullname, fromuser);
            EntityCollection emailfromuser = service.RetrieveMultiple(fromusers);
            Guid emailfromuserid = new Guid();
            foreach (var efrom in emailfromuser.Entities)
            {
                emailfromuserid = efrom.Id;

            }

            EntityReference emailfromusers = new EntityReference(Constants.Arguments.SystemUser, emailfromuserid);
            Entity fromparty = new Entity(Constants.Activity.activityparty);
            fromparty.Attributes.Add(Constants.Activity.partyid, emailfromusers);

            Entity toparty = null;
            if (to != Guid.Empty)
            {
                EntityReference emailtousers = new EntityReference(entityTypeTo, to);
                toparty = new Entity(Constants.Activity.activityparty);
                toparty.Attributes.Add(Constants.Activity.partyid, emailtousers);
            }
            else if (sTo != string.Empty)
            {
                toparty = new Entity(Constants.Activity.activityparty);
                toparty.Attributes.Add(Constants.Activity.ccList, sTo);
            }

            EntityCollection frompartycollection = new EntityCollection();
            frompartycollection.EntityName = Constants.Arguments.SystemUser;
            frompartycollection.Entities.Add(fromparty);

            EntityCollection topartycollection = new EntityCollection();
            topartycollection.Entities.Add(toparty);

            EntityCollection ccListcollection = new EntityCollection();
            if (!string.IsNullOrEmpty(CCEmailIds))
            {
                //string CCEmailList = CCEmailIds;
                string[] cclist = CCEmailIds.Split(';');

                foreach (var cc in cclist)
                {
                    Entity rsa_EmailCCReciepent = new Entity(Constants.Activity.activityparty);

                    rsa_EmailCCReciepent[Constants.Activity.ccList] = cc;
                    ccListcollection.Entities.Add(rsa_EmailCCReciepent);
                }
            }
            if (KAMId != Guid.Empty)
            {
                EntityReference KAMUser = new EntityReference(Constants.Arguments.SystemUser, KAMId);
                Entity kamCC = new Entity(Constants.Activity.activityparty);
                kamCC.Attributes.Add(Constants.Activity.partyid, KAMUser);
                ccListcollection.Entities.Add(kamCC);
            }


            Entity email = new Entity(Constants.EmailEntityAttributes.entityname);

            if (unPM != Guid.Empty)
            {
                email[Constants.EmailEntityAttributes.ownerid] = new EntityReference(Constants.Arguments.SystemUser, unPM);
            }
            else
            {
                email[Constants.EmailEntityAttributes.ownerid] = new EntityReference(Constants.Arguments.SystemUser, emailfromuserid);
            }
            email[Constants.EmailEntityAttributes.regardingobjectid] = new EntityReference(entityTypeRegarding, regarding);

            if (topartycollection.Entities.Count > 0)
            {
                email[Constants.EmailEntityAttributes.to] = topartycollection;
            }
            else
                email[Constants.EmailEntityAttributes.to] = ccListcollection;

            email[Constants.EmailEntityAttributes.from] = frompartycollection;
            email[Constants.EmailEntityAttributes.cc] = ccListcollection;

            email[Constants.EmailEntityAttributes.category] = new OptionSetValue(category);
            email[Constants.EmailEntityAttributes.state] = new OptionSetValue(state);
            email[Constants.EmailEntityAttributes.status] = new OptionSetValue(status);
            email[Constants.EmailEntityAttributes.type] = new OptionSetValue(1);
            email[Constants.EmailEntityAttributes.name] = actName;
            email[Constants.EmailEntityAttributes.cdescription] = actDesc;

            Entity GlobalEmailTemplate = GetEmailTemplate(templateName, service, tracingservice);

            if (enquiryId != string.Empty)
            {
                if (subject != string.Empty)
                {
                    email[Constants.EmailEntityAttributes.subject] = "RE: " + subject + "-" + enquiryId;
                }
                else
                {
                    email[Constants.EmailEntityAttributes.subject] = GetDataFromXML(GlobalEmailTemplate.Attributes[Constants.EmailTemplate.subject].ToString(), Constants.EmailTemplate.match);
                }
            }
            else
            {
                if (subject != string.Empty)
                {
                    email[Constants.EmailEntityAttributes.subject] = "RE: " + subject;
                }
                else
                {
                    email[Constants.EmailEntityAttributes.subject] = GetDataFromXML(GlobalEmailTemplate.Attributes[Constants.EmailTemplate.subject].ToString(), Constants.EmailTemplate.match);
                }
            }

            email[Constants.EmailEntityAttributes.description] = GetDataFromXML(GlobalEmailTemplate.Attributes[Constants.EmailTemplate.body].ToString(), Constants.EmailTemplate.match);

            #region get server url
            QueryExpression serUrlQuery = new QueryExpression("rsa_configurationentity");
            //serUrlQuery.Criteria.AddCondition("rsa_name", ConditionOperator.Equal, "Server Url");
            serUrlQuery.Criteria.AddCondition("rsa_name", ConditionOperator.Equal, "RSALogo");
            serUrlQuery.ColumnSet.AllColumns = true;

            Entity serUrlRec = service.RetrieveMultiple(serUrlQuery).Entities.FirstOrDefault();
            if (serUrlRec.Contains("rsa_value"))
            {
                //rsaLogo = serUrlRec["rsa_value"].ToString() + "/KSA/WebResources/rsa_/scripts/images/rsalogo.jpg";
                //rsaLogo = "http://ctsc00192164301/RSAWeb/Images/rsalogo.jpg";
                rsaLogo = serUrlRec["rsa_value"].ToString();
            }
            #endregion

            email.Attributes[Constants.EmailEntityAttributes.description] = email.Attributes[Constants.EmailEntityAttributes.description].ToString().Replace("RSALOGO", "<img src='" + rsaLogo + "'/>");

            emailid = service.Create(email);
            if (emailid != null)
            {
                isemailcreated = true;
            }
            else
            {
                isemailcreated = false;
            }
            return isemailcreated;
        }

        /// <summary>
        /// Get Data From Template
        /// </summary>
        /// <param name="value"></param>
        /// <param name="attributename"></param>
        /// <returns></returns>
        private object GetDataFromXML(string value, string attributename)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            XDocument document = XDocument.Parse(value);
            XElement element = document.Descendants().Where(ele => ele.Attributes().Any(attr => attr.Name == attributename)).FirstOrDefault();
            return element == null ? string.Empty : element.Value;
        }

        /// <summary>
        /// Get Email Template Details For Email
        /// </summary>
        /// <param name="title"></param>
        /// <param name="service"></param>
        /// <param name="tracingservice"></param>
        /// <returns></returns>
        private Entity GetEmailTemplate(string title, IOrganizationService service, ITracingService tracingservice)
        {
            Entity emailtemplate = null;
            string queryGlobalEmailTemplate = string.Format(Constants.FetchXmls.queryEmailTemplate, title);
            EntityCollection emailtemplates = service.RetrieveMultiple(new FetchExpression(queryGlobalEmailTemplate));
            if (emailtemplates.Entities.Count == 1)
            {
                return emailtemplates.Entities[0];
            }
            else
            {
                return emailtemplate;
            }
        }

        /// <summary>
        /// Set Auto Number
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="service"></param>
        public Entity SetAutoNumber(string entityName, IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            Entity autoNumber = new Entity("rsa_autonumber");
            StringBuilder strFetchXml = new StringBuilder();
            Dictionary<string, Int32> autoNum = new Dictionary<string, int>();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>");
            strFetchXml.Append("<entity name='rsa_autonumber'>");
            strFetchXml.Append("<attribute name='rsa_autonumberid'/>");
            strFetchXml.Append("<attribute name='rsa_format'/>");
            strFetchXml.Append("<attribute name='rsa_number'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_name' operator='eq' value='" + entityName + "' />");
            strFetchXml.Append("</filter>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));

                if (resultXml != null && resultXml.Entities.Count > 0)
                {
                    var num = (from opp in resultXml.Entities
                               select new
                               {
                                   Number = new
                                   {
                                       Format = opp["rsa_format"],
                                       Number = opp["rsa_number"],
                                       ID = opp["rsa_autonumberid"]
                                   }
                               }).FirstOrDefault();


                    if (num != null)
                    {
                        autoNumber["rsa_format"] = Convert.ToString(num.Number.Format);
                        autoNumber["rsa_number"] = Convert.ToInt32(num.Number.Number);
                        autoNumber["rsa_autonumberid"] = new Guid(Convert.ToString(num.Number.ID));
                    }
                }
            }
            catch (Exception ex)
            { }

            return autoNumber;
        }


        public EntityCollection RetrieveSLARecord(Guid entityId, SLAcalculator.Stage entityName,IOrganizationService service)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            strFetchXml.Append("<entity name='rsa_sla'>");
            strFetchXml.Append("<attribute name='rsa_slaid'/>");
            strFetchXml.Append("<attribute name='rsa_status'/>");
            strFetchXml.Append("<attribute name='rsa_starttime'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_starttime' operator='not-null' />");
            strFetchXml.Append("<condition attribute='rsa_endtime' operator='null' />");

            if (Convert.ToInt32(entityName) == 5)
                strFetchXml.Append("<condition attribute='rsa_activityid' operator='eq'  value='" + entityId + "' />");
            else if (Convert.ToInt32(entityName) == 1)
                strFetchXml.Append("<condition attribute='rsa_leadid' operator='eq'  value='" + entityId + "' />");
            else if (Convert.ToInt32(entityName) == 2)
                strFetchXml.Append("<condition attribute='rsa_opportunityid' operator='eq'  value='" + entityId + "' />");
            else if (Convert.ToInt32(entityName) == 3)
                strFetchXml.Append("<condition attribute='rsa_quoteid' operator='eq'  value='" + entityId + "' />");
            else if (Convert.ToInt32(entityName) == 4)
                strFetchXml.Append("<condition attribute='rsa_policyid' operator='eq'  value='" + entityId + "' />");

            strFetchXml.Append("</filter>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = service.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }


        public EntityCollection ReadSLAConfig(Guid LOBCode, bool isActLead, SLAcalculator.Stage stage, SLAcalculator.Status status, IOrganizationService orgService)
        {
            EntityCollection resultXml = new EntityCollection();
            string result = string.Empty;
            StringBuilder strFetchXml = new StringBuilder();

            strFetchXml.AppendFormat(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            strFetchXml.Append("<entity name='rsa_slaconfig'>");
            strFetchXml.Append("<attribute name='rsa_slaconfigid'/>");
            strFetchXml.Append("<attribute name='rsa_slaviolatedred'/>");
            strFetchXml.Append("<attribute name='rsa_slametstatusgreen'/>");
            strFetchXml.Append("<attribute name='rsa_slametstatusamber'/>");
            strFetchXml.Append("<filter type='and'>");
            strFetchXml.Append("<condition attribute='rsa_stage' operator='" + Convert.ToString(stage) + "' />");
            strFetchXml.Append("<condition attribute='rsa_status' operator='" + Convert.ToString(status) + "' />");

            if (isActLead == false)
                strFetchXml.Append("<condition attribute='rsa_lob' operator='eq' value='" + LOBCode + "' />");

            strFetchXml.Append("</filter>");
            strFetchXml.Append("</entity>");
            strFetchXml.Append("</fetch>");

            try
            {
                resultXml = orgService.RetrieveMultiple(new FetchExpression(strFetchXml.ToString()));
            }
            catch (Exception ex)
            { }

            return resultXml;
        }
    }

    
}
