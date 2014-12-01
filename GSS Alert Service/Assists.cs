using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using Alert_Service_Classes;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Net.Mail;
using System.Timers;
using GSS_Alert_Service.SDFC;
using System.Web.Services.Protocols;
using System.Collections;

namespace GSS_Alert_Service
{
    class Assists
    {
        AssistAlerterConfig Config = new AssistAlerterConfig();
        string ConfigFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "GSS_Tools\\SFDCAssistAlertConfig.bin");
        SmtpClient MailClient = new SmtpClient("pa-smtp.vmware.com");
        List<MailMessage> AlertsToSend = new List<MailMessage>();
        Timer SFDC_Check_Timer = new Timer(60000);
        Timer Alert_Send_Timer = new Timer(10000);
        string query = "SELECT Name, Id, GSS_Case_Number__r.CaseNumber, CreatedBy.Name, GSS_Requesting_To_Person__c, CreatedDate, GSS_Product_Training_Gap__c, GSS_Description__c, Owner.Name FROM GSS_Request__c WHERE GSS_Status__c = 'Requested' AND (GSS_Case_Center__c IN ({0}) OR GSS_Case_Center__c = '') AND Owner.Name IN ({1})";
        SforceService binding;
        DataTable results = new DataTable();
        List<AssistRequest> Requests = new List<AssistRequest>();
        private EventLog AppLog;

        public Assists()
        {
            AppLog = new EventLog();
            if (!EventLog.SourceExists("GSS Alerts Service - Assist System"))
            {
                EventLog.CreateEventSource("GSS Alerts Service - Assist System", "GSS Alerts Service");
            }
            AppLog.Source = "GSS Alerts Service - Assist System";
            AppLog.Log = "GSS Alerts Service";

            results.Columns.Add("Number");
            results.Columns.Add("ID");
            results.Columns.Add("CaseNumber");
            results.Columns.Add("RequestedBy");
            results.Columns.Add("RequestedTo");
            results.Columns.Add("CreatedDate");
            results.Columns.Add("Product");
            results.Columns.Add("Description");
            results.Columns.Add("Queue");
        }

        protected override void OnStart(string[] args)
        {
            LoadConfig();
            SFDC_Check_Timer.Elapsed += new ElapsedEventHandler(SFDC_Check_Timer_Elapsed);
            Alert_Send_Timer.Elapsed += new ElapsedEventHandler(Alert_Send_Timer_Elapsed);
            SFDC_Check_Timer.Start();
            Alert_Send_Timer.Start();
        }

        void Alert_Send_Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            MailMessage[] Queue = AlertsToSend.ToArray();
            foreach (MailMessage Alert in Queue)
            {
                try
                {
                    MailClient.Send(Alert);
                }
                catch (Exception ex)
                {
                    AppLog.WriteEntry(string.Format("Sending Alert failed! To: {0} | Subject: {1} | CC: {2} | Exception: {3}", Alert.To.ToString(), Alert.Subject, Alert.CC.ToString(), ex.Message), EventLogEntryType.Error, 2501);
                }
                finally
                {
                    AlertsToSend.Remove(Alert);
                }
            }
        }

        void SFDC_Check_Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            foreach  (AlertConfig Conf in Config.Alerts)
            {
                if ((DateTime.Now - Conf.LastSent).TotalMinutes >= Conf.Frequency)
                {
                    if ((DateTime.Now.DayOfWeek != DayOfWeek.Saturday) && (DateTime.Now.DayOfWeek != DayOfWeek.Sunday) && (DateTime.Now.Hour >= 7) && (DateTime.Now.Hour <= 18))
                    {
                    Requests.Clear();
                    GetRequests(Conf, ref Requests);
                    CheckAlerts(Conf, Requests);
                    }
                }
            }
        }

        private void CheckAlerts(AlertConfig Conf, List<AssistRequest> Requests)
        {
                Hashtable queues = new Hashtable();
                MailMessage thisMail = new MailMessage(Conf.EmailFrom,Conf.EmailTo);
                thisMail.IsBodyHtml = true;
                int i = 0;
                foreach (AssistRequest thisRequest in Requests)
                {
                    TimeSpan x = DateTime.Now - thisRequest.CreatedDate;
                    if (x.TotalDays <= Conf.PastDays || Conf.PastDays == 0)
                    {
                        thisMail.Body = thisMail.Body + string.Format("<b>Number:</b> <a href=\"https://na6.salesforce.com/{7}\">{0}</a><br><b>Queue:</b> {7}<br><b>Case:</b> {1}<br><b>Date Submitted:</b> {2}<br><b>Submitted By:</b> {3} <br><b>Requesting To:</b> {4}<br><b>Product:</b> {5}<br><b>Description:</b><br>{6}<br><hr>", thisRequest.Number, thisRequest.CaseNumber.Substring(4), thisRequest.CreatedDate.ToShortDateString(), thisRequest.RequestedBy.Substring(4), thisRequest.RequestedTo, thisRequest.Product, thisRequest.Description.Replace("\n", "<br>"),thisRequest.Queue);
                        i++;
                        if (!queues.ContainsKey(thisRequest.Queue))
                        {
                            queues.Add(thisRequest.Queue, 1);
                        }
                        else
                        {
                            queues[thisRequest.Queue] = (int)queues[thisRequest.Queue] + 1;
                        }
                    }
                }
                if (i > 0)
                {
                    string qcount = "";
                    foreach (string queue in queues.Keys)
                    {
                        qcount += string.Format("{0}: {1}", queue, queues[queue].ToString()) + "<br>";
                    }
                    thisMail.Subject = string.Format("{0} assists in queue",i.ToString());
                    thisMail.Body = "Queue Summary:<br>" + qcount + "===============================<br>" + thisMail.Body;
                    AlertsToSend.Add(thisMail);
                    Conf.LastSent = DateTime.Now;
                }
        }

        private void GetRequests(AlertConfig Conf, ref List<AssistRequest> Requests)
        {
            results.Rows.Clear();
            AssistRequest thisRequest = new AssistRequest();

            string Queues = "";
            for (int i = 0; i < Conf.QueueNames.Count(); i++)
			{
                Queues = Queues + "'" + Conf.QueueNames[i] + "'";
                if (i != Conf.QueueNames.Count() - 1)
                {
                    Queues = Queues + ",";
                }
            }

            string Centers = "";
            for (int i = 0; i < Conf.Centers.Count(); i++)
			{
                Centers = Centers + "'" + Conf.Centers[i] + "'";
                if (i != Conf.Centers.Count() - 1)
                {
                    Centers = Centers + ",";
                }
            }

            string thisQuery = string.Format(query, Centers, Queues);

            SFDC_Query(ref results, thisQuery);
            foreach (DataRow thisRow in results.Rows)
            {
                thisRequest = new AssistRequest();
                thisRequest.CaseNumber = thisRow["CaseNumber"].ToString();
                thisRequest.Number = thisRow["Number"].ToString();
                thisRequest.Description = thisRow["Description"].ToString();
                thisRequest.Product = thisRow["Product"].ToString();
                thisRequest.RequestedTo = thisRow["RequestedTo"].ToString();
                thisRequest.RequestedBy = thisRow["RequestedBy"].ToString();
                thisRequest.CreatedDate = DateTime.Parse(thisRow["CreatedDate"].ToString());
                thisRequest.Id = thisRow["Id"].ToString();
                thisRequest.Queue = thisRow["Queue"].ToString().Substring(4);
                Requests.Add(thisRequest);
            }
            if (Requests.Count > 0)
            {
                AppLog.WriteEntry(string.Format("{0} Requests found.", Requests.Count.ToString()),EventLogEntryType.Information);
            }
        }

        protected override void OnStop()
        {
            SFDC_Check_Timer.Stop();
            Alert_Send_Timer.Stop();
        }

        private void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                Stream ConfigStream = File.OpenRead(ConfigFile);
                BinaryFormatter deserializer = new BinaryFormatter();
                Config = (AssistAlerterConfig)deserializer.Deserialize(ConfigStream);
                ConfigStream.Close();
            }
        }

        #region SFDC Interaction
        void SFDC_Query(ref DataTable result, string query)
        {
            if (Login())
            {
                QueryResult qr = null;
                DataRow row = result.NewRow();
                binding.QueryOptionsValue = new QueryOptions();
                binding.QueryOptionsValue.batchSize = 250;
                binding.QueryOptionsValue.batchSizeSpecified = true;

                try
                {
                    qr = binding.query(query);
                    bool done = false;
                    if (qr.size > 0)
                    {
                        Debug.WriteLine("User sees " + qr.records.Length.ToString() + " records." + System.Environment.NewLine);
                        while (!done)
                        {
                            for (int i = 0; i < qr.records.Length; i++)
                            {
                                row = results.NewRow();
                                //for (int j = 0; j < qr.records[i].Any.Length; j++)
                                //{
                                //    Debug.WriteLine(qr.records[i].Any[j].InnerText + "     ");
                                //    row[j] = qr.records[i].Any[j].InnerText;
                                //}
                                //Debug.WriteLine(System.Environment.NewLine);
                                //if (qr.done)
                                //{
                                //    done = true;
                                //}
                                //else
                                //{
                                //    qr = binding.queryMore(qr.queryLocator);
                                //    Debug.WriteLine("Ooh! Getting more!" + System.Environment.NewLine);
                                //}

                                GSS_Request__c request = qr.records[i] as GSS_Request__c;
                                for (int j = 0; j < 9; j++)
                                {
                                    row["Number"] = request.Name;
                                    row["ID"] = request.Id;
                                    row["CaseNumber"] = request.GSS_Case_Number__r.CaseNumber;
                                    row["RequestedBy"] = request.CreatedBy.Name;
                                    row["RequestedTo"] = request.GSS_Requesting_To_Person__c;
                                    row["CreatedDate"] = request.CreatedDate;
                                    row["Product"] = request.GSS_Product_Training_Gap__c;
                                    row["Description"] = request.GSS_Description__c;
                                    row["Queue"] = request.Owner.Name1;
                                }

                                if (qr.done)
                                    done = true;
                                else
                                {
                                    qr = binding.queryMore(qr.queryLocator);
                                }

                                results.Rows.Add(row);
                            }
                        }
                    }
                }
                catch (SoapException ex)
                {
                    AppLog.WriteEntry(string.Format("Login Failed! Error: {0}", ex.Message));
                    Debug.WriteLine(string.Format("Login Failed! Error: {0}", ex.Message));
                }
            }
        }

        private bool Login()
        {
            binding = new SforceService();
            binding.Timeout = 60000;
            LoginResult lr;
            string passtoken = Config.Password + Config.Token;
            try
            {
                lr = binding.login(Config.Username, passtoken);
                Debug.WriteLine("User Logged In! Session ID = " + lr.sessionId.ToString() + System.Environment.NewLine);
                binding.Url = lr.serverUrl;
                binding.SessionHeaderValue = new SessionHeader();
                binding.SessionHeaderValue.sessionId = lr.sessionId;
                return true;
            }
            catch (SoapException ex)
            {
                AppLog.WriteEntry("Login Failed: " + ex.Message);
                return false;
            }
        }
        #endregion SFDC Interaction
    }
}
