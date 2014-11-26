using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;

namespace GSS_Alert_Service
{
    public partial class Alert_Service : ServiceBase
    {
        private Stopwatch runTimer;
        private System.Timers.Timer heartbeat;
        private System.Timers.Timer CommitTimer;
        private System.Timers.Timer AssistTimer;
        private System.Timers.Timer MailSystemTimer;
        public static bool DebugEnabled = false;
        private int HeartbeatCount = 0;
        private int DebugCommitCycles = 0, DebugAssistCycles = 0, DebugMailCycles = 0;
        private int HeartbeatCasesFound = 0, TotalCasesFound = 0;
        private int HeartbeatAssistsFound = 0, TotalAssistsFound = 0;
        private int HeartbeatEmailSent = 0, TotalEmailsSent = 0;

        public Alert_Service()
        {
            // Setup taskbar icon and event logger
            InitializeComponent();

            // Setup event logging
            this.AutoLog = false;
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists("GSS Alerts Service"))
            {
                EventLog.CreateEventSource("GSS Alerts Service", "GSS Alerts Service");
            }
            eventLog1.Source = "GSS Alerts Service";
            eventLog1.Log = "GSS Alerts Service";

            // Check Registry for debug setting
            try
            {
                DebugEnabled = bool.Parse(RegistryEngine.ReadRegistry("DebugEnabled"));
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(string.Format("Information:/nDebug key not found, continuing at normal logging levels."), EventLogEntryType.Information, 1);
            }
            
            if (DebugEnabled)
            {
                eventLog1.WriteEntry("Warning: Debugging has been enabled for the service.  This will cause a significant amount of log churn!/n/n***Deubg messages logged as EventID 0.", EventLogEntryType.Warning, 1);
            }

            // Setup up time tracker
            runTimer = new Stopwatch();

            // Setup heartbeat monitor
            heartbeat = new System.Timers.Timer();
            if (DebugEnabled)
                heartbeat.Interval = 10000; // Debug Value            
            else
                heartbeat.Interval = 3600000; // Production value
            heartbeat.Elapsed += new System.Timers.ElapsedEventHandler(this.OnHeartbeatTimer);
            if (DebugEnabled)
            {
                eventLog1.WriteEntry("Debug:/nHeartbeat timer setup complete./nHeartbeat interval set to 10s.", EventLogEntryType.Information, 0);
            }

            // Setup commit subsystem
            CommitTimer = new System.Timers.Timer();
            try
            {
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(@"Debug:/nReading HKLM\SOFTWARE\GSSTools\CommitInterval key...", EventLogEntryType.Information, 0);
                }
                CommitTimer.Interval = int.Parse(RegistryEngine.ReadRegistry("CommitInterval"));
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(string.Format("Debug:/nCommit check interval set to {0} ms.", CommitTimer.Interval.ToString()), EventLogEntryType.Information, 0);
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(string.Format("Warning:/nCommit check interval key does not exist!\nKey created with default value of 5 seconds./n/nStack Trace:/n/n{0}", ex.ToString()), EventLogEntryType.Warning, 100);
                try
                {
                    RegistryEngine.WriteRegistry("CommitInterval", "5000");
                }
                catch (Exception e)
                {
                    eventLog1.WriteEntry(string.Format("Error:/nRegistry write failure for key: CommitInterval/n/nStack Trace:/n/n{0}", e.ToString()), EventLogEntryType.Error,500);
                }
            }
            CommitTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnCommitTimer);

            // Setup assist subsystem
            AssistTimer = new System.Timers.Timer();
            try
            {
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(@"Debug:/nReading HKLM\SOFTWARE\GSSTools\AssistInterval key...", EventLogEntryType.Information, 0);
                }
                AssistTimer.Interval = int.Parse(RegistryEngine.ReadRegistry("AssistInterval"));
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(string.Format("Debug:/nAssist check interval set to {0} ms.", AssistTimer.Interval.ToString()), EventLogEntryType.Information, 0);
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(string.Format("Warning: Assist check interval key does not exist!\nKey created with default value of 1 minute./n/nStack Trace:/n/n{0}", ex.ToString()), EventLogEntryType.Warning, 101);
                try
                {
                    RegistryEngine.WriteRegistry("AssistInterval", "60000");
                }
                catch (Exception e)
                {
                    eventLog1.WriteEntry(string.Format("Error:/nRegistry write failure for key: AssistInterval/n/nStack Trace:/n/n{0}", e.ToString()), EventLogEntryType.Error, 500);
                }
            }
            AssistTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnAssistTimer);

            // Setup mail subsystem
            MailSystemTimer = new System.Timers.Timer();
            try
            {
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(@"Debug:/nReading HKLM\SOFTWARE\GSSTools\MailInterval key...", EventLogEntryType.Information, 0);
                }
                MailSystemTimer.Interval = int.Parse(RegistryEngine.ReadRegistry("MailInterval"));
                if (DebugEnabled)
                {
                    eventLog1.WriteEntry(string.Format("Debug:/nSend mail interval set to {0} ms.", MailSystemTimer.Interval.ToString()), EventLogEntryType.Information, 0);
                }
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(string.Format("Warning: Mail interval key does not exist!\nKey created with default value of 2.5 seconds./n/nStack Trace:/n/n{0}", ex.ToString()), EventLogEntryType.Warning, 102);
                try
                {
                    RegistryEngine.WriteRegistry("MailInterval", "2500");
                }
                catch (Exception e)
                {
                    eventLog1.WriteEntry(string.Format("Error:/nRegistry write failure for key: MailInterval/n/nStack Trace:/n/n{0}", e.ToString()), EventLogEntryType.Error, 500);
                }
            }
            MailSystemTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnMailSystemTimer);
        }

        protected override void OnStart(string[] args)
        {
            // Update status to Start Pending
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nService start pending set."), EventLogEntryType.Information, 0);
            }

            // Start the heartbeat services 
            runTimer.Reset();
            runTimer.Start();
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nRun timer stopwatch running: {0}", runTimer.IsRunning), EventLogEntryType.Information, 0);
            }
            heartbeat.Start();
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nHeartbeat service enabled: {0}", heartbeat.Enabled), EventLogEntryType.Information, 0);
            }

            // Start the commit subsystem
            CommitTimer.Start();
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nCommit subsystem enabled: {0}", CommitTimer.Enabled), EventLogEntryType.Information, 0);
            }

            // Start the assist subsystem
            AssistTimer.Start();
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nAssist subsystem enabled {0}", AssistTimer.Enabled), EventLogEntryType.Information, 0);
            }

            // Start the mail subsystem
            MailSystemTimer.Start();
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/n"), EventLogEntryType.Information, 0);
            }

            // Complete!
            eventLog1.WriteEntry("GSS Alert Service started successfully.", EventLogEntryType.Information, 5);

            // Start up complete, set status to Running
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nSetting service to stop pending."), EventLogEntryType.Information, 0);
            }
            // Update status to Stop Pending
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Stop all timers
            runTimer.Stop();
            heartbeat.Stop();
            CommitTimer.Stop();
            AssistTimer.Stop();
            MailSystemTimer.Stop();
            GUI.Dispose();
            eventLog1.WriteEntry(string.Format("GSS Alert Service stopped successfully after running for {0} minutes./n{1} cases and {2} assists were found./n{3} emails were sent.", runTimer.Elapsed.TotalMinutes, TotalCasesFound, TotalAssistsFound, TotalEmailsSent), EventLogEntryType.Information, 6);

            // Stop complete, set status to Stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
        }

        #region Service Status Code
        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public long dwServiceType;
            public ServiceState dwCurrentState;
            public long dwControlsAccepted;
            public long dwWin32ExitCode;
            public long dwServiceSpecificExitCode;
            public long dwCheckPoint;
            public long dwWaitHint;
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
        #endregion
        #region Timer Events
        private void OnHeartbeatTimer (object sender, System.Timers.ElapsedEventArgs args)
        {
            HeartbeatCount++;
            
            Debug.WriteLine("Heartbeat Timer called!");
            eventLog1.WriteEntry(string.Format("Heartbeat {0}:/nThe service has been online for {1} minutes./n{2} cases found this heartbeat ({3} total)./n{4} assists found this heartbeat ({5} total)./n{6} emails sent this heartbeat ({7} total).",HeartbeatCount , runTimer.Elapsed.TotalMinutes.ToString(), HeartbeatCasesFound, TotalCasesFound, HeartbeatAssistsFound, TotalAssistsFound, HeartbeatEmailSent, TotalEmailsSent), EventLogEntryType.Information, 2);

            HeartbeatCasesFound = 0;
            HeartbeatAssistsFound = 0;
            HeartbeatEmailSent = 0;
        }

        private void OnCommitTimer (object sender, System.Timers.ElapsedEventArgs args)
        {
            if (DebugEnabled)
            {
                DebugCommitCycles++;
                eventLog1.WriteEntry(string.Format("Debug:/nCommit subsystem called./nCycle: {0}/nCases Heartbeat / Total: {1} / {2}", DebugCommitCycles, HeartbeatCasesFound, TotalCasesFound), EventLogEntryType.Information, 0);
            }

                

            Debug.WriteLine("Commit Timer called!");
            // TODO: Add commit logic
        }

        private void OnAssistTimer (object sender, System.Timers.ElapsedEventArgs args)
        {
            if (DebugEnabled)
            {
                DebugAssistCycles++;
                eventLog1.WriteEntry(string.Format("Debug:/nAssist subsystem called./nCycle: {0}/nAssists Heartbeat / Total: {1} / {2}", DebugAssistCycles, HeartbeatAssistsFound, TotalAssistsFound), EventLogEntryType.Information, 0);
            }

            Debug.WriteLine("Assist Timer called!");
            // TODO: Add assist logic
        }

        private void OnMailSystemTimer (object sender, System.Timers.ElapsedEventArgs args)
        {
            if (DebugEnabled)
            {
                DebugMailCycles++;
                eventLog1.WriteEntry(string.Format("Debug:/nMail subsystem called./nCycle: {0}/nMail Sent Heartbeat / Total: {1} / {2}", DebugMailCycles, HeartbeatEmailSent, TotalEmailsSent), EventLogEntryType.Information, 0);
            }

            Debug.WriteLine("Mail Timer called!");
            // TODO: Add mail logic
        }
        #endregion Timer Events

        private void notifyIcon1_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (DebugEnabled)
            {
                eventLog1.WriteEntry(string.Format("Debug:/nGUI handler called!"), EventLogEntryType.Information, 0);
            }
            if (!GUI.Visible)
                GUI.Show();
            else
                GUI.Hide();
        }
    }
}
