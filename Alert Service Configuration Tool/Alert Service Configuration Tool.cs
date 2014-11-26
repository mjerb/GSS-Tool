using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GSS_Alert_Service
{
    public partial class Alert_Service_Configuration_Tool : Form
    {
        public bool isSFDCValidated = true;

        public Alert_Service_Configuration_Tool()
        {
            InitializeComponent();
        }

        private void Alert_Service_Configuration_Tool_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
        }

        private void SFDCUsername_TextChanged(object sender, EventArgs e)
        {
            if (Alert_Service.DebugEnabled)
            {
                Alert_Service.eventLog1.WriteEntry(string.Format("Debug:/nSFDC Username changed!"), EventLogEntryType.Information, 0);
            }
            isSFDCValidated = false;
        }

        private void SFDCPassword_TextChanged(object sender, EventArgs e)
        {
            if (Alert_Service.DebugEnabled)
            {
                Alert_Service.eventLog1.WriteEntry(string.Format("Debug:/nSFDC Password changed!"), EventLogEntryType.Information, 0);
            }
            isSFDCValidated = false;
        }

        private void SFDCToken_TextChanged(object sender, EventArgs e)
        {
            if (Alert_Service.DebugEnabled)
            {
                Alert_Service.eventLog1.WriteEntry(string.Format("Debug:/nSFDC Token changed!"), EventLogEntryType.Information, 0);
            }
            isSFDCValidated = false;
        }

        private void SFDCSave_Click(object sender, EventArgs e)
        {

        }
    }
}
