using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ASCOM.Utilities;
using ASCOM.scopefocus;

namespace ASCOM.scopefocus
{
    [ComVisible(false)]					// Form not registered for COM!
    public partial class SetupDialogForm : Form
    {
        public SetupDialogForm()
        {
            InitializeComponent();
            // Initialise current values of user settings from the ASCOM Profile
            InitUI();
        }

        private void cmdOK_Click(object sender, EventArgs e) // OK button event handler
        {

            if (textBox2.Text == "")
            {
                MessageBox.Show("You must specify a value for Steps per Degree");
               
                return;
            }
            else
            {
                using (ASCOM.Utilities.Profile p = new Utilities.Profile())
                {
                    p.DeviceType = "Rotator";
                    p.WriteValue(Rotator.driverID, "ComPort", (string)comboBoxComPort.SelectedItem);
                    p.WriteValue(Rotator.driverID, "SetPos", checkBox1.Checked.ToString());
                    // 6-16-16 added 2 lines below
                    //   p.WriteValue(Rotator.driverID, "Reverse", reverseCheckBox1.Checked.ToString());  // motor sitting shaft up turns clockwise with increasing numbers if NOT reversed
                    p.WriteValue(Rotator.driverID, "ContHold", checkBox2.Checked.ToString());


                    p.WriteValue(Rotator.driverID, "StepsPerDegree", textBox2.Text.ToString());
                    //   p.WriteValue(Focuser.driverID, "RPM", textBoxRpm.Text);
                    if (checkBox1.Checked)
                    {
                        p.WriteValue(Rotator.driverID, "Pos", textBox1.Text.ToString());
                    }
                    //    p.WriteValue(Focuser.driverID, "TempDisp", radioCelcius.Checked ? "C" : "F");
                }
                Dispose();




                // Place any validation constraint checks here
                // Update the state variables with results from the dialogue
                Rotator.comPort = (string)comboBoxComPort.SelectedItem;
                Rotator.traceState = chkTrace.Checked;
            }
        }
        private void cmdCancel_Click(object sender, EventArgs e) // Cancel button event handler
        {
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e) // Click on ASCOM logo event handler
        {
            try
            {
                System.Diagnostics.Process.Start("http://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (System.Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }

        private void checkTextBox() //don't allow close with steps/dgree blank or set pos checked and blank position
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
                cmdOK.Enabled = false; 
            else if ((string.IsNullOrWhiteSpace(textBox1.Text) && checkBox1.Checked))
                cmdOK.Enabled = false;
            else 
                cmdOK.Enabled = true;
        }


        private void InitUI()
        {
           
            chkTrace.Checked = Rotator.traceState;
            // set the list of com ports to those that are currently available
            comboBoxComPort.Items.Clear();
            comboBoxComPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());      // use System.IO because it's static
            // select the current port if possible
            if (comboBoxComPort.Items.Contains(Rotator.comPort))
            {
                comboBoxComPort.SelectedItem = Rotator.comPort;
            }
            using (ASCOM.Utilities.Profile p = new Utilities.Profile())
            {
                p.DeviceType = "Rotator";
                textBox2.Text = p.GetValue(Rotator.driverID, "StepsPerDegree");
                if (p.GetValue(Rotator.driverID, "ContHold") == "True")
                    checkBox2.Checked = true;
                else
                    checkBox2.Checked = false;
            }
            if (!checkBox1.Checked)
                textBox1.Enabled = false;
            checkTextBox();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool enable = false;
            if (checkBox1.Checked)
                enable = true;


            //  label2.Enabled = enable;
            textBox1.Enabled = enable;
            checkTextBox();
        }

        private void chkTrace_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTrace.Checked)
                Rotator.traceState = true;
            else
                Rotator.traceState = false;
        }



        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            checkTextBox();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            checkTextBox();
        }
    }
}