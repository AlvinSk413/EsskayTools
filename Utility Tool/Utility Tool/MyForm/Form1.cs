using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MyForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<System.Windows.Forms.Control> store = new List<System.Windows.Forms.Control>();
        List<string> funct = new List<string>();
        //////To Remove the Previously Displayed sub form in the Main Form Panel.////////
        private void remove()
        {
            for (int j = 0; j < store.Count; j++)
            {
                panel1.Controls.Remove(store[j]);
            }
        }
        //////Align the Sub Form to center in the main form//////
        private void Align_Center(System.Windows.Forms.Control a)
        {
            
            a.Location = new Point(
            panel1.ClientSize.Width / 2 - a.Size.Width / 2,
            panel1.ClientSize.Height / 2 - a.Size.Height / 2);
            a.Anchor = AnchorStyles.None;
           
        }

        private void ToolMessage(string s)
        {

            toolTip1.IsBalloon = true;
            toolTip1.SetToolTip(pictureBox2, s);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Attribute Processor")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                att.frmattributeprocessor tool1 = new att.frmattributeprocessor(); 
                tool1.TopLevel = false;
                tool1.FormBorderStyle = FormBorderStyle.None;
                
                panel1.Controls.Add(tool1);
                Align_Center(tool1);
                tool1.Visible = true;
                //ToolMessage("Creates multiple attribute files for different phases.");
                store.Add(tool1);
                //Form1.ActiveForm.Text = "Attribute Processor";
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Bolt Length Stick Check")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                BoltLengthStickCheck.BoltLengthStickCheck tool2 = new BoltLengthStickCheck.BoltLengthStickCheck();
                tool2.TopLevel = false;
                tool2.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool2);
                Align_Center(tool2);
                tool2.Visible = true;
                //ToolMessage("Checks whether all the bolts have\nadequate stick thru length.");
                store.Add(tool2);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }
            else if (comboBox1.Text == "Bolt Part Checker")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                BoltPartChecker.BoltPartChecker tool3 = new BoltPartChecker.BoltPartChecker();
                tool3.TopLevel = false;
                tool3.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool3);
                Align_Center(tool3);
                tool3.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool3);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Connection Summary")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                ConnectionSummary.ConnectionSummary tool4 = new ConnectionSummary.ConnectionSummary();
                tool4.TopLevel = false;
                tool4.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool4);
                Align_Center(tool4);
                tool4.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool4);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Embed Plate Count From Wall")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                EmbedPlateCountFromWall.EmbedPlateCountFromWall tool5 = new EmbedPlateCountFromWall.EmbedPlateCountFromWall();
                tool5.TopLevel = false;
                tool5.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool5);
                Align_Center(tool5);
                tool5.Visible = true;
                ToolMessage("Counts the number of steel members\nembedded inside the wall.");
                store.Add(tool5);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }


            else if (comboBox1.Text == "Near Far Shear Tab Finder")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                NearFarShearTabFinder.NearFarShearTabFinder tool6 = new NearFarShearTabFinder.NearFarShearTabFinder();
                tool6.TopLevel = false;
                tool6.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool6);
                Align_Center(tool6);
                tool6.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool6);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "NPTF Check")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                NPTFCheck.NPTFCheck tool7 = new NPTFCheck.NPTFCheck();
                tool7.TopLevel = false;
                tool7.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool7);
                Align_Center(tool7);
                tool7.Visible = true;
                ToolMessage("Checks the drawings for showing “NPTF”, if the\nBeam in the Tekla model has studs in the top flange.");
                store.Add(tool7);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Number Order")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                NumberOrder.NumberOrder tool8 = new NumberOrder.NumberOrder();

                tool8.TopLevel = false;
                tool8.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool8);
                Align_Center(tool8);
                tool8.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool8);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Package Checker")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                PackageChecker.PackageChecker tool9 = new PackageChecker.PackageChecker();
                tool9.TopLevel = false;
                tool9.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool9);
                Align_Center(tool9);
                tool9.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool9);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Plane Rotation Depth Info")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                PlaneRotationDepthInfo.PlaneRotationDepthInfo tool10 = new PlaneRotationDepthInfo.PlaneRotationDepthInfo();
                tool10.TopLevel = false;
                tool10.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool10);
                Align_Center(tool10);
                tool10.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool10);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "ABM Length Vs Live Length")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                ABMLengthVSLiveLength.ABMLengthVSLiveLength tool11 = new ABMLengthVSLiveLength.ABMLengthVSLiveLength();
                tool11.TopLevel = false;
                tool11.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool11);
                Align_Center(tool11);
                tool11.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool11);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Kips Report 21.0")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                KipsReport21._0.KipsReport21 tool12 = new KipsReport21._0.KipsReport21();
                tool12.TopLevel = false;
                tool12.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool12);
                Align_Center(tool12);
                tool12.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool12);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Code Compilance")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                WindowsFormsApplication1.Main_form tool13 = new WindowsFormsApplication1.Main_form();
                tool13.TopLevel = false;
                tool13.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool13);
                Align_Center(tool13);
                tool13.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool13);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Connection Processor")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                ConnectionProcessor.frmconnectionprocess tool14 = new ConnectionProcessor.frmconnectionprocess();
                tool14.TopLevel = false;
                tool14.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool14);
                Align_Center(tool14);
                tool14.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool14);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }

            else if (comboBox1.Text == "Orientation Check")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                Orientation.frmOrientation tool15 = new Orientation.frmOrientation();
                tool15.TopLevel = false;
                tool15.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool15);
                Align_Center(tool15);
                tool15.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool15);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }
            else if (comboBox1.Text == "Quality Checking Tool")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                QMS.frmQC tool15 = new QMS.frmQC();
                tool15.TopLevel = false;
                tool15.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool15);
                Align_Center(tool15);
                tool15.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool15);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }
            else if (comboBox1.Text == "Drawing Log")
            {
                DateTime startTime = DateTime.Now;
                if (store.Count > 0)
                {
                    remove();
                }
                label1.Text = "Processing...";
                DrawingLog.frmdrglog tool16 = new DrawingLog.frmdrglog();
                tool16.TopLevel = false;
                tool16.FormBorderStyle = FormBorderStyle.None;
                panel1.Controls.Add(tool16);
                Align_Center(tool16);
                tool16.Visible = true;
                ToolMessage("Utillity Tool");
                store.Add(tool16);
                DateTime endTime = DateTime.Now;
                TimeSpan processtime = endTime.Subtract(startTime);
                label1.Text = "Ready...<" + processtime.ToString(@"mm\:ss") + ">";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBox1.Checked;
        }
    }
}
