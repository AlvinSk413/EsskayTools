using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Model.UI;

namespace Orientation
{
    public partial class frmOrientation : Form
    {
        public static string skApplicationName = "Orientation Check";
        public static string skApplicationVersion = "2208.04";
        public string skTSVersion = "2021";

        public string msgCaption = "";
        public TSM.Model MyModel;

        public frmOrientation()
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");
            InitializeComponent();
            chkErrorOnly.Checked = true;
            chkUIOption.Checked = true;
            this.TopMost = true;
        }


        private void frmOrientation_Load(object sender, EventArgs e)
        {
            msgCaption = this.Text + " " + skApplicationVersion;

            //Esskay Security Validation
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");

            MyModel = skTSLib.ModelAccess.ConnectToModel();

            if (MyModel == null)
            {
                skWinLib.accesslog(skApplicationName, skApplicationVersion, "EsskayValidation", "Could not establish connection with Tekla [ " + skTSVersion + " ]."/*, skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration*/);
                MessageBox.Show("Could not establish connection with Tekla [ " + skTSVersion + " ].\n\nEnsure Tekla [ " + skTSVersion + " ] model is open,if still issue's contact support team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                System.Environment.Exit(0);
                //esskayjn = txtproject.Text.ToString();

            }

            else
            {

                this.Text = skApplicationName + " [" + skApplicationVersion + "]";
                lblsbar2.Text = skWinLib.username;
                chkUIOption.Checked = false;

            }

        }



        private void OrientationChecker()
        {
            dgOrientation.Rows.Clear();

            Tekla.Structures.Model.UI.ModelObjectSelector MyModelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();

            //filter to be added for better processing time will be done next stage 
            ModelObjectEnumerator MyModelObjectEnumerator = MyModelObjectSelector.GetSelectedObjects();

            pbar1.Visible = true;
            pbar1.Maximum = MyModelObjectEnumerator.GetSize();
            pbar1.Value = 0;

            //Check for column, post, beam orientation are as per esskay requirement (Check1 - Checked is false)
            while (MyModelObjectEnumerator.MoveNext())
            {
                ValidateMemberOrientation(MyModelObjectEnumerator);
                pbar1.Value = pbar1.Value + 1;
            }

            //setting for data grid
            dgOrientation.AutoResizeRows();
            dgOrientation.AutoResizeColumns();
            dgOrientation.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders);
            pbar1.Visible = false;

        }


        //  private void ValidateMemberOrientation(Beam MyBeam)
        private void ValidateMemberOrientation(ModelObjectEnumerator MyModelObjectEnumerator)
        {
            Beam MyBeam = MyModelObjectEnumerator.Current as Beam;
            if (MyBeam != null)
            {
                bool IsError = true;
                string PosPlane = MyBeam.Position.Plane.ToString();
                string PosRotation = MyBeam.Position.Rotation.ToString();
                string PosDepth = MyBeam.Position.Depth.ToString();

                if (chkUIOption.Checked == true)
                {
                    if (PosPlane == cmbPlane.Text && PosRotation == cmbRotation.Text && PosDepth == cmbDepth.Text)
                        IsError = false;
                }
                else if (chkTSMemberType.Checked == true)
                {
                    if (MyBeam.Type == Beam.BeamTypeEnum.BEAM)
                    {
                        if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "BEHIND")
                            IsError = false;
                    }
                    else if (MyBeam.Type == Beam.BeamTypeEnum.COLUMN)
                    {
                        if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "MIDDLE")
                            IsError = false;
                        else if (PosPlane == "MIDDLE" && PosRotation == "BACK" && PosDepth == "MIDDLE")
                            IsError = false;
                    }
                }
                else
                {
                    if (!MyBeam.Name.ToUpper().Contains("PL"))
                    {
                        if (MyBeam.Name.ToUpper().Contains("COLUMN") || MyBeam.Name.ToUpper().Contains("BEAM"))
                        {
                            if (MyBeam.Name.ToUpper().Contains("BEAM"))
                            {
                                {
                                    if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "BEHIND")
                                        IsError = false;
                                }
                            }

                            else if (MyBeam.Name.ToUpper().Contains("COLUMN"))
                            {
                                if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "MIDDLE")
                                    IsError = false;
                                else if (PosPlane == "MIDDLE" && PosRotation == "BACK" && PosDepth == "MIDDLE")
                                    IsError = false;
                            }
                        }
                    }
                }

                if (IsError == true)
                {
                    UpdateReportUI(MyBeam, true);
                }
                else
                {
                    if (chkErrorOnly.Checked == false)
                        UpdateReportUI(MyBeam, IsError);
                }
            }


            else
            {
                //beam is null hence error

                dgOrientation.Rows.Add();
                DataGridViewRow myrow = dgOrientation.Rows[dgOrientation.Rows.Count - 1];
                myrow.HeaderCell.Value = (dgOrientation.Rows.Count).ToString();

                ModelObject MyModelObject = MyModelObjectEnumerator.Current as ModelObject;
                string errorType = "Error";
                if (MyModelObject.Identifier.IsValid())
                {
                    myrow.Cells["guid"].Value = MyModelObject.Identifier.GUID.ToString();
                    errorType = errorType + "-" + MyModelObject.GetType().Name.ToString();
                }
                else
                    myrow.Cells["guid"].Value = "guid:";

                //check and update remark color
                System.Drawing.Color myColor = System.Drawing.Color.Red;
                myrow.Cells["Remark"].Value = errorType;
                myrow.Cells["Remark"].Style.BackColor = myColor;

            }

            #region
            //if (MyBeam != null)
            //{
            //    bool IsError = true;
            //    string PosPlane = MyBeam.Position.Plane.ToString();
            //    string PosRotation = MyBeam.Position.Rotation.ToString();
            //    string PosDepth = MyBeam.Position.Depth.ToString();


            //    if (chkUIOption.Checked == true)
            //    {
            //        if (PosPlane == cmbPlane.Text && PosRotation == cmbRotation.Text && PosDepth == cmbDepth.Text)
            //            IsError = false;
            //    }

            //    else if (chkTSMemberType.Checked == true)
            //    {
            //        if (MyBeam.Type == Beam.BeamTypeEnum.BEAM)
            //        {
            //            if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "BEHIND")
            //                IsError = false;
            //        }
            //        else if (MyBeam.Type == Beam.BeamTypeEnum.COLUMN)
            //        {
            //            if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "MIDDLE")
            //                IsError = false;
            //            else if (PosPlane == "MIDDLE" && PosRotation == "BACK" && PosDepth == "MIDDLE")
            //                IsError = false;
            //        }
            //    }

            //    else
            //    {
            //        if (!MyBeam.Name.ToUpper().Contains("PL"))
            //        {
            //            if (MyBeam.Name.ToUpper().Contains("COLUMN") || MyBeam.Name.ToUpper().Contains("BEAM"))
            //            {
            //                if (MyBeam.Name.ToUpper().Contains("BEAM"))
            //                {
            //                    {
            //                        if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "BEHIND")
            //                            IsError = false;
            //                    }
            //                }

            //                else if (MyBeam.Name.ToUpper().Contains("COLUMN"))
            //                {
            //                    if (PosPlane == "MIDDLE" && PosRotation == "TOP" && PosDepth == "MIDDLE")
            //                        IsError = false;
            //                    else if (PosPlane == "MIDDLE" && PosRotation == "BACK" && PosDepth == "MIDDLE")
            //                        IsError = false;
            //                }

            //            }
            //        }
            //    }

            //    if (IsError == true)
            //    {
            //        UpdateReportUI(MyBeam, true);
            //    }
            //    else
            //    {
            //        if (chkErrorOnly.Checked == false)
            //            UpdateReportUI(MyBeam, IsError);

            //    }
            //}

            //else
            //{
            //    //beam is null hence error

            //    dgOrientation.Rows.Add();
            //    DataGridViewRow myrow = dgOrientation.Rows[dgOrientation.Rows.Count - 1];
            //    myrow.HeaderCell.Value = (dgOrientation.Rows.Count).ToString();
            //    //check and update remark color
            //    System.Drawing.Color myColor = System.Drawing.Color.Red;
            //    myrow.Cells["Remark"].Value = "Error-CheckType";
            //    myrow.Cells["Remark"].Style.BackColor = myColor;

            //}
            #endregion

        }

        private void UpdateReportUI(Beam MyBeam, bool IsError)
        {
            //get the phase number
            MyBeam.GetPhase(out TSM.Phase MyPhase);

            //add a new empty row
            dgOrientation.Rows.Add();
            DataGridViewRow myrow = dgOrientation.Rows[dgOrientation.Rows.Count - 1];
            myrow.HeaderCell.Value = (dgOrientation.Rows.Count).ToString();
            myrow.Cells["guid"].Value = MyBeam.Identifier.GUID.ToString();
            myrow.Cells["partname"].Value = MyBeam.Name.ToString();
            myrow.Cells["profile"].Value = MyBeam.Profile.ProfileString;
            myrow.Cells["phase"].Value = MyPhase.PhaseNumber.ToString();
            myrow.Cells["XSMark"].Value = MyBeam.GetPartMark().ToString();


            //member position
            myrow.Cells["OnPlane"].Value = MyBeam.Position.Plane.ToString();
            myrow.Cells["Rotation"].Value = MyBeam.Position.Rotation.ToString();
            myrow.Cells["AtDepth"].Value = MyBeam.Position.Depth.ToString();

            //check main part
            if (MyBeam.GetAssembly().GetMainPart().Identifier.GUID == MyBeam.Identifier.GUID)
                myrow.Cells["Isprimary"].Value = 1;
            else
                myrow.Cells["Isprimary"].Value = 0;


            //check and update remark color
            System.Drawing.Color myColor = System.Drawing.Color.Green;
            myrow.Cells["Remark"].Value = "Ok";
            if (IsError == true)
            {
                myrow.Cells["Remark"].Value = "Error-Check";
                myColor = System.Drawing.Color.Red;
            }
            myrow.Cells["Remark"].Style.BackColor = myColor;

        }



        private void chkontop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkontop.Checked;
        }



        private void btncheck_Click(object sender, EventArgs e)
        {
            //PicBoxOrientaion.Visible = false;
            DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "EsskayOrientationChecking", "", "Pl.Wait, procssing for routing rule...", "btnRoutingRules_Click", lblsbar1, lblsbar2);

            OrientationChecker();
            PicBoxOrientaion.Visible = false;
            skTSLib.setCompletion(skApplicationName, skApplicationName, "Complete", "", "Total : " + dgOrientation.Rows.Count.ToString(), "", lblsbar1, lblsbar2, s_tm);

        }

        private void btnexport_Click(object sender, EventArgs e)
        {
            //create the sample.txt file in new folder 
                                  

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\\New folder\\sample.txt");

            try
            {
                string sLine = "";

                for (int r = 0; r <= dgOrientation.Rows.Count - 1; r++)
                {

                    for (int c = 0; c <= dgOrientation.Columns.Count - 1; c++)
                    {
                        sLine = sLine + dgOrientation.Rows[r].Cells[c].Value;
                        if (c != dgOrientation.Columns.Count - 1)
                        {
                            sLine = sLine + "   -   ";
                        }
                    }

                    //The exported text is written to the text file, one line at a time.

                    file.WriteLine(sLine);
                    sLine = "";
                }

                file.Close();
                System.Windows.Forms.MessageBox.Show("Export Complete.", "Program Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            catch (System.Exception err)
            {
                System.Windows.Forms.MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                file.Close();
            }
        }

        private void btnclear_Click(object sender, EventArgs e)
        {

            //progress bar off
            pbar1.Visible = true;

            //clear dgreport
            dgOrientation.Rows.Clear();


            //set status message
            lblsbar1.Text = "Ready";
            lblsbar2.Text = skWinLib.username;

            //show the image for user
            PicBoxOrientaion.Visible = true;

            //clear the various option
            chkUIOption.Checked = false;
            cmbPlane.Text = "";
            cmbRotation.Text = "";
            cmbDepth.Text = "";
            chkTSMemberType.Checked = false;

        }

        private void dgOrientation_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            skTSLib.selectDataGridselectedrows(dgOrientation, 0, "Guid", true);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void chkUIOption_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUIOption.Checked == true)
                chkTSMemberType.Checked = false;
            //update groupbox
            grpuseroption.Enabled = chkUIOption.Checked;

        }

        private void chkTSMemberType_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTSMemberType.Checked == true)
                chkUIOption.Checked = false;
            //update groupbox
            grpuseroption.Enabled = chkUIOption.Checked;
        }

        private void btnmempos_Click(object sender, EventArgs e)
        {
            Tekla.Structures.Model.UI.ModelObjectSelector MyModelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();

            //filter to be added for better processing time will be done next stage 
            ModelObjectEnumerator MyModelObjectEnumerator = MyModelObjectSelector.GetSelectedObjects();
            while (MyModelObjectEnumerator.MoveNext())
            {
                Beam MyBeam = MyModelObjectEnumerator.Current as Beam;
                if (MyBeam.Identifier.IsValid())
                {
                    //member position
                    cmbDepth.Text = MyBeam.Position.Depth.ToString();
                    cmbPlane.Text = MyBeam.Position.Plane.ToString();
                    cmbRotation.Text = MyBeam.Position.Rotation.ToString();
                    return;
                }
            }

            try
            {
                Picker MyPicker = new Picker();
                Tekla.Structures.Model.ModelObject MyModelObject = MyPicker.PickObject(Picker.PickObjectEnum.PICK_ONE_PART, "Select single Beam :");
                if (MyModelObject != null)
                {
                    Beam MyBeam = MyModelObject as Beam;
                    if (MyBeam.Identifier.IsValid())
                    {
                        //member position
                        cmbDepth.Text = MyBeam.Position.Depth.ToString();
                        cmbPlane.Text = MyBeam.Position.Plane.ToString();
                        cmbRotation.Text = MyBeam.Position.Rotation.ToString();
                    }
                }
            }
            catch (Exception MyException)
            {
                cmbDepth.Text = "";
                cmbPlane.Text = "";
                cmbRotation.Text = "";
                if (MyException.Message.ToUpper() != "USER INTERRUPT")
                    MessageBox.Show("Picker Error " + "\n\n\nMessage  : " + MyException.Message.ToString());



            }
        }
    }
}