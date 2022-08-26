using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
using System.Collections;

 namespace att
{
    public partial class frmattributeprocessor : Form
    {

        //esskay default variable 14Dec21
        public const string skApplicationName = "AttributeProcessor";
        public const string skApplicationVersion = "1.5";
        public string msgCaption = "";
        public string applicationlog = "";
        //Ver1.2 15/Dec/2021 Vijayakumar
        //Santosh Kumar B request for changing phase or other inside the attibute file is now added now.
        //Ver1.3 15/Dec/2021 Vijayakumar
        //if ((chkinsidealso.Checked == false) || (line.ToUpper().IndexOf("VERSION") >= 0) || (line.ToUpper().IndexOf("COUNT") >= 0))

        public const string logpath = "d:\\esskay;\\\\192.168.2.3\\Automation Log";


        //registry 14Dec21
        public const string userRoot = "HKEY_CURRENT_USER\\SOFTWARE\\ESSKAY";
        string keyName = "";


        public frmattributeprocessor()
        {
            InitializeComponent();
            msgCaption = skApplicationName + " Ver." + skApplicationVersion + "Copyright ESSKAY Inc.@2021";
            this.Text = skApplicationName + " Ver." + skApplicationVersion;
            keyName = userRoot + "\\" + skApplicationName;
            Registry.SetValue(keyName, "Version", skApplicationVersion);

            txtphase.Text = "Test";
            txtphase.Text = "";

            txtstart.Text = "Test";
            txtstart.Text = "";

            txtend.Text = "Test";
            txtend.Text = "";

            //load existing stored values
            string regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "Path", "");
            if (regvalue != "")
                txtattributefolder.Text = regvalue;

            regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "Start", "");
            if (regvalue != "")
                txtstart.Text = regvalue;

            regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "End", "");
            if (regvalue != "")
                txtend.Text = regvalue;

            regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "Find", "");
            if (regvalue != "")
                txtattfind.Text = regvalue;

            regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "Replace", "");
            if (regvalue != "")
                txtattreplace.Text = regvalue;

            regvalue = string.Empty;
            regvalue = (string)Registry.GetValue(keyName, "Phase", "");
            if (regvalue != "")
                txtphase.Text = regvalue;

            msgCaption = skApplicationName + " Ver." + skApplicationVersion;

            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Initialize", "");
        }

        private void butclear_Click(object sender, EventArgs e)
        {
            lblsbar.Text = "Please wait,processing...";
            skWinLib.accesslog( skApplicationName, skApplicationVersion, "Clear", "");
            txtattreplace.Text = "";
            txtattfind.Text = "";
            txtattributefolder.Text = "";
            txtend.Text = "";
            txtstart.Text = "";
            txtstart.Text = "";
            txtphase.Text = "";
            txtphase.Enabled = true;

            chklstAttribute.Items.Clear();
            chklstswapattribute.Items.Clear();
            lblsbar.Text = "Ready";

        }

        
        private void CreatePhase_filter()
        {
            string sUserName = System.Environment.GetEnvironmentVariable("USERNAME");
            System.IO.DirectoryInfo di = new DirectoryInfo("./" + txtattributefolder.Text);
            string fname = di.FullName + "\\" + sUserName + "_filter.SObjGrp";
            TextWriter stringWriter = new StringWriter();
            using (TextWriter streamWriter = new StreamWriter(fname))
            {
                streamWriter.WriteLine("TITLE_OBJECT_GROUP ");
                streamWriter.WriteLine("{");
                streamWriter.WriteLine("    Version= 1.04 ");
                streamWriter.WriteLine("    Count= 3 ");
                streamWriter.WriteLine("    SECTION_OBJECT_GROUP ");
                streamWriter.WriteLine("    {");
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        1");
                streamWriter.WriteLine("        co_part ");
                streamWriter.WriteLine("        proPrimaryPart ");
                streamWriter.WriteLine("        albl_Primary_part ");
                streamWriter.WriteLine("        == ");
                streamWriter.WriteLine("        albl_Equals ");
                streamWriter.WriteLine("        1 ");
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        Empty ");
                streamWriter.WriteLine("        }");
                streamWriter.WriteLine("    }");
                streamWriter.WriteLine("    SECTION_OBJECT_GROUP ");
                streamWriter.WriteLine("    {");
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        1");
                streamWriter.WriteLine("        co_part ");
                streamWriter.WriteLine("        proPHASE ");
                streamWriter.WriteLine("        albl_Phase ");
                streamWriter.WriteLine("        == ");
                streamWriter.WriteLine("        albl_Equals ");
                streamWriter.WriteLine(txtphase.Text);
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        Empty ");
                streamWriter.WriteLine("        }");
                streamWriter.WriteLine("    }");
                streamWriter.WriteLine("    SECTION_OBJECT_GROUP ");
                streamWriter.WriteLine("    {");
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        1");
                streamWriter.WriteLine("        co_part ");
                streamWriter.WriteLine("        proPrimaryPart ");
                streamWriter.WriteLine("        Range ");
                streamWriter.WriteLine("        == ");
                streamWriter.WriteLine("        albl_Equals ");
                streamWriter.WriteLine(txtphase.Text);
                streamWriter.WriteLine("        0 ");
                streamWriter.WriteLine("        Empty ");
                streamWriter.WriteLine("        }");
                streamWriter.WriteLine("    }");

                streamWriter.Close();
            }

        }

        private void butChangeAttribute_Click(object sender, EventArgs e)
        {
            lblsbar.Text = "Please wait,processing...";
            //strore default values
            skWinLib.accesslog( skApplicationName, skApplicationVersion, "ChangePhase", "");
            Registry.SetValue(keyName, "Path", txtattributefolder.Text);

            Registry.SetValue(keyName, "Find", txtattfind.Text);
            Registry.SetValue(keyName, "Replace", txtattreplace.Text);
            Registry.SetValue(keyName, "Phase", txtphase.Text);
            //read the existing files
            string path = txtattributefolder.Text;
            string newpath = txtattributefolder.Text + "\\newAttibutes\\";

            string newattributefolder = path + "\\newattibute\\";
            int addct = 0;
            int updct = 0;

            if (System.IO.Directory.Exists(newattributefolder))
            {
                DialogResult dialogResult = MessageBox.Show(newattributefolder + " already Exists!!!. Do you want to overwrite???", msgCaption, MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {

                    return;
                }
            }
            else
                System.IO.Directory.CreateDirectory(newattributefolder);

            int ct = 0;
            try
            {
                string newphase = string.Empty;
                string isalreadychk = "\\";
                string oldattributefilename = "";
                string newattributefilename = "";
                lblsbar.Text = "Please wait,changing attributes...";

                foreach (int indexChecked in this.chklstAttribute.CheckedIndices)

                {

                    newphase = chklstAttribute.Items[indexChecked].ToString();
//
 //               }
 //               for (int i = 0; i < chklstAttribute.Items.Count; i++)
                //{
                    newphase = chklstAttribute.Items[indexChecked].ToString();
                    
                    string[] repdirectoryList = path.Split(new Char[] { ';' });
                    foreach (string repdirectory in repdirectoryList)
                    {

                        if (Directory.Exists(repdirectory))
                        {
                            DirectoryInfo di = new DirectoryInfo(repdirectory);
                            FileInfo[] fi = di.GetFiles("*.*");
                            foreach (FileInfo nextFile in fi)
                            {
                                ArrayList newattributefiledata = new ArrayList();
                                if (nextFile != null)
                                {

                                    // Create an instance of StreamReader to read from a file.
                                    // The using statement also closes the StreamReader.
                                    bool chkflag1 = false;
                                    bool chkflag2 = false;
                                    bool chkflag3 = false;
                                    bool phaseflag = false;
                                    oldattributefilename = nextFile.Name;
                                    using (StreamReader sr = new StreamReader(path + "\\" + oldattributefilename))
                                    {
                                        String line;
                                        int linect = 0;
                                        addct++;
                                        while ((line = sr.ReadLine()) != null)
                                        {

                                            linect++;
                                            if (chkflag1 == true && chkflag2 == true && chkflag3 == true && phaseflag == false && linect == 79)
                                            {
                                                phaseflag = true;
                                                chkflag1 = false;
                                                chkflag2 = false;
                                                chkflag3 = false;
                                                newattributefiledata.Add("        " + newphase);
                                            }
                                            else
                                            {
                                                if (line.Contains("albl_Phase"))
                                                {
                                                    chkflag1 = true;
                                                    linect = 76;
                                                }
                                                if (line.Contains("=="))
                                                    chkflag2 = true;
                                                if (line.Contains("albl_Equals"))
                                                    chkflag3 = true;

                                                if ((chkinsidealso.Checked == false) || (line.ToUpper().IndexOf("VERSION") >= 0) || (line.ToUpper().IndexOf("COUNT") >= 0))
                                                    newattributefiledata.Add(line);
                                                else
                                                {
                                                    if (line.ToUpper().IndexOf(txtphase.Text.ToUpper()) != -1)
                                                    {
                                                        newattributefiledata.Add(line.Replace(txtphase.Text, newphase));
                                                        updct++;
                                                    }                                                        
                                                    else
                                                        newattributefiledata.Add(line);

                                                }
                                                
                                            }
                                        }
                                        //sr.flush();
                                        sr.Close();
                                        //create new attribute file
                                        {

                                            newattributefilename = oldattributefilename.Replace("_" + txtphase.Text, "_" + newphase);

                                            if (newattributefilename == oldattributefilename)
                                            {
                                                int isfound = isalreadychk.IndexOf("\\" + oldattributefilename + "\\");
                                                if (isfound == -1)

                                                {
                                                    MessageBox.Show(oldattributefilename + " does not have phase: _" + txtphase.Text + " in filename hence not changed.\n\nFor your information.", msgCaption);
                                                    isalreadychk = isalreadychk + oldattributefilename + "\\";
                                                }

                                            }
                                            else
                                                using (StreamWriter newattribefile = new StreamWriter(newattributefolder + newattributefilename, true))
                                                {
                                                    foreach (string attdata in newattributefiledata)
                                                    {
                                                        newattribefile.WriteLine(attdata);
                                                    }
                                                    newattribefile.Close();
                                                    ct++;
                                                }

                                        }
                                    }


                                }
                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
            skWinLib.worklog(skApplicationName, skApplicationVersion, "ChangePhase", "; File Count: " + addct.ToString() + "; Change Count: " + updct.ToString());
            lblsbar.Text = "Ready";
            MessageBox.Show("Process completed.\n\n" + ct.ToString() + " attributes created.", msgCaption);
        }

        private bool Isfindrepalcealreadyxists(string reportname)
        {
            bool flag = false;
            for (int i = 0; i < chklstswapattribute.Items.Count; i++)
            {
                string chklistreportname = chklstswapattribute.Items[i].ToString();
                if (chklistreportname == reportname)
                {
                    flag = true;
                    return flag;
                }
            }


            return flag;
        }

        private bool Isalreadyxists(string reportname)
        {
            bool flag = false;
            for (int i = 0; i < chklstAttribute.Items.Count; i++)
            {
                string chklistreportname = chklstAttribute.Items[i].ToString();
                if (chklistreportname == reportname)
                {
                    flag = true;
                    return flag;
                }
            }


            return flag;
        }

        public Boolean IsNumber(String s)
        {
            Boolean value = true;
            if (s == String.Empty || s == null)
            {
                value = false;
            }
            else
            {
                foreach (Char c in s.ToCharArray())
                {
                    value = value && Char.IsDigit(c);
                    if (value == false)
                        return value;
                }
            }
            return value;
        }
        private void butadd_Click(object sender, EventArgs e)
        {
            lblsbar.Text = "Please wait,processing...";
            skWinLib.accesslog( skApplicationName, skApplicationVersion, "Add", "");
            int addct = 0;

            int start = -1;
            if (IsNumber(txtstart.Text))
                start = Convert.ToInt32(txtstart.Text);

            int end = -1;
            if (IsNumber(txtend.Text))
                end = Convert.ToInt32(txtend.Text);



            bool abortflag = false;

            //check input
            if (start <= 0)
            {
                MessageBox.Show("Start number should be greater than zero and can't be null.\n\n Aborting Now!!!");
                abortflag = true;
                goto mySkip;
            }
            if (end <= 0)
            {
                MessageBox.Show("End number should be greater than zero and can't be null.");
                abortflag = true;
                goto mySkip;
            }
            if (end < start)
            {
                MessageBox.Show("End number missing or should be greater than Start Number.");
                abortflag = true;

            }
        mySkip:
            if (abortflag == false)
            {
                lblsbar.Text = "Please wait,adding adding...";
                for (int i = start; i <= end; i++)
                {
                    if (Isalreadyxists(i.ToString()) == false)
                    {
                        chklstAttribute.Items.Add(i);
                        chklstAttribute.SetItemChecked(chklstAttribute.Items.Count -1, true);
                        addct++;
                    }

                }
                Registry.SetValue(keyName, "Start", txtstart.Text);
                Registry.SetValue(keyName, "End", txtend.Text);
                //clear inputs
                txtphase.Enabled = false;
                txtstart.Text = "";
                txtend.Text = "";

                skWinLib.worklog(skApplicationName, skApplicationVersion, "Add", "; Phase Count: " + addct.ToString() );

            }


            lblsbar.Text = "Ready";
        }

        private void butswapadd_Click(object sender, EventArgs e)
        {


            string myfind = txtattfind.Text;
            string myreplace = txtattreplace.Text;

            bool abortflag = false;
            //check input
            if (myfind.Length <= 0)
            {
                MessageBox.Show("Find can't be null.\n\n Aborting Now!!!");
                abortflag = true;
                goto mySkip;
            }
            if (myreplace.Length <= 0)
            {
                MessageBox.Show("Replace can't be null.\n\n Aborting Now!!!");
                abortflag = true;
            }
        mySkip:
            if (abortflag == false)
            {
                if (Isfindrepalcealreadyxists(txtattfind.Text + " >> " + txtattreplace.Text) == false)
                    chklstswapattribute.Items.Add(txtattfind.Text + " >> " + txtattreplace.Text);
                else
                    MessageBox.Show(txtattreplace.Text + " already exist hence not added.");

                //clear inputs
                //txtattfind.Text = "";
                txtattreplace.Text = "";
            }


        }




        private void txtphase_TextChanged(object sender, EventArgs e)
        {
            if ((txtphase.Text.ToString().Trim().Length >= 1) && (txtattributefolder.Text.ToString().Trim().Length >= 1))
                butChangeAttribute.Enabled = true;
            else
                butChangeAttribute.Enabled = false;
        }

        private void txtattributefolder_TextChanged(object sender, EventArgs e)
        {
            if ((txtphase.Text.ToString().Trim().Length >= 1) && (txtattributefolder.Text.ToString().Trim().Length >= 1))
                butChangeAttribute.Enabled = true;
            else
                butChangeAttribute.Enabled = false;
        }

        private void txtstart_TextChanged(object sender, EventArgs e)
        {
            if ((txtstart.Text.ToString().Trim().Length >= 1) && (txtend.Text.ToString().Trim().Length >= 1))
                butphaseadd.Enabled = true;
            else
                butphaseadd.Enabled = false;
        }

        private void txtend_TextChanged(object sender, EventArgs e)
        {
            if ((txtstart.Text.ToString().Trim().Length >= 1) && (txtend.Text.ToString().Trim().Length >= 1))
                butphaseadd.Enabled = true;
            else
                butphaseadd.Enabled = false;
        }

        private void frmattributeprocessor_Load(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
