using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

// Tekla Structures namespaces
using Tekla.Structures;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using Tekla.Structures.ModelInternal;
using Tekla.Structures.Solid;
using Tekla.Structures.Geometry3d;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSMUI = Tekla.Structures.Model.UI;
using Tekla.Structures.Catalogs;
using Tekla.Structures.Drawing;
using TSD = Tekla.Structures.Drawing;

using MySql.Data.MySqlClient;


using System.Globalization;
using Tekla.Structures.Model.UI;

namespace DrawingLog
{
    public partial class frmdrglog : Form
    {

        public readonly Model _Model = new Model();
        public static GraphicsDrawer GraphicsDrawer = new GraphicsDrawer();
        public readonly static Tekla.Structures.Model.UI.Color TextColor = new Tekla.Structures.Model.UI.Color(1, 0, 1);


        public string skApplicationName = "DRGEDTCHKLOG";
        public string skApplicationVersion = "2206.27";

        public string skTSVersion = "2021";
        public bool skdebug = false;
        public int debugcount = 1;

        public string msgCaption = "";
        public string skApplicationLog = "";
        public string skWorkLog = "";


        //ncm setting for user access and interface setting
        public int id_emp = -1; //i.e id_emp = 297 for SKS360/P.Vijayakumar 
        public string id_empname = "-1"; //P.Vijayakumar
        public string emp_role = "-1"; //Modeller, Checker, Project Manager, Team Lead
        public int ui_stage = -1; //New, inprocess, back drafting, back checking, held, cancelled, completed
        public int id_stage = -1; //database id
        public int id_process = -1; //0 Modelling complete,5 DPM , 10 Checking, 20 PM, 30 Back Drafting, 40 PM, 50 Back Checking, 60 PM, 70 held, 75 cancelled, 80 RELOOP MODELLER, 85 DPM, 90 RELOOP CHECKER, 95 DPM, 100 COMPLETE, 110QA 
        public int id_project = -1; //database id
        public int id_project_report = -1; //id for that project and report
        public string esskayjn = "-1"; //SK1700, SK1701,etc...
        public string projectname = "-1"; //Project Basie
        public string client = string.Empty; //Prospectus 
        public string clientjn = string.Empty; //2623
        public string skreportname = string.Empty;
        public bool isowner = false;
        public bool sk_TOGGLE_ts_input = false;
        public bool sk_TOGGLE_ts_display = false;
        public bool sk_TOGGLE_project_display = false;
        public bool sk_pdf_file_display = false;
        public bool TeklaAdditionalColumnFlag = false;

        //static ip working fine
        public const string connString = "server=(address=192.168.1.62:3306,priority=100),(address=122.165.187.180:8006,priority=50),(address=103.92.203.198:8006,priority=50);user id = esskay; password=esskay1505;database=epmatdb";

        public MySqlConnection myconn = new MySqlConnection(connString);
        public MySqlCommand mycmd = new MySqlCommand();
        public string mysql = "";
        public MySqlDataReader myrdr = null;

        ArrayList projectmast = new ArrayList();
        ArrayList empmast = new ArrayList();
        ArrayList drg_detailer_data = new ArrayList();
        ArrayList drg_checker_data = new ArrayList();
        public TSM.Model MyModel;



        //there will be different activity for editing and checking hence myactivity will be modified based on user selection
        ArrayList myactivity = new ArrayList();
        ArrayList mydrglist = new ArrayList();

        TSD.DrawingHandler dh = new TSD.DrawingHandler();
        public frmdrglog()
        {
            InitializeComponent();

        }


        private void btnselectdrg_Click(object sender, EventArgs e)
        {
            ////inform the user about the drawing details 
            lblsbar1.Text = "Please Wait,processing [DWG]...";

            //get the selected drawing details like mark, revision, title, type
            ArrayList drgdata = GetDrawingDetail();


            //dgdrglist.Rows.Clear();

            //create drglist for avoiding duplicate entry
            mydrglist.Clear();
            foreach (DataGridViewRow myrow in dgdrglist.Rows)
            {
                string rowmark = string.Empty;
                string rowtype = string.Empty;
                if (myrow.Cells["DrgMark"].Value != null)
                    rowmark = myrow.Cells["DrgMark"].Value.ToString();

                if (myrow.Cells["DrgType"].Value != null)
                    rowtype = myrow.Cells["DrgType"].Value.ToString();

                if ((rowmark + rowtype).Length >=1)
                    mydrglist.Add(rowmark + rowtype);
            }
            

            foreach (string mydata in drgdata)
            {
                string[] mydatasplt = mydata.Split(new Char[] { '|' });
                //drgdata.Add("DWG|" + drawingType + "|" + currentDrawing.Mark + "|" + currentDrawing.Name + "|" + drgrev + "|" + currentDrawing.Title1 + "|" + currentDrawing.Title2 + "|" + currentDrawing.Title3 + "|" + currentDrawing.GetSheet().ToString() + "|" + currentDrawing.Layout.ToString() + "|" + currentDrawing.CreationDate.ToString() + "|" + currentDrawing.ModificationDate.ToString());
                string drgmark = mydatasplt[2].ToString().Replace("[","").Replace("]", "").Replace(".", "");
                string drgtype = mydatasplt[1].ToString().ToUpper();
                if (drgtype.Contains(".ASSEMBL"))
                    drgtype = "A";
                else if (drgtype.Contains(".SINGLE"))
                    drgtype = "W";
                else if (drgtype.Contains(".GA"))
                    drgtype = "G";
                else if (drgtype.Contains(".MULT"))
                    drgtype = "M";
                else if (drgtype.Contains(".CAST"))
                    drgtype = "C";
                string chkmarktype = drgmark + drgtype;
                if (!mydrglist.Contains(chkmarktype))
                {
                    mydrglist.Add(chkmarktype);
                    string drgname = mydatasplt[3].ToString();
                    string drgrev = mydatasplt[4].ToString();
                    int rowid = dgdrglist.Rows.Add();
                    dgdrglist.Rows[rowid].Cells["DrgMark"].Value = drgmark;
                    dgdrglist.Rows[rowid].Cells["DrgName"].Value = drgname;
                    dgdrglist.Rows[rowid].Cells["DrgRev"].Value = drgrev;
                    dgdrglist.Rows[rowid].Cells["DrgType"].Value = drgtype;

                    if (drgtype == "G")
                    {
                        string Title1 = mydatasplt[5].ToString().ToUpper();
                        if (Title1.Length >=1)
                            dgdrglist.Rows[rowid].Cells["DrgTitle1"].Value = Title1;
                        else
                            dgdrglist.Rows[rowid].Cells["DrgTitle1"].Value = "";
                    }

                    //windows login name
                    string winlogin = skWinLib.username;
                    dgdrglist.Rows[rowid].Cells["winlogin"].Value = winlogin;

                    //system id
                    dgdrglist.Rows[rowid].Cells["systemname"].Value = skWinLib.systemname;

                    //check if user have already selected empname in the selection box if selected then selected empname to be shown here
                    if (cmbusername.Text.Length >= 1)
                    {

                        winlogin = cmbusername.Text.Substring(cmbusername.Text.IndexOf("|") + 1);
                        dgdrglist.Rows[rowid].Cells["ownerlogin"].Value = winlogin;
                    }
                    else
                        dgdrglist.Rows[rowid].Cells["ownerlogin"].Value = winlogin;


                    //get empname
                    ArrayList tmp = getDBUserID(winlogin);
                    if (tmp.Count >= 1)
                    {
                        string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                        //epmat empname
                        dgdrglist.Rows[rowid].Cells["empname"].Value = empsplit[2].ToString();
                        //change fore color to red if winlogin user differs
                        if (empsplit[2].ToString() != lbluser.Text.ToString())
                            dgdrglist.Rows[rowid].Cells["empname"].Style.ForeColor = System.Drawing.Color.Red;
                        else
                            dgdrglist.Rows[rowid].Cells["empname"].Style.ForeColor = System.Drawing.Color.Black;

                    }

                    //check if user have already selected activity in the selection box if selected then selected activity to be shown here
                    if (cmbmyactivity.Text.Length >= 1)
                        dgdrglist.Rows[rowid].Cells["process"].Value = cmbmyactivity.Text.ToString();

                   

                }

            }
            lblsbar1.Text = "Ready";

        }

        //private void btnselectdrg_Click(object sender, EventArgs e)
        //{
        //    //get the selected drawing details like mark, revision, title, type
        //    ArrayList drgdata = GetDrawingDetail();


        //    lblsbar1.Text = "Please Wait,checking [DWG]...";
        //    ////inform the user about the drawing details 
        //    //dgdrglist.Rows.Clear();

        //    //create drglist for avoiding duplicate entry
        //    mydrglist.Clear();
        //    foreach (DataGridViewRow myrow in dgdrglist.Rows)
        //    {
        //        string rowmark = string.Empty;
        //        string rowtype = string.Empty;
        //        if (myrow.Cells["DrgMark"].Value != null)
        //            rowmark = myrow.Cells["DrgMark"].Value.ToString();

        //        if (myrow.Cells["DrgType"].Value != null)
        //            rowtype = myrow.Cells["DrgType"].Value.ToString();

        //        if ((rowmark + rowtype).Length >= 1)
        //            mydrglist.Add(rowmark + rowtype);
        //    }


        //    foreach (string mydata in drgdata)
        //    {
        //        string[] mydatasplt = mydata.Split(new Char[] { '|' });
        //        //drgdata.Add("DWG|" + drawingType + "|" + currentDrawing.Mark + "|" + currentDrawing.Name + "|" + drgrev + "|" + currentDrawing.Title1 + "|" + currentDrawing.Title2 + "|" + currentDrawing.Title3 + "|" + currentDrawing.GetSheet().ToString() + "|" + currentDrawing.Layout.ToString() + "|" + currentDrawing.CreationDate.ToString() + "|" + currentDrawing.ModificationDate.ToString());
        //        string drgmark = mydatasplt[2].ToString().Replace("[", "").Replace("]", "").Replace(".", "");
        //        string drgtype = mydatasplt[1].ToString().ToUpper();
        //        if (drgtype.Contains(".ASSEMBL"))
        //            drgtype = "A";
        //        else if (drgtype.Contains(".SINGLE"))
        //            drgtype = "W";
        //        else if (drgtype.Contains(".GA"))
        //            drgtype = "G";
        //        else if (drgtype.Contains(".MULT"))
        //            drgtype = "M";
        //        else if (drgtype.Contains(".CAST"))
        //            drgtype = "C";
        //        string chkmarktype = drgmark + drgtype;
        //        if (!mydrglist.Contains(chkmarktype))
        //        {
        //            mydrglist.Add(chkmarktype);
        //            string drgname = mydatasplt[3].ToString();
        //            string drgrev = mydatasplt[4].ToString();
        //            int rowid = dgdrglist.Rows.Add();
        //            dgdrglist.Rows[rowid].Cells["DrgMark"].Value = drgmark;
        //            dgdrglist.Rows[rowid].Cells["DrgName"].Value = drgname;
        //            dgdrglist.Rows[rowid].Cells["DrgRev"].Value = drgrev;
        //            dgdrglist.Rows[rowid].Cells["DrgType"].Value = drgtype;

        //            if (drgtype == "G")
        //            {
        //                string Title1 = mydatasplt[5].ToString().ToUpper();
        //                if (Title1.Length >= 1)
        //                    dgdrglist.Rows[rowid].Cells["Title1"].Value = Title1;
        //                else
        //                    dgdrglist.Rows[rowid].Cells["Title1"].Value = "";
        //            }

        //            //windows login name
        //            string winlogin = skWinLib.username;
        //            dgdrglist.Rows[rowid].Cells["winlogin"].Value = winlogin;

        //            //system id
        //            dgdrglist.Rows[rowid].Cells["systemid"].Value = skWinLib.systemname;

        //            //check if user have already selected empname in the selection box if selected then selected empname to be shown here
        //            if (cmbusername.Text.Length >= 1)
        //            {

        //                winlogin = cmbusername.Text.Substring(cmbusername.Text.IndexOf("|") + 1);
        //                dgdrglist.Rows[rowid].Cells["modempid"].Value = winlogin;
        //            }
        //            else
        //                dgdrglist.Rows[rowid].Cells["modempid"].Value = winlogin;


        //            //get empname
        //            ArrayList tmp = getDBUserID(winlogin);
        //            if (tmp.Count >= 1)
        //            {
        //                string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
        //                //epmat empname
        //                dgdrglist.Rows[rowid].Cells["empname"].Value = empsplit[2].ToString();
        //                //change fore color to red if winlogin user differs
        //                if (empsplit[2].ToString() != lbluser.Text.ToString())
        //                    dgdrglist.Rows[rowid].Cells["empname"].Style.ForeColor = System.Drawing.Color.Red;
        //                else
        //                    dgdrglist.Rows[rowid].Cells["empname"].Style.ForeColor = System.Drawing.Color.Black;

        //            }

        //            //check if user have already selected activity in the selection box if selected then selected activity to be shown here
        //            if (cmbmyactivity.Text.Length >= 1)
        //                dgdrglist.Rows[rowid].Cells["process"].Value = cmbmyactivity.Text.ToString();



        //        }

        //    }
        //    lblsbar1.Text = "Ready";

        //}
        //////testing
        ////string name = "lohesh.pdf";
        ////string file = "D:\\Vijayakumar\\01_Development\\SK-073-Drawing Log\\" + name;
        ////byte[] mybytes = System.IO.File.ReadAllBytes(file);

        //////mysql
        ////mycmd.CommandText = "INSERT INTO epmatdb.pocpdf(file,name) VALUES(@file, @name)";
        ////mycmd.Parameters.Clear();
        ////mycmd.Parameters.Add("@file", MySqlDbType.MediumBlob);
        ////mycmd.Parameters.Add("@name", MySqlDbType.MediumText);
        ////mycmd.Parameters["@file"].Value = mybytes;
        ////mycmd.Parameters["@name"].Value = name;

        //////est. conn
        ////mycmd.Connection = myconn;
        ////if (myconn.State == ConnectionState.Closed)
        ////myconn.Open();

        //////insert new record
        ////int flag = mycmd.ExecuteNonQuery();
        //////testing
        ////byte[] mybytes = null;

        //////est. conn
        ////mycmd.Connection = myconn;
        ////if (myconn.State == ConnectionState.Closed)
        ////myconn.Open();

        ////string sql = "SELECT id,file,name FROM epmatdb.pocpdf";
        ////mycmd.CommandText = sql;
        ////MySqlDataReader rdr = mycmd.ExecuteReader();

        ////while (rdr.Read())
        ////{
        ////mybytes = rdr[1] as byte[];
        ////string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        ////string name = "lohesh" + timestamp + ".pdf";
        ////string file = "D:\\Vijayakumar\\01_Development\\SK-073-Drawing Log\\" + name;

        ////System.IO.File.WriteAllBytes(file, mybytes);
        ////System.Diagnostics.Process.Start(file);
        ////}
        ////rdr.Close();
        ////myconn.Close();


        private void frmdrglog_Load(object sender, EventArgs e)
        {
            skdebug = skWinLib.DebugFlag();

            //Update User Interface Application Name and Latest Version
            //this.Text = skApplicationName + " [" + skApplicationVersion + "]";

            //MessageBox.Show("TS-2021 \n\n222");
            msgCaption = this.Text + " " + skApplicationVersion;

            //Esskay Security Validation
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");
            //MessageBox.Show("TS-2021 \n\n2");

            //Check whether mySQL connection established if not inform user.
            if (IsCheckconnection() == false)
            {
                skWinLib.accesslog(skApplicationName, skApplicationVersion, "EsskayValidation", "Could not establisth connection with MySQL.", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
                MessageBox.Show("Could not establisth connection with MySQL.\n\n\nCheck your connection again and ensure server is running, if still issue's contact support team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                System.Windows.Forms.Application.ExitThread();
            }
            else
            {
                //MessageBox.Show("TS-2021 \n\n222");
                MyModel = skTSLib.ModelAccess.ConnectToModel();

                if (MyModel == null)
                {
                    skWinLib.accesslog(skApplicationName, skApplicationVersion, "EsskayValidation", "Could not establish connection with Tekla [ " + skTSVersion + " ].", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
                    MessageBox.Show("Could not establish connection with Tekla [ " + skTSVersion + " ].\n\nEnsure Tekla [ " + skTSVersion + " ] model is open,if still issue's contact support team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    System.Environment.Exit(0);
                    //esskayjn = txtproject.Text.ToString();
                }
                else
                {
                    esskayjn = skTSLib.SKProjectNumber;
                    txtskproject.Text = esskayjn;
                    //get project information
                    Load_Project();

                    //setting for data grid
                    //dgdrglist.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    //dgdrglist.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                    //dgdrglist.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders);


                    //get windows login to avoid multiple factor authentication
                    String empid = skWinLib.username;

                    if (skWinLib.username.ToUpper() == "VISHWARAM")
                    {
                        empid = "vishwaramp";
                    }
                    else if (skWinLib.username.ToUpper() == "KESAVAN")
                    {
                        //                KESAVAN / SKS026
                        empid = "kesavan";
                        //lblrole.Text = "Checker";
                    }
                    else if (skWinLib.username.ToUpper() == "NIMAIN")
                    {
                        //                NIMAIN / SKS002
                        empid = "nimain";
                        //lblrole.Text = "Checker";
                    }
                    else if (skWinLib.username.ToUpper() == "SATHISH")
                    {
                        //                SATHISH / SKS246
                        empid = "Sathish_BNS";
                        //lblrole.Text = "Checker";
                    }


                    


                    id_emp = -1;

                    //load employee
                    Load_Employee();

                    //get  users.id from empmast
                    if (empmast.Count >= 1)
                    {
                        ArrayList tmp = getDBUserID(empid);
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                            id_emp = Convert.ToInt32(empsplit[0]);
                            lbluser.Text = empsplit[2].ToString();
                        }

                    }

                    if (id_emp == -1)
                    {
                        skWinLib.accesslog(skApplicationName, skApplicationVersion, "EsskayValidation", "Could not complete ESSKAY authenticaion.\n\n\nYour login SKS%%% " + skWinLib.username + "", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
                        MessageBox.Show("Could not complete ESSKAY authenticaion.\n\n\nYour login SKS%%% " + skWinLib.username + " either not added or incorrect loginname, if still issue's contact support team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        System.Windows.Forms.Application.ExitThread();
                    }
                    else
                    {
                        //set activity
                        SetMyActivity();

                        //Update user interface with last saved information in registry
                        //UpdateUserInterface_Fromlastsave();



                        //in future restore chklst items automatically



                        //lblsbar2.Text = "";
                        this.TopMost = true;


                        //hide ts report
                        //TabControl1.TabPages.Remove(tpgts);

                        //this.tpgqc.Parent = null;
                        //TabControl1.TabPages[1].Hide = true;
                        //(TabControl1.TabPages[1] as TabPage).Enabled = false;

                        //set tool tip
                        //settooltip();

                        //for debugging set lblsbar2 visible as true
                        if (skWinLib.username.ToUpper() == "VIJAYAKUMAR" || skWinLib.username.ToUpper() == "SKS360")
                        {
                            lblsbar2.Visible = true;
                        }
                        else
                        {
                            lblsbar2.Visible = false;
                        }
                        //TabControl1.TabPages.Remove(tpgqcmirror);

                        //get empname
                        ArrayList tmp = getDBUserID(skWinLib.username);
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });

                            //get index
                            int idx = cmbusername.FindStringExact(empsplit[2].ToString() + "|" + empsplit[1].ToString());
                            if (idx != -1)
                                cmbusername.SelectedIndex = idx;


                        }

                        //if count more than one set default index as zero
                        if (cmbmyactivity.Items.Count >= 1)
                            cmbmyactivity.SelectedIndex = 0;

                        //set the default location as model folder \plotfiles
                        txtpdflocation.Text  = skTSLib.ModelPath + "\\Plotfiles\\";

                    }
                }
            }


        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            //clear all row values
            dgdrglist.Rows.Clear();
        }


        private int bulk_insert(string bulkdata)
        {
            //idproject,drgmark,drgname,drgrev,drgtype,drgtitle1,activity,winlogin,ownerlogin,systemname,date,remark  
            string insert_statement = "INSERT INTO epmatdb.drawing(idproject,drgmark,drgname,drgrev,drgtype,drgtitle1,activity,winlogin,ownerlogin,systemname,date,remark) VALUES ";
            mycmd.CommandText = insert_statement + "(" + bulkdata + ")";
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            int insct = mycmd.ExecuteNonQuery();
            return insct;
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "Save", "", "Pl.Wait,Saving...", "btnSave_Click", lblsbar1, lblsbar1);
            string msg = string.Empty;
            DateTime chk_dt = DateTime.Now;
            DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
            long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

            //loop through for bulk insert
            string bulkdata = ")";
            foreach (DataGridViewRow myrow in dgdrglist.Rows)
            {
                //idproject,drgmark,drgname,drgrev,drgtype,drgtitle1,winlogin,ownerlogin,systemname,date,remark                
                string drgmark = string.Empty;
                string drgname = string.Empty;
                string drgrev = string.Empty;
                string drgtype = string.Empty;
                string drgtitle1 = string.Empty;
                string activity = string.Empty;                
                string ownerlogin = skWinLib.username;
                string remark = string.Empty;


                if (myrow.Cells["DrgMark"].Value != null)
                    drgmark = myrow.Cells["DrgMark"].Value.ToString();


                if (myrow.Cells["DrgName"].Value != null)
                    drgname = myrow.Cells["DrgName"].Value.ToString();

                if (myrow.Cells["DrgRev"].Value != null)
                    drgrev = myrow.Cells["DrgRev"].Value.ToString();

                if (myrow.Cells["DrgType"].Value != null)
                    drgtype = myrow.Cells["DrgType"].Value.ToString();

                if (myrow.Cells["drgTitle1"].Value != null)
                    drgtitle1 = myrow.Cells["drgTitle1"].Value.ToString();

                if (myrow.Cells["process"].Value != null)
                    activity = myrow.Cells["process"].Value.ToString();

                if (myrow.Cells["ownerlogin"].Value != null)
                    ownerlogin = myrow.Cells["ownerlogin"].Value.ToString();

                if (myrow.Cells["remark"].Value != null)
                    remark = myrow.Cells["remark"].Value.ToString();

                //idproject,drgmark,drgname,drgrev,drgtype,drgtitle1,activity,winlogin,ownerlogin,systemname,date,remark  
                bulkdata = id_project + ",'" + drgmark + "','" + drgname + "','" + drgrev + "','" + drgtype + "','" + drgtitle1 +  "','" + activity + "','" + skWinLib.username + "','" + ownerlogin + "','" + skWinLib.systemname + "'," + timestamp  + ",'" + remark + "') , (" + bulkdata;
            }
        
            if (bulkdata.Length >= 1)
            {
                bulkdata = bulkdata.Replace(") , ()", "");
                bulk_insert(bulkdata);
            }

            //clear all rows
            dgdrglist.Rows.Clear();

            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "Save", "", msg, "btnSave_Click", lblsbar1, lblsbar1, s_tm);
        }
        //private void btnsave_Click(object sender, EventArgs e)
        //{
        //    DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "Save", "", "Pl.Wait,Saving...", "btnSave_Click", lblsbar1, lblsbar1);
        //    string msg = string.Empty;
        //    DateTime chk_dt = DateTime.Now;
        //    DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
        //    long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

        //    //loop through for bulk insert
        //    string bulkdata = ")";
        //    foreach (DataGridViewRow myrow in dgdrglist.Rows)
        //    {
        //        //idproject,drgmark,drgname,drgrev,drgtype,idlogin,iduser,systemname,date
        //        string id = string.Empty;
        //        string mark = string.Empty;
        //        string DrgName = string.Empty;
        //        string rev = string.Empty;
        //        string drgtype = string.Empty;
        //        string pdfname = string.Empty;
        //        string pdfpath = string.Empty;

        //        string idlogin = skWinLib.username;

        //        if (myrow.Cells["id"].Value != null)
        //            id = myrow.Cells["id"].Value.ToString();

        //        if (myrow.Cells["DrgMark"].Value != null)
        //            mark = myrow.Cells["DrgMark"].Value.ToString();


        //        if (myrow.Cells["DrgName"].Value != null)
        //            DrgName = myrow.Cells["DrgName"].Value.ToString();

        //        if (myrow.Cells["DrgRev"].Value != null)
        //            rev = myrow.Cells["DrgRev"].Value.ToString();

        //        if (myrow.Cells["DrgType"].Value != null)
        //            drgtype = myrow.Cells["DrgType"].Value.ToString();

        //        if (myrow.Cells["pdfname"].Value != null)
        //            pdfname = myrow.Cells["pdfname"].Value.ToString();

        //        if (myrow.Cells["pdfpath"].Value != null)
        //            pdfpath = myrow.Cells["pdfpath"].Value.ToString();


        //        if (myrow.Cells["modempid"].Value != null)
        //            idlogin = myrow.Cells["modempid"].Value.ToString();



        //        //for bulk insert
        //        if (id == string.Empty)
        //        {
        //            //idproject,drgmark,drgname,pdfname,pdf,drgrev,drgtype,winlogin,iduser,systemname,date
        //            if (pdfpath.Length >= 1 && pdfname.Length >= 1)
        //            {
        //                string file = pdfpath + pdfname + ".pdf";
        //                bulkdata = id_project + ",'" + mark + "','" + DrgName + "','" + pdfname + "','" + System.IO.File.ReadAllBytes(file) + "','" + rev + "','" + drgtype + "','" + skWinLib.username + "','" + idlogin + "','" + skWinLib.systemname + "'," + timestamp + " ) , (" + bulkdata;
        //            }
        //            else
        //            {
        //                bulkdata = id_project + ",'" + mark + "','" + DrgName + "','" + pdfname + "','NULL','" + rev + "','" + drgtype + "','" + skWinLib.username + "','" + idlogin + "','" + skWinLib.systemname + "'," + timestamp + " ) , (" + bulkdata;
        //            }
        //            //else
        //            //    pdf = new byte[0];


        //            //////mysql
        //            ////mycmd.CommandText = "INSERT INTO epmatdb.pocpdf(file,name) VALUES(@file, @name)";
        //            ////mycmd.Parameters.Clear();
        //            ////mycmd.Parameters.Add("@file", MySqlDbType.MediumBlob);
        //            ///

        //        }


        //    }
        //    if (bulkdata.Length >= 1)
        //    {
        //        bulkdata = bulkdata.Replace(") , ()", "");
        //        bulk_insert(bulkdata);
        //    }

        //    skTSLib.setCompletion(skApplicationName, skApplicationVersion, "Save", "", msg, "btnSave_Click", lblsbar1, lblsbar1, s_tm);
        //}  






        private void Load_Employee()
        {
            ////fetch username,name,id,email
            ////sks360, vijayakumar, 297, vijayakumar@esskaystructures.com
            //mysql = "select username,name,id,email from epmatdb.users";
            //mycmd.CommandText = mysql;

            //myconn = new MySqlConnection(connString);
            //if (myconn.State == ConnectionState.Closed)
            //    myconn.Open();

            //mycmd.Connection = myconn;
            //myrdr = mycmd.ExecuteReader();
            //cmbusername.Items.Clear();
            //empmast.Clear();
            ////string cumlist = "";
            //while (myrdr.Read())
            //{

            //    cmbusername.Items.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString());
            //    empmast.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString() + "|" + myrdr[2].ToString() + "|" + myrdr[3].ToString());
            //    //cumlist = cumlist + "$" + myrdr[0].ToString() + "|" + myrdr[1].ToString();
            //}
            //myrdr.Close();
            ////set sorted based on UI setting
            ////cmbusername.Sorted();
            //empmast.Sort();
            ////string[] mylist =  cumlist.Split(new Char[] { '$' });
            ////cmbusername.AutoCompleteCustomSource.AddRange(mylist);

            //fetch  user information and store in empmast;
            empmast.Clear();
            mysql = "select users.id,users.username,users.name,users.email from epmatdb.users order by id";
            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {

                cmbusername.Items.Add(myrdr[2].ToString() + "|" + myrdr[1].ToString());
                empmast.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString() + "|" + myrdr[2].ToString() + "|" + myrdr[3].ToString());
            }
            myrdr.Close();
        }

        private void btnapply_Click(object sender, EventArgs e)
        {

        }

        private void lbluser_Click(object sender, EventArgs e)
        {

        }
       

        private void rbtnChecking_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnChecking.Checked == true)
                SetMyActivity();
        }

        private void rbtnediting_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtnediting.Checked == true)
                SetMyActivity();
        }
        private void dgdrglist_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

                lblsbar2.Text = "dgdrglist_RowsAdded";
                int rowindex = e.RowIndex;
                if (e.RowIndex >= 0)
                    MyActivitySetting(rowindex);

            dgdrglist.Rows[rowindex].HeaderCell.Value = (rowindex + 1).ToString();
                //lblsbar2.Text = "";
            
        }


        //check mysql connection is feasible WFH / WFO
        private bool IsCheckconnection()
        {


            myconn.ConnectionString = connString;
            try
            {
                myconn.Open();
                if (myconn.State == ConnectionState.Closed)
                    return false;
                else
                    return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source);
                return false;
            }
        }

        //get db user id based on string login id (ie. SKS360 >> Vijayakumar)
        private ArrayList getDBUserID(string login_id)
        {
            ArrayList myreturn = new ArrayList();
            //check if empmast is not null
            for (int i = 0; i < empmast.Count; i++)
            {
                string[] empsplit = empmast[i].ToString().Split(new Char[] { '|' });
                if (empsplit.Length >= 1)
                {
                    if (empsplit[1].ToString().ToUpper() == login_id.ToUpper())
                    {
                        myreturn.Add(empmast[i]);
                        break;
                    }
                }
            }
            return myreturn;
        }

        private void MyActivitySetting(int rowindex)
        {
            bool chk = false;
            if (dgdrglist.RowCount >= 1)
            {
                skWinLib.DataGridView_Setting_Before(dgdrglist);

                for (int i = 0; i < dgdrglist.Columns.Count; i++)
                {
                    DataGridViewRow row = this.dgdrglist.Rows[rowindex];
                    if (dgdrglist.Columns[i].Name.ToUpper() == "process".ToUpper())
                    {
                        DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["process"]);
                        cell.DataSource = myactivity;
                        chk = true;
                    }

                    if ((chk == true))
                        break;
                }


                skWinLib.DataGridView_Setting_After(dgdrglist);
            }

        }


        private void Load_Project()
        {
           
            //Get SK Project Number from Project Property Now and in future cmbproject will be locked once confirmed MD.
            try
            {
                //MyModel = skTSLib.ModelAccess.ConnectToModel();
                //if (MyModel != null)
                {
                    //esskayjn = MyModel.GetProjectInfo().Info1;
                    //esskayjn = skTSLib.SKProjectNumber;
                    if (esskayjn.Length >= 4)
                    {
                        //only one project information
                        mysql = "select id,esskayjn,name,clientjn from epmatdb.project where esskayjn = '" + esskayjn + "'";
                        myconn = new MySqlConnection(connString);
                        if (myconn.State == ConnectionState.Closed)
                            myconn.Open();
                        mycmd.CommandText = mysql;
                        mycmd.Connection = myconn;
                        myrdr = mycmd.ExecuteReader();
                        while (myrdr.Read())
                        {
                            id_project = Convert.ToInt32(myrdr[0].ToString());
                            projectname = myrdr[2].ToString();
                            clientjn = myrdr[3].ToString();
                        }
                        myrdr.Close();
                        if (id_project != -1)
                        {
                            mysql = "SELECT id,name FROM epmatdb.client where id = (SELECT contact_id FROM epmatdb.project_contact where project_id = " + id_project + ")";
                            myconn = new MySqlConnection(connString);
                            if (myconn.State == ConnectionState.Closed)
                                myconn.Open();
                            mycmd.CommandText = mysql;
                            mycmd.Connection = myconn;
                            myrdr = mycmd.ExecuteReader();                           
                            while (myrdr.Read())
                            {

                                client = myrdr[1].ToString();
                            }
                            myrdr.Close();

                            this.Text = skApplicationName + " [" + skApplicationVersion + "]" + "      " + esskayjn + " / " + projectname + " / " + client;
                            
                        }
                        else
                        {
                            MessageBox.Show("SK Project Number is not a valid project number.\nCheck the your input and try again.\n\n\nAborting Now!!!", "Esskay Project Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //System.Windows.Forms.Application.ExitThread();
                            System.Environment.Exit(0);
                        }

                    }
                    else
                    {
                        MessageBox.Show("SK Project Number not found Model Project Properties.\nUpdate SK Project Number and try again.\n\n\nAborting Now!!!", "Esskay Project Number Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //System.Windows.Forms.Application.ExitThread();
                        System.Environment.Exit(0);
                    }

                }
                //else
                //{
                //    MessageBox.Show("Either Tekla not open or Version is different.\n\nPlease recheck and contact support team if still have issue's.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //    //System.Environment.Exit(0);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Initail error: " + ex.Message);
            }
        }
        private void SetMyActivity()
        {
            myactivity.Clear();
            cmbmyactivity.Items.Clear();
            dgdrglist.Rows.Clear();
            if (rbtnediting.Checked == true)
            {
                //myactivity.Add("Fresh Editing");
                myactivity.Add("Editing");
                myactivity.Add("Back Drafting");
                //myactivity.Add("Mark Changed");
                myactivity.Add("Deleted");

            }
            else
            {
                //myactivity.Add("Fresh Checking");
                myactivity.Add("Checking");
                myactivity.Add("Back Checking");
                myactivity.Add("Cancelled");
            }
            myactivity.Add("IFA");
            myactivity.Add("IFF");
            myactivity.Add("Rework");
            myactivity.Add("Rework-Others");

            for (int i = 0; i < myactivity.Count; i++)
                cmbmyactivity.Items.Add(myactivity[i].ToString());

            //if count more than one set default index as zero
            if (cmbmyactivity.Items.Count >= 1)
                cmbmyactivity.SelectedIndex = 0;
        }

        public ArrayList GetDrawingDetail()
        {
            ArrayList drgdata = new ArrayList();
            lblsbar1.Text = "Please Wait,processing [DWG]...";
            // Check if we have connection to Tekla Structures drawings.
            if (dh.GetConnectionStatus())
            {
                TSM.Operations.Operation.DisplayPrompt("Selecting DWG Files.");


                //user selection type
                string useroption = "Select";
                //if (rbtnChecking.Checked == true)
                //    useroption = "All";

                //select the drawing based on user option
                TSD.DrawingEnumerator selectedDrawings;
                if (useroption == "Select")
                    selectedDrawings = dh.GetDrawingSelector().GetSelected();
                else
                    selectedDrawings = dh.GetDrawings();

                while (selectedDrawings.MoveNext())
                {
                    Tekla.Structures.Drawing.Drawing currentDrawing = selectedDrawings.Current;
                    string drawingType = currentDrawing.GetType().ToString();

                    //dh.SetActiveDrawing(currentDrawing, false); 
                    try
                    {
                        System.Reflection.PropertyInfo pi = selectedDrawings.Current.GetType().GetProperty("Identifier", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                        object value = pi.GetValue(selectedDrawings.Current, null);
                        Identifier Identifier = (Identifier)value;
                        TSM.Beam fakebeam = new TSM.Beam();
                        fakebeam.Identifier = Identifier;
                        string drgrev = "";
                        bool rev = fakebeam.GetReportProperty("REVISION.MARK", ref drgrev);
                        int revisionNumber = -1;
                        fakebeam.GetReportProperty("REVISION.NUMBER", ref revisionNumber);
                        drgdata.Add("DWG|" + drawingType + "|" + currentDrawing.Mark + "|" + currentDrawing.Name + "|" + drgrev + "|" + currentDrawing.Title1 + "|" + currentDrawing.Title2 + "|" + currentDrawing.Title3 + "|" + currentDrawing.GetSheet().ToString() + "|" + currentDrawing.Layout.ToString() + "|" + currentDrawing.CreationDate.ToString() + "|" + currentDrawing.ModificationDate.ToString());
                    }
                    catch (Exception ex)
                    {
                        TSM.Operations.Operation.DisplayPrompt("Error: " + ex.Message);
                    }


                }
                TSM.Operations.Operation.DisplayPrompt("DWG Files Exported.");
            }
            else
            {
                MessageBox.Show("Tekla Application either not running or no drawing(s) created.", "msgCaption", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            lblsbar1.Text = "Ready.";
            return drgdata;
        }

        private void btnuser_Click(object sender, EventArgs e)
        {
            string empnameid = cmbusername.Text;
            if (empnameid.Length  >= 1)
            {
                string[] mydatasplt = empnameid.Split(new Char[] { '|' });
                string rowempname = mydatasplt[0].ToString();
                string rowmodempid = mydatasplt[1].ToString();

                foreach (DataGridViewRow myrow in dgdrglist.SelectedRows)
                {  
                        myrow.Cells["empname"].Value = rowempname;
                        myrow.Cells["ownerlogin"].Value = rowmodempid;

                }
            }

        }

        private void btnprocess_Click(object sender, EventArgs e)
        {
            string myactivity = cmbmyactivity.Text;
            if (myactivity.Length  >= 1)
            {
                foreach (DataGridViewRow myrow in dgdrglist.SelectedRows)
                {
                    myrow.Cells["process"].Value = myactivity;
                }
            }
        }

        private void btnupload_Click(object sender, EventArgs e)
        {
            bool flag = false;
            bool default_pdf_path_flag = false;
            bool pdfnullflag = false;
            //check if selected row is greater than zero
            if (dgdrglist.SelectedRows.Count >=1)
            {
                //default path 
                //string default_pdf_path = skTSLib.ModelPath + "\\Plotfiles\\";
                string default_pdf_path = txtpdflocation.Text;

                //check last word is back slash or not add this
                char[] mychar = default_pdf_path.ToCharArray();
                if (mychar[mychar.Length-1] != '\\')
                    default_pdf_path = default_pdf_path + "\\";

                //array list for pdf
                string[] pdfList = null;

                //check if default_pdf_path folder exists and pdf file available
                if (System.IO.Directory.Exists(default_pdf_path) == true)
                {
                    default_pdf_path_flag = true;
                    pdfList = System.IO.Directory.GetFiles(default_pdf_path);
                }

ReLoopWithoutDefaultPath:;
                //peform based on pdflist    
                if (pdfList.Length == 0)
                {
                    default_pdf_path_flag = false;
                    default_pdf_path = skTSLib.ModelPath + "\\";
                    //check if default_pdf_path folder exists if not prompts for location
                    //if (System.IO.Directory.Exists(default_pdf_path) == false)
                    //{
                        OpenFileDialog myopenFileDialog = new OpenFileDialog();
                        myopenFileDialog.Filter = "PDF File|*.pdf";
                        myopenFileDialog.Title = "Select a PDF File";
                        myopenFileDialog.Multiselect = true;
                        myopenFileDialog.InitialDirectory = default_pdf_path;
                        if (myopenFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            //get default path
                            default_pdf_path = System.IO.Path.GetDirectoryName(myopenFileDialog.FileName) + "\\";
                            pdfList = myopenFileDialog.FileNames;
                        }
                    //}
                }


                //peform based on pdflist             
                if (pdfList.Length >= 1)
                {
                    //attach pdf file for that row 
                    foreach (DataGridViewRow myrow in dgdrglist.SelectedRows)
                    {

                        string drgmark = myrow.Cells["DrgMark"].Value.ToString();
                        string drgname = myrow.Cells["DrgName"].Value.ToString();
                        string drgrev = myrow.Cells["DrgRev"].Value.ToString();
                        string drgtype = myrow.Cells["DrgType"].Value.ToString();
                        if (drgtype.ToUpper() == "G")
                            drgmark = drgname;
                        string pdf = string.Empty;
                        if (myrow.Cells["pdfname"].Value == null)
                        {
                            if (dgdrglist.SelectedRows.Count == 1)
                            {
                                
                                foreach (string pdfname in pdfList)
                                {
                                    pdf = pdfname;

                                    //get the pdf path for storage
                                    pdf = pdf.Replace(default_pdf_path, "");

                                    //process pdf to remove this from file name
                                    pdf = pdf.Replace(".pdf", "");
                                    pdf = pdf.Replace(".PDF", "");
                                    pdf = pdf.Replace(".Pdf", "");

                                    //Example
                                    //B1012 = B1012
                                    //B1012_R0 = B1012_R0
                                    //B1012_0 = B1012_0
                                    //B1012-R0 = B1012-R0
                                    //B1012-0 = B1012-0
                                    //check direct pdf + rev and drgmark
                                    if (pdf.ToUpper() == drgmark.ToUpper() || (pdf.ToUpper() == drgmark.ToUpper() + "_R" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "_" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "-R" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "-" + drgrev.ToUpper()))
                                    {
                                        myrow.Cells["pdfname"].Value = pdf;
                                        myrow.Cells["pdfpath"].Value = default_pdf_path;
                                        flag = true;
                                        break;
                                    }
                                }
                                if (flag == false)
                                {
                                    OpenFileDialog myopenFileDialog = new OpenFileDialog();
                                    myopenFileDialog.Filter = "PDF File|*.pdf";
                                    myopenFileDialog.Title = "Select a PDF File";
                                    myopenFileDialog.Multiselect = true;
                                    myopenFileDialog.InitialDirectory = default_pdf_path;
                                    if (myopenFileDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        //get default path
                                        default_pdf_path = System.IO.Path.GetDirectoryName(myopenFileDialog.FileName) + "\\";
                                        pdfList = myopenFileDialog.FileNames;
                                    }
                                    if (pdfList.Length >= 1)
                                    {
                                        pdf = pdfList[0].ToString();

                                        //get the pdf path for storage
                                        pdf = pdf.Replace(default_pdf_path, "");

                                        //process pdf to remove this from file name
                                        pdf = pdf.Replace(".pdf", "");
                                        pdf = pdf.Replace(".PDF", "");
                                        pdf = pdf.Replace(".Pdf", "");

                                        myrow.Cells["pdfname"].Value = pdf;
                                        myrow.Cells["pdfpath"].Value = default_pdf_path;
                                        flag = true;
                                        break;
                                    }
                                }


                            }
                            else
                            {
                                pdfnullflag = true;
                                //get the pdf of each file name
                                foreach (string pdfname in pdfList)
                                {
                                    pdf = pdfname;

                                    //get the pdf path for storage
                                    pdf = pdf.Replace(default_pdf_path, "");

                                    //process pdf to remove this from file name
                                    pdf = pdf.Replace(".pdf", "");
                                    pdf = pdf.Replace(".PDF", "");
                                    pdf = pdf.Replace(".Pdf", "");

                                    //Example
                                    //B1012 = B1012
                                    //B1012_R0 = B1012_R0
                                    //B1012_0 = B1012_0
                                    //B1012-R0 = B1012-R0
                                    //B1012-0 = B1012-0
                                    //check direct pdf + rev and drgmark
                                    if (pdf.ToUpper() == drgmark.ToUpper() || (pdf.ToUpper() == drgmark.ToUpper() + "_R" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "_" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "-R" + drgrev.ToUpper()) || (pdf.ToUpper() == drgmark.ToUpper() + "-" + drgrev.ToUpper()))
                                    {
                                        myrow.Cells["pdfname"].Value = pdf;
                                        myrow.Cells["pdfpath"].Value = default_pdf_path;
                                        flag = true;
                                    }
                                }
                            }

                            
                        }

                    }
                }

                if (default_pdf_path_flag == true && flag == false && pdfnullflag == true)
                {
                    default_pdf_path_flag = false;
                    pdfList = new string[] { };
                    goto ReLoopWithoutDefaultPath;
                }    
            
            }
        }

        private void btndnload_Click(object sender, EventArgs e)
        {
            //tolerance set as 1/16
            double tolerance = (1 / 32.0) * 25.4;
            int ct = 1;
            ArrayList ObjectsToSelect = new ArrayList();
            ArrayList ignoredProfileType = new ArrayList();
            string signoredProfileType = string.Empty;

            TransformationPlane CurrentTP = _Model.GetWorkPlaneHandler().GetCurrentTransformationPlane();

            TSM.ModelObjectEnumerator MyEnum = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
            int totselct = MyEnum.GetSize();
            int FaceIndex = 0;
            while (MyEnum.MoveNext())
            {
                Tekla.Structures.Model.Beam myBeam = MyEnum.Current as Tekla.Structures.Model.Beam;
                if (myBeam != null)
                {
                    string myProfileType = string.Empty ;
                    myBeam.GetReportProperty("PROFILE_TYPE", ref myProfileType);
                    myProfileType = myProfileType.ToUpper().Trim();
                    if ((myProfileType == "I") || (myProfileType == "C") || (myProfileType == "O") || (myProfileType == "B") || (myProfileType == "L") || (myProfileType == "M")  || (myProfileType == "RO") || (myProfileType == "RU"))
                    {
                        //Set the transformation plane to use the beam's coordinate system in order to get the beam's extremes in the local coordinate system
                        _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(myBeam.GetCoordinateSystem()));

                        //Update the beam's extremes to the new transformation plane
                        myBeam.Select();

                        //get the profile depth and hieght for I, C and O Shape
                        //double height = 0.0;
                        //myBeam.GetReportProperty("HEIGHT", ref height);
                        //double width = 0.0;
                        //myBeam.GetReportProperty("WIDTH", ref width);

                        Tekla.Structures.Geometry3d.Point LocalStartPoint = myBeam.StartPoint;
                        Tekla.Structures.Geometry3d.Point LocalEndPoint = myBeam.EndPoint;

                        //LocalStartPoint.Z = LocalStartPoint.Z - 1000;
                        //LocalEndPoint.Z = LocalEndPoint.Z - 1000;
                        //GraphicsDrawer.DrawText(myBeam.StartPoint, FormatPointCoordinates(myBeam.StartPoint, "Start"), TextColor);
                        //GraphicsDrawer.DrawText(myBeam.EndPoint, FormatPointCoordinates(myBeam.EndPoint, "End"), TextColor);

                        if (myBeam != null)
                        {
                            ArrayList MyList = new ArrayList();
                            ArrayList MyFaceNormalList = new ArrayList();
                            Solid Solid = myBeam.GetSolid();
                            IEnumerator FaceEnum = Solid.GetFaceEnumerator();

                            while (FaceEnum.MoveNext())
                            {
                                Face MyFace = FaceEnum.Current as Face;
                                if (MyFace != null)
                                {

                                    MyFaceNormalList.Add(MyFace.Normal);
                                    LoopEnumerator MyLoopEnum = MyFace.GetLoopEnumerator();
                                    while (MyLoopEnum.MoveNext())
                                    {
                                        FaceIndex++;
                                        Tekla.Structures.Geometry3d.Point PastVertex = null;
                                        Loop MyLoop = MyLoopEnum.Current as Loop;
                                        if (MyLoop != null)
                                        {
                                            VertexEnumerator MyVertexEnum = MyLoop.GetVertexEnumerator() as VertexEnumerator;
                                            while (MyVertexEnum.MoveNext())
                                            {
                                                Tekla.Structures.Geometry3d.Point MyVertex = MyVertexEnum.Current as Tekla.Structures.Geometry3d.Point;
                                                //GraphicsDrawer.DrawText(MyVertex, FormatPointCoordinates(MyVertex, ct.ToString()), TextColor);
                                                ct++;
                                                if (PastVertex != null)
                                                {
                                                    //get difference between x & y
                                                    double Difference_X = Math.Abs(MyVertex.X - PastVertex.X) - tolerance;
                                                    double Difference_Y = Math.Abs(MyVertex.Y - PastVertex.Y) - tolerance;
                                                    double Difference_Z = Math.Abs(MyVertex.Z - PastVertex.Z) - tolerance;
                                                    if (Difference_X > 0.0 && Difference_Y > 0.0)
                                                    {

                                                        if (!MyList.Contains(MyVertex))
                                                        {
                                                            MyList.Add(MyVertex);
                                                            GraphicsDrawer.DrawText(PastVertex, FormatPointCoordinates(PastVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            GraphicsDrawer.DrawText(MyVertex, FormatPointCoordinates(MyVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            //MessageBox.Show("Bevel Cut @ XY");
                                                            TSM.Operations.Operation.DisplayPrompt("Bevel Cut @ XY");
                                                            goto ExitBevelCut;
                                                        }
                                                    }
                                                    // get difference between Y & Z
                                                    if (Difference_Y > 0.0 && Difference_Z > 0.0)
                                                    {

                                                        if (!MyList.Contains(MyVertex))
                                                        {
                                                            MyList.Add(MyVertex);
                                                            GraphicsDrawer.DrawText(PastVertex, FormatPointCoordinates(PastVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            GraphicsDrawer.DrawText(MyVertex, FormatPointCoordinates(MyVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            TSM.Operations.Operation.DisplayPrompt("Bevel Cut @ YZ");
                                                            goto ExitBevelCut;
                                                        }
                                                    }
                                                    // get difference between x & z
                                                    if (Difference_X > 0.0 && Difference_Z > 0.0)
                                                    {

                                                        if (!MyList.Contains(MyVertex))
                                                        {
                                                            MyList.Add(MyVertex);
                                                            FaceIndex++;
                                                            GraphicsDrawer.DrawText(PastVertex, FormatPointCoordinates(PastVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            GraphicsDrawer.DrawText(MyVertex, FormatPointCoordinates(MyVertex, FaceIndex, MyFace.Normal), TextColor);
                                                            TSM.Operations.Operation.DisplayPrompt("Bevel Cut @ XZ");
                                                            goto ExitBevelCut;
                                                        }
                                                    }
                                                }
                                                PastVertex = MyVertex;

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    
                    }
                    else
                    {
                        if (!ignoredProfileType.Contains(myProfileType))
                        {
                            ignoredProfileType.Add(myProfileType);
                            signoredProfileType = myProfileType + "\n" + signoredProfileType;
                        }

                    }
                }               

ExitBevelCut:;
            }
            _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);
            TSM.Operations.Operation.DisplayPrompt("Ready");

        }

        private void ShowExtremesInOtherObjectCoordinates(Tekla.Structures.Model.ModelObject ReferenceObject, Beam Beam)
        {
            //Set the transformation plane to use the beam's coordinate system in order to get the beam's extremes in the local coordinate system
            TransformationPlane CurrentTP = _Model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(Beam.GetCoordinateSystem()));

            //Update the beam's extremes to the new transformation plane
            Beam.Select();
            Tekla.Structures.Geometry3d.Point LocalStartPoint = Beam.StartPoint;
            Tekla.Structures.Geometry3d.Point LocalEndPoint = Beam.EndPoint;

            //Get the beam's extremes in the reference object's coordinates
            CoordinateSystem ReferenceObjectCS = ReferenceObject.GetCoordinateSystem();
            Matrix TransformationMatrix = MatrixFactory.ByCoordinateSystems(Beam.GetCoordinateSystem(), ReferenceObjectCS);



            //Transform the extreme points to the new coordinate system
            Tekla.Structures.Geometry3d.Point BeamStartPoint = TransformationMatrix.Transform(LocalStartPoint);
            Tekla.Structures.Geometry3d.Point BeamEndPoint = TransformationMatrix.Transform(LocalEndPoint);

            _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);

            //Transform the points where to show the texts to current work plane coordinate system
            Matrix TransformationToCurrent = MatrixFactory.FromCoordinateSystem(ReferenceObjectCS);
            Tekla.Structures.Geometry3d.Point BeamStartPointInCurrent = TransformationToCurrent.Transform(BeamStartPoint);
            Tekla.Structures.Geometry3d.Point BeamEndPointInCurrent = TransformationToCurrent.Transform(BeamEndPoint);

            //Display results
            DrawCoordinateSytem(ReferenceObjectCS);
            GraphicsDrawer.DrawText(BeamStartPointInCurrent, FormatPointCoordinates(BeamStartPoint), TextColor);
            GraphicsDrawer.DrawText(BeamEndPointInCurrent, FormatPointCoordinates(BeamEndPoint), TextColor);
        }

        private void ShowExtremesInOtherObjectCoordinates(Beam Beam)
        {
            //Set the transformation plane to use the beam's coordinate system in order to get the beam's extremes in the local coordinate system
            TransformationPlane CurrentTP = _Model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(Beam.GetCoordinateSystem()));

            //Update the beam's extremes to the new transformation plane
            Beam.Select();
            Tekla.Structures.Geometry3d.Point LocalStartPoint = Beam.StartPoint;
            Tekla.Structures.Geometry3d.Point LocalEndPoint = Beam.EndPoint;

            //Get the beam's extremes in the reference object's coordinates

            Vector BeamAxisX = new Vector(Beam.StartPoint.X + 100, Beam.StartPoint.Y, Beam.StartPoint.Z);
            Vector BeamAxisY = new Vector(Beam.StartPoint.X, Beam.StartPoint.Y + 100, Beam.StartPoint.Z);
            CoordinateSystem ReferenceObjectCS = new CoordinateSystem(Beam.StartPoint, BeamAxisX, BeamAxisY);
            Matrix TransformationMatrix = MatrixFactory.ByCoordinateSystems(Beam.GetCoordinateSystem(), ReferenceObjectCS);

            //Transform the extreme points to the new coordinate system
            Tekla.Structures.Geometry3d.Point BeamStartPoint = TransformationMatrix.Transform(LocalStartPoint);
            Tekla.Structures.Geometry3d.Point BeamEndPoint = TransformationMatrix.Transform(LocalEndPoint);

            _Model.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);

            //Transform the points where to show the texts to current work plane coordinate system
            Matrix TransformationToCurrent = MatrixFactory.FromCoordinateSystem(ReferenceObjectCS);
            Tekla.Structures.Geometry3d.Point BeamStartPointInCurrent = TransformationToCurrent.Transform(BeamStartPoint);
            Tekla.Structures.Geometry3d.Point BeamEndPointInCurrent = TransformationToCurrent.Transform(BeamEndPoint);

            //Display results
            DrawCoordinateSytem(ReferenceObjectCS);
            GraphicsDrawer.DrawText(BeamStartPointInCurrent, FormatPointCoordinates(BeamStartPoint), TextColor);
            GraphicsDrawer.DrawText(BeamEndPointInCurrent, FormatPointCoordinates(BeamEndPoint), TextColor);
        }

        //Draws the coordinate system in which the values are shown
        private static void DrawCoordinateSytem(CoordinateSystem CoordinateSystem)
        {
            DrawVector(CoordinateSystem.Origin, CoordinateSystem.AxisX, "X");
            DrawVector(CoordinateSystem.Origin, CoordinateSystem.AxisY, "Y");
        }

        //Draws the vector of the coordinate system
        private static void DrawVector(Tekla.Structures.Geometry3d.Point StartPoint, Vector Vector, string Text)
        {
            Tekla.Structures.Model.UI.Color Color = new Tekla.Structures.Model.UI.Color(0, 1, 1);
            const double Radians = 0.43;

            Vector = Vector.GetNormal();
            Vector Arrow01 = new Vector(Vector);

            Vector.Normalize(500);
            Tekla.Structures.Geometry3d.Point EndPoint = new Tekla.Structures.Geometry3d.Point(StartPoint);
            EndPoint.Translate(Vector.X, Vector.Y, Vector.Z);
            GraphicsDrawer.DrawLineSegment(StartPoint, EndPoint, Color);

            GraphicsDrawer.DrawText(EndPoint, Text, Color);

            Arrow01.Normalize(-100);
            Vector Arrow = ArrowVector(Arrow01, Radians);

            Tekla.Structures.Geometry3d.Point ArrowExtreme = new Tekla.Structures.Geometry3d.Point(EndPoint);
            ArrowExtreme.Translate(Arrow.X, Arrow.Y, Arrow.Z);
            GraphicsDrawer.DrawLineSegment(EndPoint, ArrowExtreme, Color);

            Arrow = ArrowVector(Arrow01, -Radians);

            ArrowExtreme = new Tekla.Structures.Geometry3d.Point(EndPoint);
            ArrowExtreme.Translate(Arrow.X, Arrow.Y, Arrow.Z);
            GraphicsDrawer.DrawLineSegment(EndPoint, ArrowExtreme, Color);
        }

        //Draws the arrows of the vectors
        private static Vector ArrowVector(Vector Vector, double Radians)
        {
            double X, Y, Z;

            if (Vector.X == 0 && Vector.Y == 0)
            {
                X = Vector.X;
                Y = (Vector.Y * Math.Cos(Radians)) - (Vector.Z * Math.Sin(Radians));
                Z = (Vector.Y * Math.Sin(Radians)) + (Vector.Z * Math.Cos(Radians));
            }
            else
            {
                X = (Vector.X * Math.Cos(Radians)) - (Vector.Y * Math.Sin(Radians));
                Y = (Vector.X * Math.Sin(Radians)) + (Vector.Y * Math.Cos(Radians));
                Z = Vector.Z;
            }

            return new Vector(X, Y, Z);
        }

        //Shows the point coordinates with only two decimals
        private static string FormatPointCoordinates(Tekla.Structures.Geometry3d.Point Point)
        {
            string Output = String.Empty;

            Output = "(" + Point.X.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Y.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Z.ToString("0.00", CultureInfo.InvariantCulture) + ")";

            return Output;
        }

        private static string FormatPointCoordinates(Tekla.Structures.Geometry3d.Point Point, string Output)
        {

            Output = Output + " (" + Point.X.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Y.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Z.ToString("0.00", CultureInfo.InvariantCulture) + ")";

            return Output;
        }
        private static string FormatPointCoordinates(Tekla.Structures.Geometry3d.Point Point, int index, Vector myface )
        {
            string Output = String.Empty;

            Output = "   " + index.ToString() + " (" + myface.X.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     myface.Y.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     myface.Z.ToString("0.00", CultureInfo.InvariantCulture) + ")" + " \n(" + Point.X.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Y.ToString("0.00", CultureInfo.InvariantCulture) + ", " +
                     Point.Z.ToString("0.00", CultureInfo.InvariantCulture) + ")";

            return Output;
        }

        private void lblsbar1_Click(object sender, EventArgs e)
        {
            
        }

        private void txtskproject_Leave(object sender, EventArgs e)
        {
            //SK Project number is greater than zero then get the project details
            if (txtskproject.Text.Length >=1 )
            {
                if (esskayjn != txtskproject.Text)
                {
                    //set esskay jobnumber
                    esskayjn = txtskproject.Text;
                    //get the sk project number and clear all data
                    Load_Project();

                }
            }


        }

        private void txtskproject_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtskproject_KeyDown(object sender, KeyEventArgs e)
        {
            //if enter key is hit then assume user need this esskay job number
            Console.WriteLine(e.KeyCode.ToString().ToUpper().Trim());
            if (e.KeyCode.ToString().ToUpper().Trim() == "RETURN")
            {
                txtskproject_Leave(sender, e);
            }   
        }

        private void dgdrglist_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgdrglist_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //check whether user CellMouseDoubleClick on hyperlink
            if (e.ColumnIndex == 2)
            {
                if ((this.dgdrglist.Rows[e.RowIndex].Cells["pdfname"].Value != null) && (this.dgdrglist.Rows[e.RowIndex].Cells["pdfpath"].Value != null))
                {
                    string pdfname = this.dgdrglist.Rows[e.RowIndex].Cells["pdfname"].Value.ToString();
                    string pdfpath = this.dgdrglist.Rows[e.RowIndex].Cells["pdfpath"].Value.ToString();
                    string pdf = pdfpath + pdfname + ".pdf";
                    System.Diagnostics.Process.Start(pdf);
                }
            }

        }

        private void btnget_Click(object sender, EventArgs e)
        {
            //fetch the data from database
            //only one project information
            mysql = "select id,esskayjn,name,clientjn from epmatdb.project where esskayjn = '" + esskayjn + "'";
            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {
                id_project = Convert.ToInt32(myrdr[0].ToString());
                projectname = myrdr[2].ToString();
                clientjn = myrdr[3].ToString();
            }
        }
    }
}
