//Omm Muruga 01Dec2021
// ###########################################################################################
// ### Name          : Pipe Support Modelling Tool
// ### Version       : Beta for V2020.0
// ###               : Visual C# / Visual Studio 2019
// ### Created       : December 2021
// ### Modified      : 
// ### Author        : Vijayakumar
// ### Released      : 
// ### Comment       :
// ### Company       : Esskay design and structures pvt. ltd.
// ### Description   : Tool model pipe support as per business requirement
// ###                 
// ###                 
// ###########################################################################################
//

//// Version 5.x
//using Microsoft.Graph.Core;

using myExcel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using XLApp = Microsoft.Office.Interop.Excel.Application;
using Border = Microsoft.Office.Interop.Excel.Border;
using Range = Microsoft.Office.Interop.Excel.Range;
using XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight;
using XlLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle;
using XlConstants = Microsoft.Office.Interop.Excel.Constants;
using Microsoft.VisualBasic;//add reference


using System.Runtime.InteropServices;
using System.Security;
using System.Globalization;


//using System.Data.SqlClient;
using MySql.Data.MySqlClient;

using System;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
// Tekla Structures namespaces
using Tekla.Structures;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
//using IfcGuid;

//using TSG = Tekla.Structures.Geometry3d;
//using Tekla.Structures.Geometry3d;
//using Tekla.Structures.Solid;
using Tekla.Structures.Model.UI;
using TSMUI = Tekla.Structures.Model.UI;
//using T3D = Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.Operations;
//using Tekla.Structures.Drawing;
//using TSD = Tekla.Structures.Drawing;







namespace QMS
{
    

    public enum SKProcess 
    {

        Modelling = 0,

        Modelling_Completed = 10,

        Modelling_PM_Approved = 20,

        ModelChecking_Completed = 30,

        ModelChecking_PM_Approved = 40,

        ModelBackDraft_Completed = 50,

        ModelBackDraft_PM_Approved = 60,

        //ModelChecking_Conflict_Completed = 70,

        ModelBackCheck_Completed = 70,

        ModelBackCheck_PM_Approved = 80,

        ModelSeverity_Completed = 85,

        ModelSeverity_PM_Approved = 90,

        Modelling_QC_Completed = 100,

        Drawing_Completed = 110,

        Drawing_PM_Approved = 120,

        DrawingChecking_Completed = 130,

        DrawingChecking_PM_Approved = 140,

        DrawingBackDraft_Completed = 150,

        DrawingBackDraft_PM_Approved = 160,

        DrawingChecking_Conflict_Completed = 170,

        DrawingBackCheck_Completed = 180,

        DrawingBackCheck_PM_Approved = 190,

        Drawing_QC_Completed = 200,

        QA = 300,

        Held = 410,

        Cancelled = 420,

        DrawingHeld = 460,

        DrawingCancelled = 470,

        Completed = 500
    }

    public partial class frmQC : Form
    {
        //public string skApplicationName = "NCM";
        //public string skApplicationVersion = "2205.21";
        public string skApplicationName = "QC";
        public string skApplicationVersion = "2206.27";
        //Trichy Team Demo
        //Autoemail triggerred when user creates a new report
        //Excel export bug now fixed - unsure of this issue
        public string skTSVersion = "2022";
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

        public int SK_id_process = 0;
        public int SK_id_report = 0;

        public string misentry = string.Empty;
        public string qcentry = string.Empty;

        //Tekla additional column these additional column values are not stored in database except type. Additional column can be added but additional column should be in object.inp and should fetch output in report template.
        public string TS_Extra_Column = "Assembly_Pos,Phase.Name,Name,Profile,Material,Finish,Assembly_Prefix,Part_Prefix,User_field_1,User_field_2,User_field_3,User_field_4";

        //During fetch, all data's are stored in this arraylist which will be used to compare before saving and allow only changed values to be updated or inserted into database.
        public ArrayList sk_db_performance = new ArrayList();
        public const string sk_seperator = "|@|";
        //public const string sk_seperator = char[160];
        public const string sk_nocomment = "No Comment";



        //public const string sk_replace = "No Comment";

        public ArrayList tvw_dash_data = new ArrayList();
        //Treeview setting
        TreeNode tvwroot;
        TreeNode tvwproject;
        TreeNode tvwnew;
        TreeNode tvwchecking;
        TreeNode tvwbackdraft;
        TreeNode tvwbackcheck;
        //TreeNode tvwconflict;
        TreeNode tvwseverity;
        TreeNode tvwqa;
        TreeNode tvwheld;
        TreeNode tvwcancelled;
        TreeNode tvwcompleted;




        //DataGridViewComboBoxCell cboCol = (DataGridViewComboBoxCell)d.CurrentCell;
        DataGridViewComboBoxCell dgObservation = new DataGridViewComboBoxCell();

        ArrayList qmsCategory = new ArrayList();
        ArrayList qmsSeverity = new ArrayList();


        //working fine for LAN and VPN
        //public const string connString = "server=192.168.1.62;Port=3306;user id = esskay; password=esskay1505;database=epmatdb";
        //public const string connString = "server=(address=192.168.1.62:3306,priority=10),(address=122.165.187.180:8006,priority=60),(address=103.92.203.198:8006,priority=100);user id = esskay; password=esskay1505;database=epmatdb";
        //static ip working fine
        public const string connString = "server=(address=192.168.1.62:3306,priority=100),(address=122.165.187.180:8006,priority=50),(address=103.92.203.198:8006,priority=50);user id = esskay; password=esskay1505;database=epmatdb";

        public MySqlConnection myconn = new MySqlConnection(connString);
        public MySqlCommand mycmd = new MySqlCommand();
        public string mysql = "";
        public MySqlDataReader myrdr = null;

        ArrayList projectmast = new ArrayList();
        ArrayList empmast = new ArrayList();
        ArrayList ncmdata = new ArrayList();
        ArrayList chkdata = new ArrayList();
        public TSM.Model MyModel;


        ////EnableDisable_Controls function work's based on visible of control
        ////if control visible is off then enable / disable not works
        ////if control visible is on then only enable / disable works
        //private void EnableDisable_Controls(string mycondition = "")
        //{

        //}



        //private void SetUserInteface_Tabpage(string TabPageName)
        //{
        //    if (TabPageName == "DashBoard")
        //    {
        //        SetUserInteface_DashBoard();
        //    }
        //    else if (TabPageName == "QC")
        //    {
        //        SetUserInterface_QC(id_process);
        //    }

        //}

        //private void ControlSetting(string myactivity = "")
        //{
        //    // int qcprocess = 0, bool isowner = false, string action = "New", int option = -1
        //    bool flag = false;
        //    int current_tabindex = TabControl1.SelectedIndex;
        //    int current_process = id_process;
        //    bool owner = isowner;
        //    if (current_tabindex == 0)
        //    {
        //        if (myactivity == "New")
        //        {
        //            //EnableDisableControl(flag);
        //            flag = true;
        //            //since new condition
        //            btn1.Text = "+";
        //            btn1.Enabled = flag;
        //            lblqmsdblist.Enabled = flag;
        //            cmbqmsdblist.Enabled = flag;
        //        }
        //        else if (myactivity == "Checking" )
        //        {
        //            //EnableDisableControl(flag);
        //            flag = true;
        //            //since new condition
        //            btn1.Text = "-";
        //            btn1.Enabled = false;
        //            lblqmsdblist.Enabled = flag;
        //            cmbqmsdblist.Enabled = flag;
        //        }
        //        .

        //        if (stage_displayname == "New")
        //        {
        //            btn2.Text = "-";
        //            //btn1.Visible = true;
        //            btn2.Enabled = true;
        //            btn3.Text = "REN";
        //            //btn1.Visible = true;
        //            btn3.Enabled = true;
        //            //lblqmsdblist.Visible = true;
        //            //cmbqmsdblist.Visible = true;
        //            cmbqmsdblist.Text = "";
        //            //btncomplete.Visible = false;
        //        }
        //        else if (stage_displayname == "Checking")
        //        {
        //            btn2.Text = "-";
        //            //btn1.Visible = true;
        //            btn2.Enabled = true;
        //            btn3.Text = "REN";
        //            //btn1.Visible = true;
        //            btn3.Enabled = true;
        //            //lblqmsdblist.Visible = true;
        //            //cmbqmsdblist.Visible = true;
        //            cmbqmsdblist.Text = report;
        //            cmbqmsdblist.Enabled = true;
        //        }
        //        else if (stage_displayname == "Completed")
        //        {
        //            //btncomplete.Visible = false;
        //            //lblqmsdblist.Visible = false;
        //            //cmbqmsdblist.Visible = false;
        //        }
        //    }
        //    else if (current_tabindex == 1)
        //    {

        //    }

        //}


        private void EnableDisableControl(bool flag)
        {
            //for initial load condition all controls disabled
            btnclear.Enabled = flag;
            btnfetch.Enabled = flag;
            btnSave.Enabled = flag;
            btnAdd.Enabled = flag;
            btnDelete.Enabled = flag;
            btnUnique.Enabled = flag;
            btnnextprocess.Enabled = flag;
            btntspreview2.Enabled = flag;
            btnfilter.Enabled = flag;
            //btnexport.Enabled = flag;
            btncopy.Enabled = flag;
            btnsplitsave.Enabled = flag;
            btnpreviousprocess.Enabled = flag;
            btnrebackdraft.Enabled = flag;
            btnview.Enabled = flag;

            chkzoom.Enabled = flag;
            chkreplace.Enabled = flag;

            lblqmsdblist.Enabled = flag;
            cmbqmsdblist.Enabled = flag;
            btn1.Enabled = flag;
            btn2.Enabled = flag;
            btn3.Enabled = flag;
            btncomplete.Enabled = flag;
        }

        private void SetUserInteface_DashBoard()
        {
            lblqmsdblist.Text = "Report";

            //if (tpgdash.Tag != null)
            //    cmbqmsdblist.Text = tpgdash.Tag.ToString();
            //else
            //    cmbqmsdblist.Text = "";


            btn1.Text = "+";
            //btn1.Visible = true;

            btn2.Text = "";
            //btn1.Visible = false;

            btn3.Text = "";
            //btn1.Visible = false;

            btncomplete.Text = "";
            //btncomplete.Visible = false;
            btncomplete.Enabled = false;

            cmbqmsdblist.DropDownStyle = ComboBoxStyle.Simple;

            //set all other control visible to false
            ControlVisible_DashBoardAndQC(false);
        }

        private void SetUserInterface_QC(int current_process)
        {
            ToolTip mytooltip = new ToolTip();
            //export is part dash board hence disable during qc process
            btnexport.Enabled = false;
            btnrefresh.Enabled = false;
            btnTSAll.Enabled = false;
            btnTSClear.Enabled = false;
            //btnpreview.Enabled = false;

            if (lblstage.Text == "New")
            {
                btnAdd.Enabled = true;
                btnpreviousprocess.Enabled = false;
            }
            else if (lblstage.Text == "QA" || lblstage.Text == "COMPLETED" || lblstage.Text == "HELD")
            { 

                btnpreviousprocess.Enabled = false;
                btnnextprocess.Enabled = false;
            }
            //else
            btnfetch.Enabled = true;
            btnview.Enabled = true;
            lblqmsdblist.Text = "Entry";
            //lblqmsdblist.Enabled = true;
            //cmbqmsdblist.Enabled = true;            
            cmbqmsdblist.DropDownStyle = ComboBoxStyle.DropDown;

            //    //set button text as  string.Empty
            btn1.Text = string.Empty;
            btn2.Text = string.Empty;
            btn3.Text = string.Empty;
            btncomplete.Text = string.Empty;

            btn1.Enabled = false;
            btn2.Enabled = false;
            btn3.Enabled = false;
            btncomplete.Enabled = false;

            ArrayList mydatagrid = new ArrayList();
            ArrayList mydbcolumn = new ArrayList();
            mydatagrid.Add("guid,0,1");
            mydatagrid.Add("idqc,0,1");
            mydatagrid.Add("TeklaSKtype,1,1");
            mydatagrid.Add("stage,0,1");

            //database column name for the corresponding datagrid column for fetching sql correctly
            mydbcolumn.Add("st|guid");
            mydbcolumn.Add("in|idqc");
            mydbcolumn.Add("st|guid_type");
            mydbcolumn.Add("in|id_sk_process");


            //based on process set datagrid visiblity and readonly option for each column
            //Column Name, IsVisible, IsReadOnly
            if (current_process < (int)SKProcess.Modelling_Completed)
            {
                //Modelling process
                mydatagrid.Add("Remark,1,0");
                mydbcolumn.Add("st|mod_rmk");
                btn1.Text = "R";
                mytooltip.SetToolTip(this.btn1, "Remark");

                //adjust data grid column width based on visible
                dgqc.Columns["Remark"].Width = 700;
                dgqc.Columns["TeklaSKtype"].Width = 100;
            }
            else if ((current_process >= (int)SKProcess.Modelling_PM_Approved) && (current_process < (int)SKProcess.ModelChecking_Completed))
            {
                //Checking process       
                mydatagrid.Add("Observation,1,0");
                mydatagrid.Add("Remark,1,0");
                mydatagrid.Add("beforeremark,1,1");

                mydbcolumn.Add("st|obser");
                mydbcolumn.Add("st|chk_rmk");
                mydbcolumn.Add("st|beforeremark");
                btn1.Text = "O";
                btn2.Text = "R";
                btn3.Text = "NC";
                mytooltip.SetToolTip(this.btn1, "Observation");
                mytooltip.SetToolTip(this.btn2, "Remark");
                mytooltip.SetToolTip(this.btn3, "No Comment");
                //adjust data grid column width based on visible
                dgqc.Columns["Observation"].Width = 350;
                dgqc.Columns["Remark"].Width = 100;
                dgqc.Columns["beforeremark"].Width = 250;
                dgqc.Columns["TeklaSKtype"].Width = 100;
            }
            else if ((current_process >= (int)SKProcess.ModelChecking_PM_Approved) && (current_process < (int)SKProcess.ModelBackDraft_Completed))
            {

                //back drafting process         
                mydatagrid.Add("Observation,1,1");
                mydatagrid.Add("IsAgreed,1,0");
                mydatagrid.Add("ModelUpdated,1,0");
                mydatagrid.Add("Remark,1,0");
                mydatagrid.Add("Category,1,0");
                mydatagrid.Add("beforeremark,1,1");

                mydbcolumn.Add("st|obser");
                mydbcolumn.Add("in|bd_isagree");
                mydbcolumn.Add("in|bd_ismodelupdated");
                mydbcolumn.Add("st|bd_rmk");
                mydbcolumn.Add("in|id_cat");
                mydbcolumn.Add("st|beforeremark");


                btn1.Text = "R";
                mytooltip.SetToolTip(this.btn1, "Remark");

            }
            else if ((current_process >= (int)SKProcess.ModelBackDraft_PM_Approved) && (current_process < (int)SKProcess.ModelBackCheck_Completed))
            {
                //back check process
                //mydatagrid.Add("Owner,1,1");
                mydatagrid.Add("Observation,1,1");
                mydatagrid.Add("IsAgreed,1,1");
                mydatagrid.Add("ModelUpdated,1,1");
                mydatagrid.Add("Category,1,1");
                mydatagrid.Add("BackCheck,1,0");
                mydatagrid.Add("backcheck_obser,1,0");
                mydatagrid.Add("Remark,1,0");
                mydatagrid.Add("beforeremark,1,1");
                mydatagrid.Add("mod_user_id,1,1");



                //mydbcolumn.Add("in|mod_user_id");
                mydbcolumn.Add("st|obser");
                mydbcolumn.Add("in|bd_isagree");
                mydbcolumn.Add("in|bd_ismodelupdated");
                mydbcolumn.Add("in|id_cat");
                mydbcolumn.Add("in|backcheck");
                mydbcolumn.Add("st|backcheck_obser");
                mydbcolumn.Add("st|backcheck_rmk");
                mydbcolumn.Add("st|beforeremark");
                mydbcolumn.Add("in|mod_user_id");

                dgqc.Columns["Observation"].Width = 350;
                dgqc.Columns["Remark"].Width = 100;
                dgqc.Columns["beforeremark"].Width = 250;
                dgqc.Columns["TeklaSKtype"].Width = 100;



                btn1.Text = "R";
                mytooltip.SetToolTip(this.btn1, "Remark");
            }
            else if ((current_process >= (int)SKProcess.ModelBackCheck_PM_Approved) && (current_process < (int)SKProcess.ModelSeverity_PM_Approved))
            {
                //for severity entry
                mydatagrid.Add("Observation,1,1");
                mydatagrid.Add("Category,1,1");
                mydatagrid.Add("Severity,1,0");
                mydatagrid.Add("Remark,1,0");
                mydatagrid.Add("beforeremark,1,1");

                mydbcolumn.Add("st|obser");
                mydbcolumn.Add("in|id_cat");
                mydbcolumn.Add("in|id_sev");
                mydbcolumn.Add("st|sev_rmk");
                mydbcolumn.Add("st|beforeremark");

                btn1.Text = "R";
                mytooltip.SetToolTip(this.btn1, "Remark");

            }
            else if ((current_process >= (int)SKProcess.ModelSeverity_PM_Approved))
            {

                //Sl.No GUID     Observation Category    Severity Owner   Agreed / Not Agreed Corrected By Remarks

                //completed ...all are ready only
                mydatagrid.Add("Observation,1,1");
                mydatagrid.Add("Category,1,1");
                mydatagrid.Add("Severity,1,1");
                mydatagrid.Add("Owner,1,1");
                mydatagrid.Add("IsAgreed,1,1");
                mydatagrid.Add("CorrectedBy,1,1");
                //mydatagrid.Add("BackCheck,1,1");
                mydatagrid.Add("beforeremark,1,1");


                mydbcolumn.Add("st|obser");
                mydbcolumn.Add("in|id_cat");
                mydbcolumn.Add("in|id_sev");
                mydbcolumn.Add("in|mod_user_id");
                mydbcolumn.Add("in|bd_isagree");
                mydbcolumn.Add("in|bd_user_id");
                //mydbcolumn.Add("in|backcheck");
                mydbcolumn.Add("st|beforeremark");

            }

            if (btn1.Text.Length >= 1)
                btn1.Enabled = true;


            if (btn2.Text.Length >= 1)
                btn2.Enabled = true;

            if (btn3.Text.Length >= 1)
                btn3.Enabled = true;

            if (btncomplete.Text.Length >= 1)
                btncomplete.Enabled = true;

            lblqmsdblist.Enabled = btn1.Enabled;
            cmbqmsdblist.Enabled = btn1.Enabled;

            chkreplace.Enabled = cmbqmsdblist.Enabled;


            //set all column visible status = false
            HideAllDataGridViewColumn(dgqc);

            //set data grid visible
            setDataGridColumnVisible(dgqc, mydatagrid, mydbcolumn);




            //if (emp_role.ToUpper().Trim() == "CHECKER") //&& tbpgUI.Tag == null
            //{
            tpgqc.Tag = emp_role;

            //}


            //based on setting ensure datagrid columns are updated //future we should have flag to identify during changes only this on or off should work
            Update_Tekla_MemberProperties_UserField();

        }

        //private void SetUserInterface_QC(int id_process)
        //{

        //    //tabpage dash board
        //    SetControlVisible_DashBoard(false);

        //    ////all button and check box visible set to false
        //    //ControlVisible_DashBoardAndQC(false);

        //    //set button text as  string.Empty
        //    btn1.Text = string.Empty;
        //    btn2.Text = string.Empty;
        //    btn3.Text = string.Empty;
        //    btncomplete.Text = string.Empty;

        //    lblqmsdblist.Text = "Entry";

        //    //btn1.Text = "O";
        //    ////btn1.Visible = true;

        //    //btn2.Text = "R";
        //    ////btn1.Visible = true;

        //    //btn3.Text = "";
        //    ////btn1.Visible = false;

        //    //btncomplete.Text = "";
        //    ////btncomplete.Visible = false;
        //    //btncomplete.Enabled = false;
        //    cmbqmsdblist.DropDownStyle = ComboBoxStyle.DropDown;

        //    ArrayList mydatagrid = new ArrayList();
        //    ArrayList mycolumn = new ArrayList();
        //    mydatagrid.Add("guid,0,1");
        //    mydatagrid.Add("idqc,0,1");
        //    mydatagrid.Add("type,0,1");
        //    mydatagrid.Add("stage,0,1");

        //    //based on process set datagrid visiblity and readonly option for each column
        //    //Column Name, IsVisible, IsReadOnly
        //    if (id_process < (int)SKProcess.ModelChecking_Completed)
        //    {
        //        //Checking process
        //        mydatagrid.Add("Observation,1,0");
        //        mydatagrid.Add("Remark,1,0");

        //        btn1.Text = "O";
        //        btn2.Text = "R";
        //    }
        //    else if ((id_process >= (int)SKProcess.ModelChecking_Completed) && (id_process < (int)SKProcess.ModelBackDraft_Completed))
        //    {
        //        //back drafting process         
        //        mydatagrid.Add("Observation,1,1");
        //        mydatagrid.Add("IsAgreed,1,0");
        //        mydatagrid.Add("Remark,1,0");
        //        mydatagrid.Add("Category,1,0");

        //        btn1.Text = "R";
        //    }
        //    else if ((id_process >= (int)SKProcess.ModelBackDraft_Completed) && (id_process < (int)SKProcess.ModelBackCheck_Completed))
        //    {
        //        //back check process
        //        mydatagrid.Add("Observation,1,1");
        //        mydatagrid.Add("beforeremark,1,1");
        //        mydatagrid.Add("IsAgreed,1,1");
        //        mydatagrid.Add("Category,1,1");
        //        mydatagrid.Add("BackCheck,1,0");
        //        mydatagrid.Add("Remark,1,0");

        //        btn1.Text = "R";
        //    }
        //    else if ((id_process >= (int)SKProcess.Completed))
        //    {
        //        //completed
        //        mydatagrid.Add("Observation,1,1");
        //        mydatagrid.Add("beforeremark,1,1");
        //        mydatagrid.Add("IsAgreed,1,1");
        //        mydatagrid.Add("Category,1,1");
        //        mydatagrid.Add("Severity,1,1");
        //        mydatagrid.Add("BackCheck,1,1");
        //        mydatagrid.Add("Remark,1,1");
        //    }

        //    if (btn1.Text.Length >= 1)
        //        //btn1.Visible = true;

        //        //lblqmsdblist.Visible = btn1.Visible;
        //        //cmbqmsdblist.Visible = btn1.Visible;
        //        //chkreplace.Visible = cmbqmsdblist.Visible;
        //        //if (tpgts.Tag != null)
        //        //    cmbqmsdblist.Text = tpgts.Tag.ToString();
        //        //else
        //        //    cmbqmsdblist.Text = "";


        //        //    if (btn2.Text.Length >= 1)
        //        //        btn1.Visible = true;

        //        //if (btn3.Text.Length >= 1)
        //        //    btn1.Visible = true;

        //        //if (btncomplete.Text.Length >= 1)
        //        //    btncomplete.Visible = true;
        //        if (lblstage.Text == "New")
        //        {
        //            btnAdd.Enabled = true;
        //            //btnAdd.Visible = true;
        //            btnSave.Enabled = true;
        //            //btnSave.Visible = true;
        //            btnfetch.Enabled = false;
        //            //btnfetch.Visible = false;
        //        }
        //        else
        //        {
        //            btnfetch.Enabled = true;
        //            //btnfetch.Visible = true;
        //        }

        //    ////set button control
        //    //if (id_process < (int)SKProcess.ModelChecking_Completed)
        //    //{
        //    //    btnAdd.Enabled = true;
        //    //    btnDelete.Enabled = true;
        //    //}
        //    //else
        //    //{
        //    //    btnAdd.Enabled = false;
        //    //    btnDelete.Enabled = false;
        //    //}

        //    ////for completed process there is no requirement for save, select, delete and split hence disabled
        //    //if (id_process == 500)
        //    //{
        //    //    btnSave.Enabled = false;
        //    //    btnAdd.Enabled = false;
        //    //    btnDelete.Enabled = false;
        //    //    btnsplitsave.Enabled = false;
        //    //}

        //    //////

        //    ////based on role set the User Interface
        //    //FirstTimeSetting_UserInterface(lblrole.Text, id_process);
        //    ////if (skWinLib.IsNumeric(lblstage.Tag) == true)
        //    ////FirstTimeSetting_UserInterface(lblrole.Text, Convert.ToInt32(lblstage.Tag));
        //    ///
        //    //set all column visible status = false
        //    HideAllDataGridViewColumn(dgqc);
        //    //dgqc.Columns.Clear();


        //    //set data grid visible
        //    setDataGridColumnVisible(dgqc, mydatagrid);



        //    //set all other control visible to true
        //    //SetControlVisible(true);







        //    //if (emp_role.ToUpper().Trim() == "CHECKER") //&& tbpgUI.Tag == null
        //    //{
        //    tpgqc.Tag = emp_role;

        //    //}


        //    //based on setting ensure datagrid columns are updated //future we should have flag to identify during changes only this on or off should work
        //    Update_Tekla_MemberProperties_UserField();

        //}


        public frmQC()
        {
            InitializeComponent();
            //Load_Project();
            TabControl1.Selected += new TabControlEventHandler(tabNCM_Selected);

        }

        private void settooltip()
        {
            // Create the ToolTip and associate with the Form container.
            ToolTip mytooltip = new ToolTip();

            // Set up the delays for the ToolTip.
            mytooltip.AutoPopDelay = 5000;
            mytooltip.InitialDelay = 1000;
            mytooltip.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            mytooltip.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            //tab page dash board
            mytooltip.SetToolTip(this.btn1, "Add");
            mytooltip.SetToolTip(this.btn2, "Delete");
            mytooltip.SetToolTip(this.btn3, "Rename");
            mytooltip.SetToolTip(this.btncomplete, "Complete. Move to next stage");
            mytooltip.SetToolTip(this.chktop, "Front/Back");
            mytooltip.SetToolTip(this.chkzoom, "Zoom Object(s)");
            mytooltip.SetToolTip(this.chkreplace, "Replace");

            //tab page qc
            mytooltip.SetToolTip(this.btnfetch, "Get");
            mytooltip.SetToolTip(this.btnview, "View");
            mytooltip.SetToolTip(this.btnclear, "Clear");
            mytooltip.SetToolTip(this.btnSave, "Save");
            mytooltip.SetToolTip(this.btnAdd, "Add");
            mytooltip.SetToolTip(this.btnDelete, "Delete");
            mytooltip.SetToolTip(this.btnnextprocess, "Preview Observation");
            mytooltip.SetToolTip(this.btntspreview2, "Preview Category");
            mytooltip.SetToolTip(this.btnUnique, "Group");
            mytooltip.SetToolTip(this.btncopy, "Copy");
            mytooltip.SetToolTip(this.btnexport, "Export");
            mytooltip.SetToolTip(this.btnrefresh, "Refresh");
            mytooltip.SetToolTip(this.btnpreview, "Preview Project");
            mytooltip.SetToolTip(this.btnTSAll, "Select All");
            mytooltip.SetToolTip(this.btnTSClear, "Clear All");
            mytooltip.SetToolTip(this.btnsplitsave, "Split");
            mytooltip.SetToolTip(this.btnpreviousprocess, "Move Record(s) previous process");
            mytooltip.SetToolTip(this.btnnextprocess, "Move Record(s) next process");
            mytooltip.SetToolTip(this.btnrebackdraft, "For BackDraft");


            mytooltip.SetToolTip(this.btnhlp, "Help");
        }
        private void frmPS_Load(object sender, EventArgs e)
        {
            skdebug = skWinLib.DebugFlag();

            //Update User Interface Application Name and Latest Version
            //this.Text = skApplicationName + " [" + skApplicationVersion + "]";

            //MessageBox.Show("TS-2021 \n\n222");
            msgCaption = this.Text + " " + skApplicationVersion ;

            //Esskay Security Validation
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");
            //MessageBox.Show("TS-2021 \n\n2");

            //Check whether mySQL connection established if not inform user.
            if (IsCheckconnection() == false)
            {
                skWinLib.accesslog(skApplicationName, skApplicationVersion, "EsskayValidation", "Could not establisth connection with MySQL.", skTSLib.ModelName , skTSLib.Version, skTSLib.Configuration);
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
                    txtproject.Text = esskayjn;


                    //intialize dashboard treeview 
                    //InitalizeTreeview(esskayjn);

                    //Disable all controls since first time
                    EnableDisableControl(false);

                    //setting for data grid
                    dgqc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    dgqc.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;  
                    dgqc.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders);


                    //get windows login to avoid multiple factor authentication
                    String empid = skWinLib.username;
                    //if (skWinLib.username.ToUpper() == "VIJAYAKUMAR")
                    //{
                    //    empid = "SKS360";
                    //    //lblrole.Text = "Checker";
                    //    //empid = "SKS269";
                    //    //lblrole.Text = "Modeller";
                    //}
                    //else
                    if (skWinLib.username.ToUpper() == "VISHWARAM")
                    {
                        empid = "vishwaramp";
                    }
                    //else if (skWinLib.username.ToUpper() == "ULAGANATHAN")
                    //{
                    //    //                Ulaganathan / 224
                    //    empid = "SKS269";
                    //    //lblrole.Text = "Modeller";
                    //}
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
                    emp_role = lblrole.Text;

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
                        empmast.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString() + "|" + myrdr[2].ToString() + "|" + myrdr[3].ToString());
                    }
                    myrdr.Close();


                    id_emp = -1;
                    //get  users.id from empmast
                    if (empmast.Count >= 1)
                    {
                        ArrayList tmp = getDBUserID(empid);
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                            id_emp = Convert.ToInt32(empsplit[0]);
                            lblxsuser.Text = empsplit[2].ToString();
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

                        //Update user interface with last saved information in registry
                        UpdateUserInterface_Fromlastsave();


                        //MessageBox.Show("TS-2021 \n\n3");

                        //set all column visible status, readonly status
                        chkTSAttribute.Items.Clear();

                        string[] split = TS_Extra_Column.Split(new Char[] { ',' });
                        foreach (string s in split)
                        {
                            chkTSAttribute.Items.Add(s);
                            dgqc.Columns.Add("TeklaSK" + s, s);

                        }

                        //in future restore chklst items automatically



                        //lblsbar2.Text = "";
                        this.TopMost = true;


                        //hide ts report
                        //TabControl1.TabPages.Remove(tpgts);

                        this.tpgqc.Parent = null;
                        //TabControl1.TabPages[1].Hide = true;
                        //(TabControl1.TabPages[1] as TabPage).Enabled = false;

                        //set tool tip
                        settooltip();

                        //for debugging set lblsbar2 visible as true
                        if (skWinLib.username.ToUpper() == "VIJAYAKUMAR" || skWinLib.username.ToUpper() == "SKS360")
                        {
                            lblsbar2.Visible = true;
                        }
                        else
                        {
                            lblsbar2.Visible = false;
                        }
                        TabControl1.TabPages.Remove(tpgqcmirror);

                    }
                }
            }

        }

        //private void SetHelpFile()
        //{
        //    Create new StreamReader
        //    StreamReader sr = new StreamReader(dlgOpen.FileName, Encoding.Default);
        //    Get all text from the file
        //    string str = sr.ReadToEnd();
        //    Close the StreamReader
        //    sr.Close();
        //    Show the text in the rich textbox rtbMain
        //    rtxthelp.Text = str;
        //}
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
        private ArrayList getDBUserID(int id)
        {
            ArrayList myreturn = new ArrayList();
            //check if empmast is not null
            for (int i = 0; i < empmast.Count; i++)
            {
                string[] empsplit = empmast[i].ToString().Split(new Char[] { '|' });
                if (empsplit.Length >= 1)
                {
                    if (empsplit[0].ToString() == id.ToString())
                    {
                        myreturn.Add(empmast[i]);
                        break;
                    }
                }
            }
            return myreturn;
        }
        private bool IsCheckconnection()
        {

            //        //public const string connString = "server=192.168.1.62;Port=3306;user id = esskay; password=esskay1505;database=epmatdb";
            //        //http://103.92.203.198:8001/login and http://122.165.187.180:8001/login
            //        //public const string connString = "server=103.92.203.198;Port=8001;user id = esskay; password=esskay1505;database=epmatdb";
            //        //public const string connString = "server=122.165.187.180;Port=8001;user id = esskay; password=esskay1505;database=epmatdb";

            //string connString = "server=127.0.0.1;Port=3306;user id = root; password=Mysqlpassword1;database=epmatdb";
            //connString = "server=103.92.203.198;Port=3306;user id = root; password=Mysqlpassword1;database=epmatdb";

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

        private void Checkconnection()
        {
            try
            {
                myconn.Open();
                string sql = "SELECT * FROM ncm.empmast";
                mycmd.CommandText = sql;
                MySqlDataReader rdr = mycmd.ExecuteReader();
                while (rdr.Read())
                {

                    Console.WriteLine(rdr[0] + " -- " + rdr[1]);
                }
                rdr.Close();
                myconn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }



        //function to hide all column in datagrid view
        private void HideAllDataGridViewColumn(DataGridView mydatagrid)
        {
            //loop throug and hide
            foreach (DataGridViewColumn dgcol in mydatagrid.Columns)
            {
                dgcol.Visible = false;

            }
        }

        ////function to create column in datagrid view (input name,width,is readonly, celltype, visible) 
        //private void CreateDataGridViewColumn(DataGridView mydatagrid, ArrayList mylist)
        //{
        //    //hide all column
        //    HideAllDataGridViewColumn(mydatagrid);

        //}


        //private void setDataGridColumnVisible(DataGridView mydgQMS, ArrayList mydatagrid)
        //{
        //    foreach (string mycolheader in mydatagrid)
        //    {
        //        string[] spltmycolheader = mycolheader.ToString().Split(new Char[] { ',' });
        //        //get the column name
        //        string dgcolname = spltmycolheader[0];

        //        //set visible based on condition
        //        if (spltmycolheader[1].ToString() == "1")
        //        {
        //            DataGridViewColumn newcol = mydgQMS.Columns[dgcolname];
        //            //mydgQMS.Columns.Add(mydgQMS.Columns[dgcolname] as DataGridViewColumn);
        //            mydgQMS.Columns.Add(newcol);
        //            mydgQMS.Columns[dgcolname].Visible = true;

        //            //set readonly based on condition
        //            if (spltmycolheader[2].ToString() == "1")
        //                mydgQMS.Columns[dgcolname].ReadOnly = true;
        //            else
        //                mydgQMS.Columns[dgcolname].ReadOnly = false;
        //        }




        //    }


        //    ////full row select
        //    //dgqc.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        //}


        private void setDataGridColumnVisible(DataGridView mydgQMS, ArrayList mydatagrid, ArrayList sql_columnname)
        {

            for (int i = 0; i < mydatagrid.Count; i++)
            {
                string[] spltmycolheader = mydatagrid[i].ToString().Split(new Char[] { ',' });
                //get the column name
                string dgcolname = spltmycolheader[0];

                //set visible based on condition
                if (spltmycolheader[1].ToString() == "1")
                    mydgQMS.Columns[dgcolname].Visible = true;
                else
                    mydgQMS.Columns[dgcolname].Visible = false;

                //set readonly based on condition
                if (spltmycolheader[2].ToString() == "1")
                    mydgQMS.Columns[dgcolname].ReadOnly = true;
                else
                    mydgQMS.Columns[dgcolname].ReadOnly = false;

                //set db column name as tag
                mydgQMS.Columns[dgcolname].Tag = sql_columnname[i].ToString();
            }


            ////full row select
            //dgqc.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }












        //private DrawingHandler DrawingHandler = new DrawingHandler();
        public static bool IsValidDirectory(string directory)
        {
            if (directory == null || directory == "")
                return false;
            return Directory.Exists(directory);
        }




        private void lblsbar1_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "lblsbar1_Click";
            if (lblsbar2.Visible == false)
            {
                lblsbar2.Visible = true;
            }

            else
            {
                lblsbar2.Visible = false;
                //lblsbar2.Text = "";
            }

        }
        private void Update_Tekla_MemberProperties_UserField()
        {
            bool chkstate = false;
            string colname = "";
            for (int i = 0; i < chkTSAttribute.Items.Count; i++)
            {
                if (chkTSAttribute.GetItemCheckState(i) == CheckState.Unchecked)
                {

                    chkstate = false;
                }
                if (chkTSAttribute.GetItemCheckState(i) == CheckState.Checked)
                {

                    chkstate = true;

                }
                colname = chkTSAttribute.Items[i].ToString();
                this.dgqc.Columns["TeklaSK" + colname].Visible = chkstate;
            }
        }
        private void chklsthdr_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblsbar2.Text = "chklsthdr_SelectedIndexChanged";
            Update_Tekla_MemberProperties_UserField();
        }

        private bool IsTeklaAdditionalEnabled()
        {
            for (int i = 0; i < chkTSAttribute.Items.Count; i++)
            {
                if (chkTSAttribute.GetItemCheckState(i) == CheckState.Checked)
                    return true;
            }
            return false;
        }

        private void dgQMS_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //lblsbar2.Text = "dgQMS_CellMouseDoubleClick";

            //lblsbar1.Text = e.RowIndex.ToString() + "...";
            //if (e.RowIndex >= 0)
            //{
            //    int rowindex = e.RowIndex;
            //    DataGridViewRow row = this.dgqc.Rows[rowindex];


            //    string myguid = dgqc.Rows[e.RowIndex].Cells["guid"].Value.ToString();
            //    Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier(myguid);
            //    bool flag = myIdentifier.IsValid();
            //    //Type mytype = myIdentifier.GetType();
            //    //string name = mytype.Name;
            //    //string type = dgQMS.Rows[e.RowIndex].Cells["TeklaSKtype"].Value.ToString();                
            //    //myIdentifier = new Tekla.Structures.Identifier(myguid);
            //    ArrayList ObjectsToSelect = new ArrayList();
            //    if (flag)
            //    {
            //        Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
            //        if (ModelSideObject.Identifier.IsValid())
            //        {
            //            //ModelSideObject.Select();
            //            if (!ObjectsToSelect.Contains(ModelSideObject))
            //                ObjectsToSelect.Add(ModelSideObject);
            //        }

            //        var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            //        MOS.Select(ObjectsToSelect);
            //        if (chkzoom.Checked == true) TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
            //        //lblxsuser.Text 
            //    }




            //}
            //lblsbar1.Text = "Ready";
            //lblsbar2.Text = "";
        }


        private void label1_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "label1_Click";
            //lblsbar2.Text = "";
        }









        private void dgQMS_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            lblsbar2.Text = "dgQMS_CellContentClick";
            //lblsbar2.Text = "";
        }

        private void Load_Project([Optional] string sqlCondition)
        {
            //All project information
            //mysql = "select id,esskayjn,name,clientjn from epmatdb.project " + sqlCondition;
            //myconn = new MySqlConnection(connString);
            //if (myconn.State == ConnectionState.Closed)
            //    myconn.Open();
            //mycmd.CommandText = mysql;
            //mycmd.Connection = myconn;
            //myrdr = mycmd.ExecuteReader();
            //cmbprojid.Items.Clear();
            //projectmast.Clear();
            //while (myrdr.Read())
            //{

            //    cmbprojid.Items.Add(myrdr[1].ToString().ToUpper());
            //    projectmast.Add(myrdr[1].ToString() + "|" + myrdr[2].ToString() + "|" + myrdr[3].ToString() + "|" + myrdr[0].ToString());
            //}
            //myrdr.Close();



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
                            cmbprojid.Items.Clear();
                            while (myrdr.Read())
                            {
                                cmbprojid.Items.Add(esskayjn);
                                client = myrdr[1].ToString();
                            }
                            myrdr.Close();
                            //if (client == string.Empty)
                            //{
                            //    MessageBox.Show(esskayjn + " not a valid Project.\nUpdate Correct SK Project Number in Model Project Properties and try again.\n\n\nAborting Now!!!", "Esskay Invalid Project Number", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            //    //System.Windows.Forms.Application.ExitThread();
                            //    System.Environment.Exit(0);
                            //}

                            this.Text = skApplicationName + " [" + skApplicationVersion + "]" + "      " + esskayjn + " / " + projectname + " / " + client;
                            ////set sorted based on UI setting
                            ////if (cmbprojid.Sorted == true)
                            ////    projectmast.Sort();

                            //cmbprojid.DropDownStyle = ComboBoxStyle.Simple;
                            //cmbprojid.Text = esskayjn;
                            ////cmbprojid.Items = 0;
                            ////set the project_id in cmbprojid.tag
                            //string[] projectinfo = projectmast[cmbprojid.SelectedIndex].ToString().Split(new Char[] { '|' });
                            ////cmbprojid.Tag = projectinfo[3];
                            //id_project = Convert.ToInt32(projectinfo[3].ToString());

                            //if (cmbprojid.SelectedIndex == -1)
                            //{
                            //    MessageBox.Show(esskayjn + " not a valid Project.\nUpdate Correct SK Project Number in Model Project Properties and try again.\n\n\nAborting Now!!!", "Esskay Invalid Project Number", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            //    //System.Windows.Forms.Application.ExitThread();
                            //    System.Environment.Exit(0);
                            //}

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
                    cmbprojid.Enabled = false;
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

        private bool Update_DashBoard_TreeView(TreeNode currentprocess, TreeNode nextprocess, int new_id, int newprocess_id)
        {
            lblsbar2.Text = "Update_DashBoard_TreeView,currentprocess:" + currentprocess.Text.ToString() + ",new_id:" + new_id.ToString() + ",newprocess_id:" + newprocess_id.ToString();

            for (int i = 0; i < currentprocess.Nodes.Count; i++)
            {
                if (currentprocess.Nodes[i].Text.ToUpper().Trim() == lblreport.Text.ToUpper().Trim())
                {
                    //delete new node
                    currentprocess.Nodes[i].Remove();
                    //add to in process node
                    TreeNode tmp = nextprocess.Nodes.Add(lblreport.Text.Trim());
                    lblstage.Text = tmp.Parent.Text;

                    id_stage = new_id;
                    id_process = newprocess_id;
                    tmp.Tag = id_process + "|" + id_stage;
                    return true;

                }
            }

            return false;
        }

        private void InitalizeTreeview(string SK_ProjectNumber)
        {
            //update treeview
            tvwncm.Nodes.Clear();
            tvwroot = tvwncm.Nodes.Add("QC - DO IT CORRECT FIRST TIME AND EVERYTIME.");
            tvwproject = tvwroot.Nodes.Add(SK_ProjectNumber);
            tvwnew = tvwproject.Nodes.Add("New");
            tvwchecking = tvwproject.Nodes.Add("Checking");
            tvwbackdraft = tvwproject.Nodes.Add("Back Drafting");
            tvwbackcheck = tvwproject.Nodes.Add("Back Checking");
            //tvwconflict = tvwproject.Nodes.Add("Conflict");
            tvwseverity = tvwproject.Nodes.Add("For Severity");
            tvwqa = tvwproject.Nodes.Add("QA");
            tvwheld = tvwproject.Nodes.Add("Held");
            tvwcancelled = tvwproject.Nodes.Add("Cancelled");
            tvwcompleted = tvwproject.Nodes.Add("Completed");
            tvwncm.ExpandAll();

            //tvwroot.ForeColor = System.Drawing.Color.Red;
            tvwroot.BackColor = System.Drawing.Color.LightBlue;
            tvwheld.ForeColor = System.Drawing.Color.Red;
            tvwcancelled.ForeColor = System.Drawing.Color.Red;

        }

        //refresh dash board treeview
        //next version shall have each node update rather then updating all node for better performance
        private void Update_DashBoard_TreeView()
        {
            //Update Treeview DashBoard
            InitalizeTreeview(esskayjn);

            //ArrayList mydashboarddata = GetReportStage_Based_Project_User(Convert.ToInt32(cmbprojid.Tag), Convert.ToInt32(lblxsuser.Tag));
            //ArrayList mydashboarddata = GetReportStage_Based_Project_User(id_project ,id_emp);
            ArrayList reportdata = GetReportStage_Based_Project(id_project);

            if (reportdata != null)
            {
                if (reportdata.Count >= 1)
                {
                    ArrayList mydashboarddata = reportdata[1] as ArrayList;
                    TreeNode childnode = null;
                    TreeNode child_dashdata = null;
                    foreach (string mydata in mydashboarddata)
                    {
                        //check whether already we have |||
                        string tmp = mydata;
                        //tmp = tmp.Replace(sk_seperator, "");
                        //if ((tmp.Length + 18) == mydata.Length)
                        {
                            tmp = mydata.Replace(sk_seperator, "|");
                            string[] splt_mydata = tmp.Split(new Char[] { '|' });
                            int id_process = Convert.ToInt32(splt_mydata[1]);
                            int id_ncm_progress = Convert.ToInt32(splt_mydata[0]);
                            string splt_report = "";
                            foreach (string report in reportdata[0] as ArrayList)
                            {
                                string tmp2 = report.Replace(sk_seperator, "|");
                                string[] splt_mydata2 = tmp2.Split(new Char[] { '|' });
                                int id_ncm_progress2 = Convert.ToInt32(splt_mydata2[0]);
                                if (id_ncm_progress == id_ncm_progress2)
                                {
                                    splt_report = splt_mydata2[1].ToString();
                                    break;
                                }
                            }
                            if (splt_report != "")
                            {
                                childnode = null;
                                string process_owner = string.Empty;
                                if ((id_process < (int)SKProcess.Modelling_PM_Approved))
                                {
                                    //new
                                    childnode = tvwnew.Nodes.Add(splt_report);
                                    //process_owner = splt_mydata[2].ToString();

                                }
                                else if ((id_process >= (int)SKProcess.Modelling_PM_Approved) && (id_process < (int)SKProcess.ModelChecking_Completed))
                                {
                                    //checking
                                    childnode = tvwchecking.Nodes.Add(splt_report);
                                    //process_owner = splt_mydata[3].ToString();

                                }
                                else if ((id_process >= (int)SKProcess.ModelChecking_PM_Approved) && (id_process < (int)SKProcess.ModelBackDraft_Completed))
                                {
                                    //for DPM Approval and for back drafting
                                    childnode = tvwbackdraft.Nodes.Add(splt_report);
                                    //process_owner = splt_mydata[4].ToString();

                                }
                                else if ((id_process >= (int)SKProcess.ModelBackDraft_PM_Approved) && (id_process < (int)SKProcess.ModelBackCheck_Completed))
                                {
                                    //for back checking
                                    childnode = tvwbackcheck.Nodes.Add(splt_report);
                                    //process_owner = splt_mydata[5].ToString();
                                }
                                else if ((id_process >= (int)SKProcess.ModelBackCheck_PM_Approved) && (id_process < (int)SKProcess.ModelSeverity_Completed))
                                {

                                    //for severity setting
                                    childnode = tvwseverity.Nodes.Add(splt_report);
                                    //process_owner = splt_mydata[6].ToString();

                                }
                                else if ((id_process >= (int)SKProcess.ModelSeverity_PM_Approved) && (id_process < (int)SKProcess.QA))
                                {

                                    //for RESERVED FOR FUTURE
                                    childnode = tvwqa.Nodes.Add(splt_report);

                                }
                                //else if (id_process == (int)SKProcess.Held)
                                //{
                                //    //completed
                                //    childnode = tvwheld.Nodes.Add(splt_report);
                                //}
                                //else if  (id_process == (int)SKProcess.Cancelled)
                                //{
                                //    //held
                                //    childnode = tvwcancelled.Nodes.Add(splt_report);
                                //}

                                else if ((id_process == (int)SKProcess.Completed))
                                {
                                    //qa
                                    childnode = tvwcompleted.Nodes.Add(splt_report);
                                }

                                //add node tag with id_report
                                if (childnode != null)
                                    childnode.Tag = splt_mydata[0].ToString() + "|" + splt_mydata[1].ToString();

                                if (id_emp.ToString() == process_owner || process_owner.Length == 0)
                                    childnode.ForeColor = System.Drawing.Color.Black;
                                else
                                    childnode.ForeColor = System.Drawing.Color.Red;

                            }
                            else
                            {
                                string chk1 = "1";
                            }
     

                            //red color applicable for checking stage only other all activity like bd, bc don't have control now 
                            //if (childnode.Parent.Text == "Checking")
                            //{
                            //    if (id_emp.ToString() != splt_mydata[3])
                            //        childnode.ForeColor = System.Drawing.Color.Red;
                            //}


                            //if (id_process < (int)SKProcess.ModelBackDraft_PM_Approved)
                            //{
                            //    int total = 0;
                            //    if (skWinLib.IsNumeric(splt_mydata[4]) == true)
                            //        total = Convert.ToInt32(splt_mydata[4]);

                            //    int val1 = 0;
                            //    if (skWinLib.IsNumeric(splt_mydata[5]) == true)
                            //        val1 = Convert.ToInt32(splt_mydata[5]);

                            //    int val2 = 0;
                            //    if (skWinLib.IsNumeric(splt_mydata[6]) == true)
                            //        val2 = Convert.ToInt32(splt_mydata[6]);

                            //    child_dashdata = childnode.Nodes.Add("Total:" + total.ToString());
                            //    if (id_process < (int)SKProcess.ModelChecking_Completed)
                            //    {
                            //        if (val1 != 0)
                            //            child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());
                            //        if (val2 != 0)
                            //            child_dashdata = childnode.Nodes.Add("No Comment:" + val2.ToString());
                            //    }
                            //    else if (id_process < (int)SKProcess.ModelBackDraft_Completed)
                            //    {
                            //        if (val1 != 0)
                            //            child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());
                            //        if (val2 != 0)
                            //            child_dashdata = childnode.Nodes.Add("Pending Cat." + val2.ToString());
                            //    }
                            //    else if (id_process < (int)SKProcess.ModelBackCheck_Completed)
                            //    {
                            //        if (val1 != 0)
                            //            child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());

                            //    }
                            //}

                        }

                    }
                }


            }
        }

        //private void Update_DashBoard_TreeView()
        //{

        //    //update treeview
        //    tvwncm.Nodes.Clear();
        //    tvwroot = tvwncm.Nodes.Add("Esskay << DO IT CORRECT FIRST TIME AND EVERYTIME >>");
        //    tvwproject = tvwroot.Nodes.Add(esskayjn);
        //    tvwnew = tvwproject.Nodes.Add("New");
        //    tvwchecking = tvwproject.Nodes.Add("Checking");
        //    tvwbackdraft = tvwproject.Nodes.Add("Back Drafting");
        //    tvwbackcheck = tvwproject.Nodes.Add("Back Checking");
        //    //tvwheld = tvwproject.Nodes.Add("Held");
        //    //tvwcancelled = tvwproject.Nodes.Add("Cancelled");
        //    tvwqa = tvwproject.Nodes.Add("QA");
        //    tvwcompleted = tvwproject.Nodes.Add("Completed");
        //    tvwncm.ExpandAll();

        //    //Update Treeview DashBoard
        //    //ArrayList mydashboarddata = GetReportStage_Based_Project_User(Convert.ToInt32(cmbprojid.Tag), Convert.ToInt32(lblxsuser.Tag));
        //    //ArrayList mydashboarddata = GetReportStage_Based_Project_User(id_project ,id_emp);
        //    ArrayList mydashboarddata = GetReportStage_Based_Project(id_project);

        //    if (mydashboarddata != null)
        //    {
        //        TreeNode childnode = null;
        //        TreeNode child_dashdata = null;
        //        foreach (string mydata in mydashboarddata)
        //        {
        //            //check whether already we have |||
        //            string tmp = mydata;
        //            tmp = tmp.Replace(sk_seperator, "");
        //            if ((tmp.Length + 18) == mydata.Length)
        //            {
        //                tmp = mydata.Replace(sk_seperator, "|");
        //                string[] splt_mydata = tmp.Split(new Char[] { '|' });
        //                string splt_report = splt_mydata[2].ToString();
        //                int id_process = Convert.ToInt32(splt_mydata[0]);
        //                int id_stage = Convert.ToInt32(splt_mydata[1]);



        //                //if ((id_process >= (int)SKProcess.Modelling_Completed) && (id_process < (int)SKProcess.ModelChecking_PM_Approved))
        //                if ((id_process >= (int)SKProcess.Modelling_Completed) && (id_process < (int)SKProcess.Modelling_PM_Approved))
        //                {
        //                    //inprogress
        //                    childnode = tvwchecking.Nodes.Add(splt_report);

        //                }
        //                else if ((id_process >= (int)SKProcess.ModelChecking_Completed) && (id_process < (int)SKProcess.ModelChecking_PM_Approved))
        //                {
        //                    //for DPM Approval and for back drafting
        //                    childnode = tvwbackdraft.Nodes.Add(splt_report);

        //                }
        //                else if ((id_process >= (int)SKProcess.ModelBackDraft_Completed) && (id_process < (int)SKProcess.ModelBackDraft_PM_Approved))
        //                {
        //                    //for back checking
        //                    childnode = tvwbackcheck.Nodes.Add(splt_report);

        //                }
        //                else if ((id_process >= (int)SKProcess.QA) && (id_process < (int)SKProcess.Completed))
        //                {
        //                    //for RESERVED FOR FUTURE
        //                    childnode = tvwqa.Nodes.Add(splt_report);

        //                }
        //                //else if (id_process == (int)SKProcess.Held)
        //                //{
        //                //    //completed
        //                //    childnode = tvwheld.Nodes.Add(splt_report);
        //                //}
        //                //else if  (id_process == (int)SKProcess.Cancelled)
        //                //{
        //                //    //held
        //                //    childnode = tvwcancelled.Nodes.Add(splt_report);
        //                //}

        //                else if ((id_process == (int)SKProcess.Completed))
        //                {
        //                    //qa
        //                    childnode = tvwcompleted.Nodes.Add(splt_report);
        //                }
        //                //add node tag with stage name
        //                if ((id_process >= (int)SKProcess.Modelling_Completed) && (id_process <= (int)SKProcess.Completed))
        //                {
        //                    childnode.Tag = splt_mydata[0] + "|" + splt_mydata[1];
        //                    if (id_emp.ToString() == splt_mydata[3])
        //                        childnode.ForeColor = System.Drawing.Color.Blue;
        //                    else
        //                        childnode.ForeColor = System.Drawing.Color.Black;
        //                    if (id_process < (int)SKProcess.ModelBackDraft_PM_Approved)
        //                    {
        //                        int total = 0;
        //                        if (skWinLib.IsNumeric(splt_mydata[4]) == true)
        //                            total = Convert.ToInt32(splt_mydata[4]);

        //                        int val1 = 0;
        //                        if (skWinLib.IsNumeric(splt_mydata[5]) == true)
        //                            val1 = Convert.ToInt32(splt_mydata[5]);

        //                        int val2 = 0;
        //                        if (skWinLib.IsNumeric(splt_mydata[6]) == true)
        //                            val2 = Convert.ToInt32(splt_mydata[6]);

        //                        child_dashdata = childnode.Nodes.Add("Total:" + total.ToString());
        //                        if (id_process < (int)SKProcess.ModelChecking_Completed)
        //                        {
        //                            if (val1 != 0)
        //                                child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());
        //                            if (val2 != 0)
        //                                child_dashdata = childnode.Nodes.Add("No Comment:" + val2.ToString());
        //                        }
        //                        else if (id_process < (int)SKProcess.ModelBackDraft_Completed)
        //                        {
        //                            if (val1 != 0)
        //                                child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());
        //                            if (val2 != 0)
        //                                child_dashdata = childnode.Nodes.Add("Pending Cat." + val2.ToString());
        //                        }
        //                        else if (id_process < (int)SKProcess.ModelBackCheck_Completed)
        //                        {
        //                            if (val1 != 0)
        //                                child_dashdata = childnode.Nodes.Add("Pending:" + val1.ToString());

        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                //hope there is ||| more than one instance
        //                //next version changes
        //            }
        //        }


        //    }
        //}
        private ArrayList get_Report_Breakup_DashBoard(ArrayList mydata, int id_ncm_process, int colid)
        {
            ArrayList myreturn = new ArrayList();
            int ct = 0;
            int nullct = 0;
            int nocmmtct = 0;
            for (int i = 0; i < mydata.Count; i++)
            {
                string[] splt_data = mydata[i].ToString().Replace(sk_seperator, "|").Split(new Char[] { '|' });

                if (Convert.ToInt32(splt_data[0].ToString()) - id_ncm_process == 0)
                {
                    ct++;
                    if (splt_data.Length >= colid)
                    {

                        if (splt_data[colid].ToString().Trim().Length == 0)
                        {
                            //null count
                            nullct++;
                        }
                        else if (splt_data[colid].ToString().Trim().ToUpper() == sk_nocomment.Trim().ToUpper())
                        {

                            //no comments
                            nocmmtct++;
                        }

                    }


                }
            }
            myreturn.Add(ct.ToString());
            myreturn.Add(nullct.ToString());
            myreturn.Add(nocmmtct.ToString());
            return myreturn;
        }
        private void Load_Employee([Optional] string sql)
        {
            //fetch username,name, email
            //	297	2021-11-25 12:15:08	18	18	2021-11-25 12:15:08	sks360@esskaystructures.com	P Vijayakumar	$2a$10$grQgpdSNm8hgyNxY0zeVP.94hVxS83entGNwUn/7T5U7FhToJJ0UG	SKS360
            if (sql == null)
                mysql = "select username,name,id from epmatdb.users";
            else
                mysql = sql;

            mycmd.CommandText = mysql;

            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();

            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            cmbchkid.Items.Clear();
            empmast.Clear();
            while (myrdr.Read())
            {

                cmbchkid.Items.Add(myrdr[0].ToString().ToUpper());
                empmast.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString() + "|" + myrdr[2].ToString());
            }
            myrdr.Close();
            //set sorted based on UI setting
            if (cmbchkid.Sorted == true)
                empmast.Sort();
        }

        private void UpdateUserInterface_Fromlastsave()
        {
            try
            {
                //fetch jobnumber and  project name
                Load_Project();

                //fetch employee details
                //SELECT user.userId, user.name from roles.roleid FROM roles JOIN  user ON Roles.userid = user.userId where roleid = 1;
                //Load_Employee();

                //Update Dash Board
                Update_DashBoard_TreeView();

                //Get SK Project Number from Project Property Now and in future cmbproject will be locked once confirmed MD.
                //cmbchkid.Text = skWinLib.username;
                //lblxsuser.Text = skWinLib.username;
                //13042022
                //lblrole.Text = "Checker";
                //emp_role = lblrole.Text;



                //load Category
                mysql = "select Category,id_Category from epmatdb.Categorymast";
                myconn = new MySqlConnection(connString);
                if (myconn.State == ConnectionState.Closed)
                    myconn.Open();
                mycmd.CommandText = mysql;
                mycmd.Connection = myconn;
                myrdr = mycmd.ExecuteReader();



                qmsCategory.Clear();
                while (myrdr.Read())
                {
                    qmsCategory.Add(myrdr[0].ToString());
                }
                myrdr.Close();


                //load Severity
                mysql = "select severity,id_severity from epmatdb.Severitymast";
                //mysql = "select idempmast from ncm.empmast where role = 'checker' ";
                mycmd.CommandText = mysql;
                mycmd.Connection = myconn;
                myrdr = mycmd.ExecuteReader();
                qmsSeverity.Clear();
                while (myrdr.Read())
                {
                    qmsSeverity.Add(myrdr[0].ToString());
                }
                myrdr.Close();

                SetUserInteface_DashBoard();
                //SetUserInteface_Tabpage("DashBoard");

                //set the datagrid column values applicable for that roles only one time setting during form load
                ////based on role set the User Interface
                //FirstTimeSetting_UserInterface(lblrole.Text, id_process);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL error: " + ex.Message);
            }




            /*
            
           
            myconn.Close();
            */



        }


        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void SetRole(string department, string role)
        {

            if (department.ToUpper().Contains("DRAWINGOFFICE") && role.ToUpper().Contains("CHECK"))
                lblrole.Text = "Checker";
            else if (department.ToUpper().Contains("DRAWINGOFFICE") && role.ToUpper().Contains("MODEL"))
                lblrole.Text = "Modeler";
            else if (department.ToUpper().Contains("DRAWINGOFFICE") && role.ToUpper().Contains("DETAIL"))
                lblrole.Text = "Detailer";
            else if ((role.ToUpper().Contains("LEAD")) || (role.ToUpper().Contains("MANAGER")))
                lblrole.Text = "Lead";
            else
                lblrole.Text = "";
        }

        private void cmbchkid_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblsbar2.Text = "cmbchkid_SelectedIndexChanged";
            //int tmp = empmast.IndexOf(cmbchkid.Text);            
            string[] empinfo = empmast[cmbchkid.SelectedIndex].ToString().Split(new Char[] { '|' });
            id_emp = Convert.ToInt32(empinfo[0]);
            id_empname = empinfo[2];
            lblrole.Text = empinfo[2];
            lblxsuser.Text = id_empname;
            //lblxsuser.Tag = id_emp;
            SetRole("DRAWINGOFFICE", lblrole.Text);
        }


        //private void cmbprojid_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    lblsbar2.Text = "cmbprojid_SelectedIndexChanged";
        //    string[] projectinfo = projectmast[cmbprojid.SelectedIndex].ToString().Split(new Char[] { '|' });
        //    txtclient.Text = projectinfo[1];
        //    txtprojname.Text = projectinfo[2];
        //    //cmbprojid.Tag = projectinfo[3];
        //    id_project = Convert.ToInt32(projectinfo[3]);
        //    this.Text = this.Text + "      " + cmbprojid.Text + " / " + txtclient.Text + " / " + txtprojname.Text;
        //}




        //private void tbpgUI_Click(object sender, EventArgs e)
        //{
        //    lblsbar2.Text = "tbpgUI_Click";
        //    //bool chkstate = false;
        //    //string colname = "";
        //    //for (int i = 0; i < chklsthdr.Items.Count; i++)
        //    //{

        //    //    if (chklsthdr.GetItemCheckState(i) == CheckState.Unchecked)
        //    //    {

        //    //        chkstate = false;
        //    //    }
        //    //    if (chklsthdr.GetItemCheckState(i) == CheckState.Checked)
        //    //    {

        //    //        chkstate = true;

        //    //    }
        //    //    colname = chklsthdr.Items[i].ToString();
        //    //    this.dgQMS.Columns[colname].Visible = chkstate;
        //    //}
        //    lblsbar2.Text = "";
        //}




        //////private void loadColumnHeader()
        //////{

        //////    string qmsAttributes = "GUID,Observation,Category,Severity,Owner,Agreed,BackCheckedby";
        //////    //string tsAttributes = "Type,Mark,Phase,IsMainPart,Name,Profile,Material,Finish,AssyPrefix,AssyStart,PartPrefix,PartStart,ABM,User1,User2,User3,User4";
        //////    string tsAttributes = "Type,Assembly_Pos,Phase.Name,Name,Profile,Material,Finish,Assembly_Prefix,Part_Prefix,User_field_1,User_field_2,User_field_3,User_field_4";

        //////    string[] split = qmsAttributes.Split(new Char[] { ',' });
        //////    int rowid = 0;
        //////    int colid = 0;
        //////    chklsthdr.Items.Clear();
        //////    //foreach (string s in split)
        //////    //{
        //////    //    chklsthdr.Items.Add(s);
        //////    //    chklsthdr.SetItemChecked(colid, true);
        //////    //    colid++;
        //////    //}

        //////    split = tsAttributes.Split(new Char[] { ',' });
        //////    foreach (string s in split)
        //////    {
        //////        chklsthdr.Items.Add(s);
        //////        chklsthdr.SetItemChecked(colid, false);
        //////        colid++;
        //////        dgQMS.Columns.Add(s, s);

        //////    }
        //////    //dgQMS.Columns.HeaderText = "S.No";
        //////    //dgQMS.Rows[0].HeaderCell.Value = "1";
        //////    dgQMS.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        //////    CategorySeveritySetting(0);
        //////}
        private void UpdateVisibleTable()
        {
            if (chkTSAttribute.CheckedItems.Count > 0)
            {

                foreach (DataGridViewColumn myDataGridViewColumn in dgqc.Columns)
                {
                    string dgcolname = myDataGridViewColumn.Name;
                    myDataGridViewColumn.Visible = false;
                    foreach (int indexChecked in this.chkTSAttribute.CheckedIndices)
                    {
                        string chkcolumn = chkTSAttribute.Items[indexChecked].ToString();
                        if (chkcolumn.ToUpper() == dgcolname.ToUpper())
                        {
                            myDataGridViewColumn.Visible = true;
                            break;
                        }
                    }

                }

            }
            else
            {
                foreach (DataGridViewColumn myDataGridViewColumn in dgqc.Columns)
                {
                    myDataGridViewColumn.Visible = false;

                }
            }

        }
        private void tabNCM_Selected(object sender, TabControlEventArgs e)
        {


            if (e.TabPageIndex == 0)
            {




            }
            else if (e.TabPageIndex == 1)
            {
                //UpdateVisibleTable();



            }


        }

        //private void tbpgSetting_Click(object sender, EventArgs e)
        //{
        //    SetUserInteface_Tabpage("DASHBOARD");
        //}
        //private void SelectDrawingObjects()
        //{

        //    TSD.DrawingObjectEnumerator MyEnum = new DrawingHandler().GetDrawingObjectSelector().GetSelected();
        //    int ct = 0;
        //    lblsbar2.Text = "SelectDrawingObjects";

        //    while (MyEnum.MoveNext())
        //    {
        //        string s_tmp = "";
        //        TSD.DrawingObject current = MyEnum.Current;
        //        string drgname = current.GetDrawing().Name;
        //        string drgtype = current.GetType().Name;
        //        dgqc.Rows.Add(drgname);
        //        int rowid = dgqc.Rows.Count - 1;
        //        dgqc.Rows[rowid].Cells["TeklaSKtype"].Value = drgtype;
        //        dgqc.Rows[rowid].HeaderCell.Value = dgqc.Rows.Count.ToString();
        //        ct++;
        //    }
        //}

        //private void newentry()
        //{

        //    DataTable table = new DataTable();
        //    table.Columns.Add("Name".ToString());
        //    table.Columns.Add("Color".ToString());
        //    DataRow dr = table.NewRow();
        //    dr["Name"] = "Mike";
        //    dr["Color"] = "blue";
        //    table.Rows.Add(dr);


        //}
        private void SelectModelObjects(int option = 1)
        {
            DataTable mytable = new DataTable();
            foreach (DataGridViewColumn mycolumn in dgqc.Columns)
            {
                if (mycolumn.Visible == true)
                    mytable.Columns.Add(mycolumn.Name);
            }


            //dgqc.Visible = false;
            int ct = 0;
            lblsbar2.Text = "SelectModelObjects";
            TSM.ModelObjectEnumerator MyEnum = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
            ArrayList myguidlist = new ArrayList();
            //TSD.DrawingObjectEnumerator drgMyEnum = new TSD.UI.DrawingObjectSelector.g

            while (MyEnum.MoveNext())
            {
                DataRow mydr = mytable.NewRow();
                TSM.ModelObject current = MyEnum.Current;
                mydr["guid"] = current.Identifier.GUID.ToString();
                mytable.Rows.Add(mydr);

                //myguidlist.Add(current.Identifier.GUID.ToString());

            }


            dgqc.DataSource = mytable;
            //dgqc.Visible = true;

            //refresh
            Update_DashBoard_TreeView();
        }

        private void SelectModelObjects(string option = "")
        {

            skWinLib.DataGridView_Setting_Before(dgqc);
            //dgqc.Visible = false;
            int ct = 0;
            lblsbar2.Text = "SelectModelObjects";
            TSM.ModelObjectEnumerator MyEnum = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
            //check for 500 + object and inform user this will take few minits
            bool manyObjectFlag = true;
            int totselct = MyEnum.GetSize();
            if (totselct >= 500)
            {
                if (MessageBox.Show(MyEnum.GetSize().ToString() + " selected will take few mints." + "\n\nCan you re-confirm to proceed futher ...?", msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    manyObjectFlag = false;
            }
            if (manyObjectFlag == true)
            {
                //TSD.DrawingObjectEnumerator drgMyEnum = new TSD.UI.DrawingObjectSelector.g


                //check for additional tekla field is checked on 
                TeklaAdditionalColumnFlag = IsTeklaAdditionalEnabled();

                //bool TeklaAdditionalTimeFlag = true;
                //bool speedupflag = false;
                while (MyEnum.MoveNext())
                {

                    TSM.ModelObject current = MyEnum.Current;
                    dgqc.Rows.Add(current.Identifier.GUID.ToString());
                    int rowid = dgqc.Rows.Count - 1;
                    DataGridViewRow myrow = dgqc.Rows[rowid];
                    if (chkPerformance.Checked == false)
                    {
                        //get type
                        if (dgqc.Columns["TeklaSKtype"].Visible == true)
                        {
                            string tsType = current.GetType().Name.ToString();
                            ///tsType = tsType.Replace("Tekla.Structures.Model.", "");
                            dgqc.Rows[dgqc.Rows.Count - 1].Cells["TeklaSKtype"].Value = tsType;
                        }
                    }




                    //if (TeklaAdditionalColumnFlag == true)
                    //{
                    //    
                    //    if (totselct >= 500)
                    //    {
                    //        if (speedupflag == false)
                    //        {
                    //            speedupflag = true;
                    //            if (MessageBox.Show(totselct.ToString() + " selected will take more time." + ".\n\n\nCan you re-confirm you need addtional tekla information..?", msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    //                TeklaAdditionalTimeFlag = true;
                    //            else
                    //                TeklaAdditionalTimeFlag = false;
                    //        }
                    //    }
                    //}
                    //if (TeklaAdditionalTimeFlag == true)
                    //{

                    //}

                    if (chkPerformance.Checked == false)
                    {
                        if (TeklaAdditionalColumnFlag == true)
                        {

                            update_Tekla_AdditionalAttributes(current, myrow);
                        }
                    }
                    //if (TeklaAdditionalColumnFlag == true)
                    //{
                    //    if (TeklaAdditionalTimeFlag == true)
                    //    {
                    //        DataGridViewRow myrow = dgqc.Rows[rowid];
                    //        update_Tekla_AdditionalAttributes(current, myrow);
                    //    }
                    //}
                    ct++;

                }
            }

            skWinLib.DataGridView_Setting_After(dgqc);

            //refresh
            Update_DashBoard_TreeView();
        }

        private void update_Tekla_AdditionalAttributes(TSM.ModelObject current, DataGridViewRow myrow)
        {
            int colid = 0;
            foreach (DataGridViewColumn mycolumn in dgqc.Columns)
            {
                if (mycolumn.Visible == true)
                {
                    if (mycolumn.Name.Contains("TeklaSK") == true)
                    {
                        string s_tmp = "";
                        current.GetReportProperty(mycolumn.Name.Replace("TeklaSK", "").ToUpper().Trim(), ref s_tmp);
                        if (s_tmp.Trim().Length >= 1)
                            myrow.Cells[colid].Value = s_tmp;
                    }
                }
                colid = colid + 1;
            }
        }

        //private void SelectModelObjects()
        //{
        //    int ct = 0;
        //    lblsbar2.Text = "SelectModelObjects";
        //    TSM.ModelObjectEnumerator MyEnum = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
        //    //TSD.DrawingObjectEnumerator drgMyEnum = new TSD.UI.DrawingObjectSelector.g
        //    while (MyEnum.MoveNext())
        //    {

        //        TSM.ModelObject current = MyEnum.Current;
        //        dgqc.Rows.Add(current.Identifier.GUID.ToString());
        //        int rowid = dgqc.Rows.Count - 1;
        //        int colid = 1;
        //        foreach (DataGridViewColumn mycolumn in dgqc.Columns)
        //        {
        //            if (mycolumn.Name.ToUpper().Trim() != "GUID")
        //            {
        //                string s_tmp = "";
        //                if (mycolumn.Name.ToUpper().Trim() == "TYPE")
        //                {
        //                    s_tmp = current.GetType().Name.ToString();

        //                }
        //                else
        //                {
        //                    if (mycolumn.Visible == true)
        //                    {
        //                        bool flag = current.GetReportProperty(mycolumn.Name.ToUpper().Trim(), ref s_tmp);
        //                        if (flag == false)
        //                            s_tmp = "";
        //                    }

        //                }
        //                if (s_tmp != "")
        //                    dgqc.Rows[rowid].Cells[colid].Value = s_tmp;
        //                colid = colid + 1;

        //            }

        //        }
        //        dgqc.Rows[rowid].HeaderCell.Value = dgqc.Rows.Count.ToString();

        //        ct++;
        //    }

        //lblsbar2.Text = "SelectModelObjects";
        //try
        //{


        //    //Set up model connection w/TS
        //    Model MyModel = new Model();
        //    if (MyModel.GetConnectionStatus() &&
        //        MyModel.GetInfo().ModelPath != string.Empty)
        //    {
        //        TransformationPlane CurrentTP = MyModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
        //        MyModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());


        //        //var objects = MyModel.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.BEAM);
        //        ModelObjectEnumerator MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
        //        int ct = 1;
        //        string mymark = "";
        //        string phname = "";
        //        string Observation = "";
        //        DataGridViewComboBoxCell Category = new DataGridViewComboBoxCell();
        //        DataGridViewComboBoxCell Severity = new DataGridViewComboBoxCell();
        //        string Owner = "";
        //        DataGridViewCheckBoxCell Agreed = new DataGridViewCheckBoxCell();
        //        DataGridViewCheckBoxCell BackCheckedby = new DataGridViewCheckBoxCell();
        //        string Type = "";
        //        string Mark = "";
        //        string Phase = "";
        //        string IsMainPart = "";
        //        string Name = "";
        //        string Profile = "";
        //        string Material = "";
        //        string Finish = "";
        //        string AssyPrefix = "";
        //        string AssyStart = "";
        //        string PartPrefix = "";
        //        string PartStart = "";
        //        string ABM = "";
        //        string User1 = "";
        //        string User2 = "";
        //        string User3 = "";
        //        string User4 = "";                  
        //        while (MyEnum.MoveNext())
        //        {

        //            mymark = "";
        //            phname = "";
        //            Observation = "";
        //            Category = null;
        //            Severity = null;
        //            Owner = "";
        //            Agreed = null;
        //            BackCheckedby = null;
        //            Type = "";
        //            Mark = "";
        //            Phase = "";
        //            IsMainPart = "";
        //            Name = "";
        //            Profile = "";
        //            Material = "";
        //            Finish = "";
        //            AssyPrefix = "";
        //            AssyStart = "";
        //            PartPrefix = "";
        //            PartStart = "";
        //            ABM = "";
        //            User1 = "";
        //            User2 = "";
        //            User3 = "";
        //            User4 = "";

        //            TSM.ModelObject current = MyEnum.Current;
        //            addNewlySelectedModelObject(current);
        //            ct++;
        //        }


        //        // original UCS
        //        MyModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);

        //        //label2.Text = listView1.Items.Count.ToString() + " Object(s) listed";
        //    } // end of if that checks to see if there is a model and one is open

        //    else
        //    {
        //        //If no connection to model is possible show error message
        //        MessageBox.Show("Tekla Structures not open or model not open.",
        //            "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}  //end of try loop

        //catch (Exception ex)
        //{
        //    MessageBox.Show("Reading tekla objects. Original error: " + ex.Message);
        //}
        //lblsbar2.Text = "";

        private void addNewlySelectedModelObject(TSM.ModelObject ModelObject)
        {


            DataGridViewComboBoxCell Category = new DataGridViewComboBoxCell();
            //DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Category"]);
            Category.DataSource = qmsCategory;
            DataGridViewComboBoxCell Severity = new DataGridViewComboBoxCell();
            DataGridViewCheckBoxCell Agreed = new DataGridViewCheckBoxCell();
            DataGridViewCheckBoxCell BackCheckedby = new DataGridViewCheckBoxCell();

            string mymark = "";
            string Observation = "";
            //string Category = "";
            //string Severity = "";
            string Owner = "";
            //string Agreed = "";
            //string BackCheckedby = "";
            string Type = "";
            string Mark = "";
            string Phase = "";
            int IsMainPart = 0;
            string Name = "";
            string Profile = "";
            string Material = "";
            string Finish = "";
            string AssyPrefix = "";
            int AssyStart = 0;
            string PartPrefix = "";
            int PartStart = 0;
            string ABM = "";
            string User1 = "";
            string User2 = "";
            string User3 = "";
            string User4 = "";


            //if (ModelObject is TSM.ContourPlate)
            //    Type = "ContourPlate";
            //else
            //    Type = "Beam";

            string mytype = ModelObject.GetType().Name;

            var guid = ModelObject.Identifier.GUID;




            ModelObject.GetReportProperty("ASSEMBLY_POS", ref Mark);
            ModelObject.GetReportProperty("PHASE.NAME", ref Phase);
            ModelObject.GetReportProperty("MAIN_PART", ref IsMainPart);
            ModelObject.GetReportProperty("NAME", ref Name);
            ModelObject.GetReportProperty("PROFILE", ref Profile);
            ModelObject.GetReportProperty("MATERIAL", ref Material);
            ModelObject.GetReportProperty("FINISH", ref Finish);
            ModelObject.GetReportProperty("ASSEMBLY_PREFIX", ref AssyPrefix);
            ModelObject.GetReportProperty("ASSEMBLY_START_NUMBER", ref AssyStart);
            ModelObject.GetReportProperty("PART_PREFIX", ref PartPrefix);
            ModelObject.GetReportProperty("PART_START_NUMBER", ref PartStart);
            if (Phase == null || Phase == "")
            {
                if (ModelObject is TSM.Weld)
                {
                    TSM.Weld myWeld = ModelObject as TSM.Weld;
                    myWeld.GetPhase(out Phase ph);
                    Phase = ph.PhaseName;
                    Name = myWeld.TypeAbove.ToString() + "/" + myWeld.TypeBelow.ToString();
                    Profile = myWeld.ShopWeld.ToString();
                    Material = myWeld.Preparation.ToString();
                    Finish = myWeld.ReferenceText.ToString();
                }
                else if (ModelObject is TSM.Detail)
                {
                    TSM.Detail myDetail = ModelObject as TSM.Detail;
                    myDetail.GetPhase(out Phase ph);
                    Phase = ph.PhaseName;
                }
                else if (ModelObject is TSM.Connection)
                {
                    TSM.Connection myComponent = ModelObject as TSM.Connection;
                    myComponent.GetPhase(out Phase ph);
                    Phase = ph.PhaseName;
                }
            }


            ABM = "";
            User1 = "";
            User2 = "";
            User3 = "";
            User4 = "";
            ModelObject.GetReportProperty("ABM", ref ABM);
            if (ABM != "")
                ModelObject.GetReportProperty("PRELIM_MARK", ref ABM);

            ModelObject.GetReportProperty("USER_FIELD_1", ref User1);
            ModelObject.GetReportProperty("USER_FIELD_2", ref User2);
            ModelObject.GetReportProperty("USER_FIELD_3", ref User3);
            ModelObject.GetReportProperty("USER_FIELD_4", ref User4);
            dgqc.Rows.Add(guid);
            int rowidx = dgqc.Rows.Count - 1;
            int colid = 6;
            CategorySeveritySetting(rowidx);
            dgqc.Rows[rowidx].HeaderCell.Value = (rowidx + 1).ToString();

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = mytype;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Mark;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Phase;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = IsMainPart;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Name;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Profile;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Material;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = Finish;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = AssyPrefix;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = AssyStart;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = PartPrefix;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = AssyStart;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = PartStart;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = User1;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = User2;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = User3;

            colid++;
            dgqc.Rows[rowidx].Cells[colid].Value = User4;


            UpdateCellAccess(lblrole.Text, rowidx);
            //lock cells and update back color based on role




        }
        private void UpdateCellAccess(string role, int rowidx)
        {

            for (int colidx = 1; colidx < dgqc.Columns.Count; colidx++)
            {
                dgqc.Columns[colidx].ReadOnly = true;

            }

            if (dgqc.Rows.Count >= 1)
            {
                if (role.ToUpper() == "CHECKER")
                {
                    dgqc.Columns["Observation"].ReadOnly = false;
                    dgqc.Rows[rowidx].Cells[dgqc.Columns["Observation"].Index].Style.BackColor = System.Drawing.Color.LightBlue;
                    //On back checking stage this to be enabled
                    //dgQMS.Columns["BackCheckedby"].ReadOnly = false;
                    //dgQMS.Rows[rowidx].Cells[dgQMS.Columns["BackCheckedby"].Index].Style.BackColor = System.Drawing.Color.LightGray;

                }
                if (role.ToUpper() == "MODELER")
                {
                    dgqc.Columns["Agreed"].ReadOnly = false;
                    //dgQMS.Columns["Remarks"].ReadOnly = false;
                }

            }
            dgqc.Columns["Observation"].ReadOnly = false;

        }

        private int getListIndex(ArrayList List, string SearchItem)
        {
            int id = 0;
            foreach (string chkCategory in List)
            {
                if (SearchItem.Trim().ToUpper() == chkCategory.Trim().ToUpper())
                {
                    return id;

                }
                id = id + 1;
            }
            return -1;
        }

        private void CategorySeveritySetting(int rowindex)
        {
            bool chk1 = false;
            bool chk2 = false;
            if (dgqc.RowCount >= 1)
            {
                skWinLib.DataGridView_Setting_Before(dgqc);

                for (int i = 0; i < dgqc.Columns.Count; i++)
                {
                    DataGridViewRow row = this.dgqc.Rows[rowindex];
                    if (dgqc.Columns[i].Name.ToUpper() == "Category".ToUpper())
                    {
                        DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Category"]);
                        cell.DataSource = qmsCategory;
                        chk1 = true;
                    }
                    if (dgqc.Columns[i].Name.ToUpper() == "Severity".ToUpper())
                    {
                        DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Severity"]);
                        cell.DataSource = qmsSeverity;
                        chk2 = true;
                    }
                    if ((chk1 == true) && (chk2 == true))
                        break;
                }


                skWinLib.DataGridView_Setting_After(dgqc);
            }

        }
        private void addNCMOptioninDataGrid(DataGridViewRowsAddedEventArgs e)
        {
            lblsbar2.Text = "addNCMOptioninDataGrid";
            if (e.RowIndex >= 0)
            {
                int rowindex = e.RowIndex;
                CategorySeveritySetting(rowindex);

            }

            //lblsbar2.Text = "";

        }
        private void dgQMS_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            lblsbar2.Text = "dgQMS_RowsAdded";
            int rowindex = e.RowIndex;
            if (e.RowIndex >= 0)
                CategorySeveritySetting(rowindex);

            dgqc.Rows[rowindex].HeaderCell.Value = (rowindex + 1).ToString();
            //lblsbar2.Text = "";
        }

        //private bool IsDrawingActive()
        //{

        //    TSD.DrawingHandler dh = new TSD.DrawingHandler();
        //    TSD.Drawing activeDrawing = null;
        //    if (dh.GetConnectionStatus())
        //    {

        //        bool m = dh.GetDrawings().SelectInstances;
        //        //TSD.DrawingEnumerator drawingListEnumerator = dh.GetDrawings();
        //        TSD.DrawingEnumerator drawingListEnumerator = dh.GetDrawings();
        //        while (drawingListEnumerator.MoveNext())
        //        {
        //            TSD.Drawing current = drawingListEnumerator.Current;
        //            string drgname = current.Name;
        //            string drgtype = current.GetType().Name;
        //            dgqc.Rows.Add(drgname);
        //            int rowid = dgqc.Rows.Count - 1;
        //            dgqc.Rows[rowid].Cells["TeklaSKtype"].Value = drgtype;
        //            dgqc.Rows[rowid].HeaderCell.Value = dgqc.Rows.Count.ToString();

        //        }
        //        if (dgqc.Rows.Count >= 1)
        //            return true;
        //        //{ }
        //        //    if (drawing.GetType().Equals(typeof(TSD.CastUnitDrawing)))
        //        //    {

        //        //        TSD.UI.DrawingSelector mydrg = dh.GetDrawingSelector();

        //        //if (mydrg !=null)
        //        //{

        //        //    TSD.DrawingObjectEnumerator MyEnum = new DrawingHandler().GetDrawingObjectSelector().GetSelected();
        //        //    int ct = 0;
        //        //    lblsbar2.Text = "SelectDrawingObjects";

        //        //    while (MyEnum.MoveNext())
        //        //    {

        //        //        TSD.DrawingObject current = MyEnum.Current;
        //        //        string drgname = current.GetDrawing().Name;
        //        //        string drgtype = current.GetType().Name;
        //        //        dgqc.Rows.Add(drgname);
        //        //        int rowid = dgqc.Rows.Count - 1;
        //        //        dgqc.Rows[rowid].Cells["TeklaSKtype"].Value = drgtype;
        //        //        dgqc.Rows[rowid].HeaderCell.Value = dgqc.Rows.Count.ToString();
        //        //        ct++;
        //        //    }

        //        //return true;
        //        //}     
        //        activeDrawing = dh.GetActiveDrawing();
        //        if (activeDrawing != null)
        //        {

        //            return true;
        //        }

        //    }
        //    return false;
        //}

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (isowner == true)
            {
                //tpgqc.Tag = cmbqmsdblist.Text;;
                //guid's can be added during new / inprocess only hence allowed until process is below 20
                if (SK_id_process < (int)SKProcess.ModelChecking_PM_Approved)
                {
                    DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Select_TS_Object", "", "Pl.Wait,selecting TS object's", "btnAdd_Click", lblsbar1, lblsbar2);

                    skWinLib.DataGridView_Setting_Before(dgqc);
                    //if (IsDrawingActive() == false)
                    SelectModelObjects("Add");
                    skWinLib.DataGridView_Setting_After(dgqc);
                    if (dgqc.RowCount >= 1)
                    {
                        //update row header
                        skWinLib.updaterowheader(dgqc);
                        btnSave.Enabled = true;
                    }
                    else
                    {
                        btnSave.Enabled = false;
                    }
                    chkzoom.Enabled = btnSave.Enabled;
                    btnDelete.Enabled = btnSave.Enabled;
                    btnclear.Enabled = btnSave.Enabled;

                    ////update row header
                    //skWinLib.updaterowheader(dgqc);


                    skWinLib.setCompletion(skApplicationName, skApplicationVersion, "Select_TS_Object", "", "", "btnAdd_Click", lblsbar1, lblsbar2, s_tm);

                }
                //else
                //{
                //    //only checker can select the model guid's and push to server
                //}
            }
            else
            {
                //since not owner no action required
                MessageBox.Show("You are not the owner for report " + tpgqc.Text + "  hence aborting now!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "btnclear_Click";
            lblsbar1.Text = "Processing clear object(s).....";
            dgqc.Rows.Clear();
            //tpgqc.Tag = cmbqmsdblist.Text;;

            //based on row count set control visiblity
            SetControlVisible_QC_DataGridRowCount();

            lblsbar1.Text = "Ready";
        }


        //        for (int rowindex = 0; rowindex<dgqc.Rows.Count; rowindex++)
        //                        {
        //                            DataGridViewRow row = this.dgqc.Rows[rowindex];
        //        //check observation is not null or empty space
        //        obser = string.Empty;

        //                            if (row.Cells["observation"].Value != null)
        //                                obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

        //                            guid = string.Empty;
        //                            //get variable value from data grid
        //                            if (row.Cells["guid"].Value != null)
        //                                guid = row.Cells["guid"].Value.ToString();


        //                            guid_type = string.Empty;
        //                            if (row.Cells["TeklaSKtype"].Value != null)
        //                                guid_type = row.Cells["TeklaSKtype"].Value.ToString();

        //                            idqc = -1;
        //                            if (row.Cells["idqc"].Value != null)
        //                            {
        //                                string tmp = row.Cells["idqc"].Value.ToString();
        //                                if (skWinLib.IsNumeric(tmp) == true)
        //                                    idqc = Convert.ToInt32(tmp);
        //                            }
        //    remark = string.Empty;
        //                            if (row.Cells["remark"].Value != null)
        //                                remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

        //                            int db_bdisagree = 0;
        //                            if (row.Cells["isagreed"].Value != null)
        //                            {
        //                                string agreed = row.Cells["isagreed"].Value.ToString();
        //                                if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
        //                                    db_bdisagree = 1;
        //                                else
        //                                    db_bdisagree = 0;

        //                            }




        //int db_bdcat = -1;
        //if (row.Cells["category"].Value != null)
        //{
        //    string category = row.Cells["category"].Value.ToString();
        //    db_bdcat = getListIndex(qmsCategory, category);
        //    //DataGridViewComboBoxCell mycat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        //    ////string chk = mycat.Items.IndexOf(category);
        //    //if (skWinLib.IsNumeric(category) == true)
        //    //    db_bdcat = 1;
        //}


        //int db_dbseverity = -1;
        //if (row.Cells["severity"].Value != null)
        //{
        //    string severity = row.Cells["severity"].Value.ToString();
        //    db_dbseverity = getListIndex(qmsSeverity, severity);
        //}

        //int db_backcheck = -1;
        //if (row.Cells["backcheck"].Value != null)
        //{
        //    string agreed = row.Cells["backcheck"].Value.ToString();
        //    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
        //        db_backcheck = 1;
        //    else
        //        db_backcheck = 0;
        //}
        ////if (idqc == 1211)
        ////{
        ////    string cc = "";
        ////}
        //if (idqc == -1)
        //{
        //    //insert fresh record into my sql
        //    int flag = ncm_insert(guid, guid_type, obser, remark, id_stage);
        //    //int flag = ncm_insert(paraflag, guid, guid_type, obser, remark, skWinLib.systemname,  id_stage, timestamp); 
        //    //for status log
        //    if (flag != -1)
        //    {
        //        ins_ct++;
        //        paraflag = true;
        //    }
        //    else
        //        err_ct++;
        //}
        //else
        //{

        //    //already record exist need to update changed field value
        //    //for better performance check with earlier stored datagrid against idqc and if difference found update those column only                    
        //    string db_cat = string.Empty;
        //    int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
        //    //int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
        //    //for status log
        //    if (flag != -1 && flag != 0)
        //        updt_ct++;
        //    else
        //        err_ct++;

        //}
        private void btnSave_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "Save", "", "Pl.Wait,Saving...", "btnSave_Click", lblsbar1, lblsbar2);
            string msg = string.Empty;
            DateTime chk_dt = DateTime.Now;
            DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
            long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

            //check whether already data avaialbe in mirror if not then perform insert for newly added guid's
            int newguidct = 0;


            if (dgqcmirror.RowCount >= 1)
            {
                ArrayList mysqldata = getDifference_Datagrid();
                ArrayList mycolumnlist = mysqldata[0] as ArrayList;
                ArrayList mysql_columnname = mysqldata[1] as ArrayList;
                ArrayList new_data = mysqldata[2] as ArrayList;
                ArrayList modify_data = mysqldata[3] as ArrayList;
                ArrayList beforesave = mysqldata[4] as ArrayList;

                //for new data
                if (new_data.Count >= 1)
                {

                    //bulk insert currently set for 5000 record at one time
                    newguidct = bulk_insert(SK_id_report, SK_id_process, 5000, new_data, id_emp, timestamp);

                    //int total = bulk_insert(id_stage, timestamp, 5000);
                    msg = newguidct.ToString() + " added.";
                    lblsbar1.Text = msg;

                }

                //for update
                if (modify_data.Count >= 1)
                {
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    msg = msg + " " + Update_ModifiedData(mysql_columnname, modify_data, id_stage, timestamp);
                    //msg = msg + " " + Update_ModifiedData(mysql_columnname, modify_data, timestamp);

                }

                if (newguidct >= 1)
                {
                    ////after saving fetch the data and update user interface
                    Fetch_Model_Data();

                    //update row header
                    skWinLib.updaterowheader(dgqc);
                }

            }
            else
            {
                //bulk insert currently set for 5000 record at one time
                newguidct = bulk_insert(SK_id_report, SK_id_process, 5000, dgqc, id_emp, timestamp);
                if (newguidct >= 1)
                {
                    //int total = bulk_insert(id_stage, timestamp, 5000);
                    msg = newguidct.ToString() + " added.";

                    //function update idqc for the newly added records in the current data grid user interface which are all blank;
                    Update_idqc_NewRecord(SK_id_report, id_emp, timestamp);

                    //update row header
                    skWinLib.updaterowheader(dgqc);
                }
                lblsbar1.Text = msg;
            }

            if (msg != "")
            {
                //refresh
                Update_DashBoard_TreeView();

                //mirror data
                MirrorTabPageData();

                //based on row count set control visiblity
                SetControlVisible_QC_DataGridRowCount();
            }
            else
                msg = "Nothing to Save!!!";
            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "Save", "", msg, "btnSave_Click", lblsbar1, lblsbar2, s_tm);
        }

        private int bulk_insert(int id_ncm_progress, string NewData)
        {
            string insert_statement = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress, chk_user_id, chk_dt) VALUES ";
            mycmd.CommandText = insert_statement + "(" + NewData + ")";
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            int insct = mycmd.ExecuteNonQuery();
            return insct;
        }

        private int bulk_insert(int id_report, int id_sk_process, int bulkrecord, ArrayList myData, int id_emp, long timestamp)
        {

            string insert_statement = "INSERT INTO epmatdb.ncm(guid, guid_type, mod_rmk, id_report, id_sk_process,  mod_user_id, mod_dt) VALUES ";
            string sql = ",";
            int rowct = 0;
            int insct = 0;
            foreach (DataGridViewRow row in myData)
            {

                //check observation is not null or empty space
                string guid = string.Empty;
                string guid_type = string.Empty;
                string remark = string.Empty;


                //get guid
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();

                //get guidtype
                if (row.Cells["TeklaSKtype"].Value != null)
                    guid_type = row.Cells["TeklaSKtype"].Value.ToString();

                //check idqc
                int idqc = -1;
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                        idqc = Convert.ToInt32(tmp);
                }

                if (row.Cells["remark"].Value != null)
                    remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

                //sql = sql + " , (" + id_project.ToString() + ",'" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_emp + ",'" + skWinLib.systemname + "'," + newreport_stage + "," + timestamp + ")";
                sql = sql + " , ('" + guid + "','" + guid_type + "','" + remark + "'," + id_report + "," + id_sk_process + "," + id_emp + "," + timestamp + ")";

                rowct++;

                //based on bulk record count perform insert 
                //for better performance
                if (rowct == bulkrecord)
                {
                    sql = sql.Replace(", ,", "");
                    mycmd.CommandText = insert_statement + sql;
                    mycmd.Connection = myconn;
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    insct = insct + mycmd.ExecuteNonQuery();
                    sql = ",";
                    rowct = 0;
                }
            }
            //replace ,, with ,
            sql = sql.Replace(", ,", "");
            mycmd.CommandText = insert_statement + sql;
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            insct = insct + mycmd.ExecuteNonQuery();
            return insct;
        }
        private int bulk_insert(int id_report, int id_sk_process, int bulkrecord, DataGridView myDataGridView, int id_emp, long timestamp)
        {

            string insert_statement = "INSERT INTO epmatdb.ncm(guid, guid_type, mod_rmk, id_report, id_sk_process,  mod_user_id, mod_dt) VALUES ";
            string sql = ",";
            int rowct = 0;
            int insct = 0;
            foreach (DataGridViewRow row in myDataGridView.Rows)
            {

                //check observation is not null or empty space
                string guid = string.Empty;
                string guid_type = string.Empty;
                string remark = string.Empty;


                //get guid
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();

                //get guidtype
                if (row.Cells["TeklaSKtype"].Value != null)
                    guid_type = row.Cells["TeklaSKtype"].Value.ToString();

                //check idqc
                int idqc = -1;
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                        idqc = Convert.ToInt32(tmp);
                }

                if (row.Cells["remark"].Value != null)
                    remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

                //sql = sql + " , (" + id_project.ToString() + ",'" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_emp + ",'" + skWinLib.systemname + "'," + newreport_stage + "," + timestamp + ")";
                sql = sql + " , ('" + guid + "','" + guid_type + "','" + remark + "'," + id_report + "," + id_sk_process + "," + id_emp + "," + timestamp + ")";

                rowct++;

                //based on bulk record count perform insert 
                //for better performance
                if (rowct == bulkrecord)
                {
                    sql = sql.Replace(", ,", "");
                    mycmd.CommandText = insert_statement + sql;
                    mycmd.Connection = myconn;
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    insct = insct + mycmd.ExecuteNonQuery();
                    sql = ",";
                    rowct = 0;
                }
            }
            //replace ,, with ,
            sql = sql.Replace(", ,", "");
            mycmd.CommandText = insert_statement + sql;
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            insct = insct + mycmd.ExecuteNonQuery();
            return insct;
        }

        //private int bulk_insert(int id_ncm_progress, int bulkrecord, DataGridView myDataGridView, int id_emp, long timestamp)
        //{

        //    string insert_statement = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress, chk_user_id, chk_dt) VALUES ";
        //    //string insert_statement = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress, chk_user_id, chk_dt) VALUES ";

        //    //string insert_statement = "INSERT INTO epmatdb.ncm(project_id,guid,guid_type,obser,chk_rmk,chk_user_id,chk_sys,id_ncm_progress,entry_date) VALUES ";
        //    string sql = ",";
        //    int rowct = 0;
        //    int insct = 0;
        //    foreach (DataGridViewRow row in myDataGridView.Rows)
        //    {

        //        //check observation is not null or empty space
        //        string obser = string.Empty;
        //        string guid = string.Empty;
        //        string guid_type = string.Empty;
        //        string remark = string.Empty;


        //        //get observation
        //        if (row.Cells["observation"].Value != null)
        //            obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

        //        //get guid
        //        if (row.Cells["guid"].Value != null)
        //            guid = row.Cells["guid"].Value.ToString();

        //        //get guidtype
        //        if (row.Cells["TeklaSKtype"].Value != null)
        //            guid_type = row.Cells["TeklaSKtype"].Value.ToString();

        //        //check idqc
        //        int idqc = -1;
        //        if (row.Cells["idqc"].Value != null)
        //        {
        //            string tmp = row.Cells["idqc"].Value.ToString();
        //            if (skWinLib.IsNumeric(tmp) == true)
        //                idqc = Convert.ToInt32(tmp);
        //        }

        //        if (row.Cells["remark"].Value != null)
        //            remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

        //        //sql = sql + " , (" + id_project.ToString() + ",'" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_emp + ",'" + skWinLib.systemname + "'," + newreport_stage + "," + timestamp + ")";
        //        sql = sql + " , ('" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_ncm_progress + "," + id_emp + "," + timestamp  + ")";

        //        rowct++;

        //        //based on bulk record count perform insert 
        //        //for better performance
        //        if (rowct == bulkrecord)
        //        {
        //            sql = sql.Replace(", ,", "");
        //            mycmd.CommandText = insert_statement + sql;
        //            mycmd.Connection = myconn;
        //            if (myconn.State == ConnectionState.Closed)
        //                myconn.Open();
        //            insct = insct + mycmd.ExecuteNonQuery();
        //            sql = ",";
        //            rowct = 0;
        //        }
        //    }
        //    //replace ,, with ,
        //    sql = sql.Replace(", ,", "");
        //    mycmd.CommandText = insert_statement + sql;
        //    mycmd.Connection = myconn;
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    insct = insct + mycmd.ExecuteNonQuery();
        //    return insct;
        //}
        private int bulk_insert(int id_ncm_progress, long timestamp, int bulkrecord, DataGridView myDataGridView)
        {

            //string insert_statement = "INSERT INTO epmatdb.ncm(project_id,guid,guid_type,obser,chk_rmk,chk_user_id,chk_sys,id_ncm_progress,entry_date) VALUES ";
            string insert_statement = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress, chk_user_id, chk_dt) VALUES ";
            string sql = ",";
            int rowct = 0;
            int insct = 0;
            foreach (DataGridViewRow row in myDataGridView.Rows)
            {

                //check observation is not null or empty space
                string obser = string.Empty;
                string guid = string.Empty;
                string guid_type = string.Empty;
                string remark = string.Empty;


                //get observation
                if (row.Cells["observation"].Value != null)
                    obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

                //get guid
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();

                //get guidtype
                if (row.Cells["TeklaSKtype"].Value != null)
                    guid_type = row.Cells["TeklaSKtype"].Value.ToString();

                //check idqc
                int idqc = -1;
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                        idqc = Convert.ToInt32(tmp);
                }

                if (row.Cells["remark"].Value != null)
                    remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());



                //sql = sql + " , (" + id_project.ToString() + ",'" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_emp + ",'" + skWinLib.systemname + "'," + newreport_stage + "," + timestamp + ")";
                sql = sql + " , (" + ",'" + guid + "','" + guid_type + "','" + obser + "','" + remark + "'," + id_ncm_progress + "," + id_emp + "," + timestamp + ")";

                rowct++;

                //based on bulk record count perform insert 
                //for better performance
                if (rowct == bulkrecord)
                {
                    sql = sql.Replace(", ,", "");
                    mycmd.CommandText = insert_statement + sql;
                    mycmd.Connection = myconn;
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    insct = insct + mycmd.ExecuteNonQuery();
                    sql = ",";
                    rowct = 0;
                }
            }
            //replace ,, with ,
            sql = sql.Replace(", ,", "");
            mycmd.CommandText = insert_statement + sql;
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            insct = insct + mycmd.ExecuteNonQuery();
            return insct;
        }

        private int ncm_insert(string guid, string guid_type, string observation, string remark, int id_ncm_progress, int id_emp, long timestamp)
        {
            mycmd.CommandText = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress,chk_user_id,chk_dt) VALUES(@guid, @guid_type, @obser, @chk_rmk, @id_ncm_progress ,@chk_user_id, @chk_dt)";

            mycmd.Parameters.Clear();
            mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
            mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
            mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
            mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
            mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);
            mycmd.Parameters.Add("@chk_user_id", MySqlDbType.Int32);
            mycmd.Parameters.Add("@chk_dt", MySqlDbType.Int64);

            mycmd.Parameters["@guid"].Value = guid;
            mycmd.Parameters["@guid_type"].Value = guid_type;
            mycmd.Parameters["@obser"].Value = observation;
            mycmd.Parameters["@chk_rmk"].Value = remark;
            mycmd.Parameters["@id_ncm_progress"].Value = id_ncm_progress;
            mycmd.Parameters["@chk_user_id"].Value = id_emp;
            mycmd.Parameters["@chk_dt"].Value = timestamp;

            //insert new record
            return mycmd.ExecuteNonQuery();


        }

        private int ncm_insert(bool flag, string guid, string guid_type, string observation, string remark, string system_name, int id_ncm_progress, int id_emp, long timestamp)
        {
            mycmd.CommandText = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress,chk_user_id,chk_dt) VALUES(@guid, @guid_type, @obser, @chk_rmk, @id_ncm_progress, @chk_user_id, @chk_dt)";
            if (flag == false)
            {
                mycmd.Parameters.Clear();
                mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
                mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
                mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
                mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
                mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);
                mycmd.Parameters.Add("@chk_user_id", MySqlDbType.Int32);
                mycmd.Parameters.Add("@chk_dt", MySqlDbType.Int64);

            }
            mycmd.Parameters["@guid"].Value = guid;
            mycmd.Parameters["@guid_type"].Value = guid_type;
            mycmd.Parameters["@obser"].Value = observation;
            mycmd.Parameters["@chk_rmk"].Value = remark;
            mycmd.Parameters["@id_ncm_progress"].Value = id_ncm_progress;
            mycmd.Parameters["@chk_user_id"].Value = id_emp;
            mycmd.Parameters["@chk_dt"].Value = timestamp;
            //insert new record
            return mycmd.ExecuteNonQuery();


        }

        //private int  ncm_insert(bool flag, string guid, string guid_type,string observation, string remark, string system_name, int process, long timestamp)
        //{
        //    mycmd.CommandText = "INSERT INTO epmatdb.ncm(project_id,guid,guid_type,obser,chk_rmk,chk_user_id,chk_sys,id_ncm_progress,entry_date) VALUES(@project_id, @guid, @guid_type, @obser, @chk_rmk, @chk_user_id, @chk_sys, @id_ncm_progress, @entry_date)";
        //    if (flag == false)
        //    {
        //        mycmd.Parameters.Clear();
        //        mycmd.Parameters.Add("@project_id", MySqlDbType.Int64);
        //        mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
        //        mycmd.Parameters.Add("@chk_user_id", MySqlDbType.LongText);
        //        mycmd.Parameters.Add("@chk_sys", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);
        //        mycmd.Parameters.Add("@entry_date", MySqlDbType.Int64);

        //    }

        //    //get variable value from data grid
        //    //guid = row.Cells["guid"].Value.ToString();
        //    //guid_type = row.Cells["TeklaSKtype"].Value.ToString();
        //    //remark = row.Cells["remark"].Value.ToString();





        //    mycmd.Parameters["@project_id"].Value = id_project;
        //    mycmd.Parameters["@guid"].Value = guid;
        //    mycmd.Parameters["@guid_type"].Value = guid_type;
        //    mycmd.Parameters["@obser"].Value = observation;
        //    mycmd.Parameters["@chk_rmk"].Value = remark;
        //    mycmd.Parameters["@chk_user_id"].Value = id_emp;
        //    mycmd.Parameters["@chk_sys"].Value = system_name;
        //    mycmd.Parameters["@id_ncm_progress"].Value = process;
        //    mycmd.Parameters["@entry_date"].Value = timestamp;



        //    //insert new record
        //    return mycmd.ExecuteNonQuery();


        //}

        //private int ncm_insert(bool flag, string guid, string guid_type, string observation, string remark, string system_name, long system_time, int process)
        //{
        //    mycmd.CommandText = "INSERT INTO epmatdb.ncm(project_id,guid,guid_type,obser,chk_rmk,chk_user_id,chk_sys,chk_dt,id_ncm_progress) VALUES(@project_id, @guid, @guid_type, @obser, @chk_rmk, @chk_user_id, @chk_sys, @chk_dt, @id_ncm_progress)";
        //    if (flag == false)
        //    {
        //        mycmd.Parameters.Clear();
        //        mycmd.Parameters.Add("@project_id", MySqlDbType.Int64);
        //        mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
        //        mycmd.Parameters.Add("@chk_user_id", MySqlDbType.LongText);
        //        mycmd.Parameters.Add("@chk_sys", MySqlDbType.VarChar);
        //        mycmd.Parameters.Add("@chk_dt", MySqlDbType.Int64);
        //        mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);

        //    }

        //    //get variable value from data grid
        //    //guid = row.Cells["guid"].Value.ToString();
        //    //guid_type = row.Cells["TeklaSKtype"].Value.ToString();
        //    //remark = row.Cells["remark"].Value.ToString();





        //    mycmd.Parameters["@project_id"].Value = id_project;
        //    mycmd.Parameters["@guid"].Value = guid;
        //    mycmd.Parameters["@guid_type"].Value = guid_type;
        //    mycmd.Parameters["@obser"].Value = observation;
        //    mycmd.Parameters["@chk_rmk"].Value = remark;
        //    mycmd.Parameters["@chk_user_id"].Value = id_emp;
        //    mycmd.Parameters["@chk_sys"].Value = system_name;
        //    mycmd.Parameters["@chk_dt"].Value = system_time;
        //    mycmd.Parameters["@id_ncm_progress"].Value = process;



        //    //insert new record
        //    return mycmd.ExecuteNonQuery();
        //}
        private int ncm_update(DataGridViewRow row, int id_report, long timestamp)
        {

            //function used to update existing records not for new entry

            string guid = string.Empty;
            string guid_type = string.Empty;
            int idqc = -1;
            string obser = string.Empty;
            string remark = string.Empty;

            int db_bdisagree = -1;
            int db_bdismodelupdated = -1;
            int db_bdcat = -1;
            int db_backcheck = -1;
            int db_dbseverity = -1;


            //get variable value from data grid
            if (row.Cells["guid"].Value != null)
                guid = row.Cells["guid"].Value.ToString();

            //guid type like bolt, part, weld, connection
            if (row.Cells["TeklaSKtype"].Value != null)
                guid_type = row.Cells["TeklaSKtype"].Value.ToString();

            //idqc 
            if (row.Cells["idqc"].Value != null)
            {
                string tmp = row.Cells["idqc"].Value.ToString();
                if (skWinLib.IsNumeric(tmp) == true)
                    idqc = Convert.ToInt32(tmp);
            }

            if (row.Cells["observation"].ReadOnly == false)
            {
                //check observation is not null or empty space
                if (row.Cells["observation"].Value != null)
                    obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());
            }

            if (row.Cells["remark"].ReadOnly == false)
            {
                //remark by user checker, back draft or back check or any other
                if (row.Cells["remark"].Value != null)
                    remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());
            }

            if (row.Cells["isagreed"].ReadOnly == false)
            {
                //bd model agree or disagree the observation given by checker
                if (row.Cells["isagreed"].Value != null)
                {
                    string agreed = row.Cells["isagreed"].Value.ToString();
                    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
                        db_bdisagree = 1;
                    else
                        db_bdisagree = 0;

                }
            }
            if (row.Cells["modelupdated"].ReadOnly == false)
            {
                //whether model is updated or not
                if (row.Cells["modelupdated"].Value != null)
                {
                    string modelupdated = row.Cells["modelupdated"].Value.ToString();
                    if (modelupdated.ToUpper().Trim() == "1" || modelupdated.ToUpper().Trim() == "TRUE")
                        db_bdismodelupdated = 1;
                    else
                        db_bdismodelupdated = 0;

                }
            }
            if (row.Cells["category"].ReadOnly == false)
            {
                //in case of observation category type
                if (row.Cells["category"].Value != null)
                {
                    string category = row.Cells["category"].Value.ToString();
                    db_bdcat = getListIndex(qmsCategory, category);

                }
                else
                    db_bdcat = 0;
            }
            if (row.Cells["backcheck"].ReadOnly == false)
            {
                //in case of observation whether back check done or not
                if (row.Cells["backcheck"].Value != null)
                {
                    string agreed = row.Cells["backcheck"].Value.ToString();
                    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
                        db_backcheck = 1;
                    else
                        db_backcheck = 0;
                }
            }
            if (row.Cells["severity"].ReadOnly == false)
            {
                //severity to be entered by DPM
                if (row.Cells["severity"].Value != null)
                {
                    string severity = row.Cells["severity"].Value.ToString();
                    db_dbseverity = getListIndex(qmsSeverity, severity);

                }
                else
                    db_dbseverity = 0;
            }

            string update_cond = " SET ";
            //check observation
            if (obser != "" && obser != null && obser.Trim().Length >= 1)
            {
                update_cond = update_cond + ", obser = '" + obser + "', chk_user_id = " + id_emp + ", chk_dt = " + timestamp;
            }

            //check agree or disagree
            if (db_bdisagree != -1)
            {
                update_cond = update_cond + ", bd_isagree = " + db_bdisagree + ", bd_user_id = " + id_emp + ", bd_dt = " + timestamp;
            }

            //check model updated or not
            if (db_bdismodelupdated != -1)
            {
                if (update_cond.Contains("bd_user_id") == false)
                    update_cond = update_cond + ", bd_ismodelupdated = " + db_bdismodelupdated + ", bd_user_id = " + id_emp + ", bd_dt = " + timestamp;
                else
                    update_cond = update_cond + ", bd_ismodelupdated = " + db_bdismodelupdated;
            }


            //check agree or disagree
            if (db_backcheck != -1)
            {
                update_cond = update_cond + ", backcheck = " + db_backcheck + ", backcheck_user_id = " + id_emp + ", backcheck_dt = " + timestamp;
            }

            //check agree or disagree
            if (db_bdcat != -1)
            {
                update_cond = update_cond + ", id_cat = " + db_bdcat + ", bd_user_id = " + id_emp + ", bd_dt = " + timestamp;
            }

            //check agree or disagree
            if (db_dbseverity != -1)
            {
                update_cond = update_cond + ", id_sev = " + db_dbseverity + ", sev_user_id = " + id_emp + ", sev_dt = " + timestamp;
            }

            //check remark
            if (remark != "" && remark != null && remark.Trim().Length >= 1)
            {
                if (id_process == (int)SKProcess.Modelling)
                    update_cond = update_cond + ", mod_rmk = '" + remark + "'";
                else if (id_process < (int)SKProcess.ModelChecking_Completed)
                    update_cond = update_cond + ", chk_rmk = '" + remark + "'";
                else if (id_process < (int)SKProcess.ModelBackDraft_Completed)
                    update_cond = update_cond + ", bd_rmk = '" + remark + "'";
                else if (id_process < (int)SKProcess.ModelBackCheck_Completed)
                    update_cond = update_cond + ", backcheck_rmk = '" + remark + "'";
            }
            if (update_cond != " SET ")
            {
                update_cond = update_cond.Replace(",,", ",");
                update_cond = update_cond.Replace("SET ,", "SET ");
                mysql = "UPDATE `ncm` " + update_cond + " WHERE idqc=" + idqc + " and id_report = " + id_report;
                mycmd.CommandText = mysql;

                //update record
                return mycmd.ExecuteNonQuery();
            }
            else
                return 0;
        }

        private int ncm_update(int id_process, int i_idqc, int id_ncm_progress, string observation = "", int db_bdisagree = -1, int category = -1, int db_backcheck = -1, int severity = -1, string remark = "")
        {

            //id_process, idqc, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark,  timestamp
            //already record exist need to update changed field value
            //for better performance check with earlier stored datagrid against idqc and if difference found update those column only 

            //int i_idqc = Convert.ToInt32(row.Cells["idqc"].Value.ToString());

            //int db_bdisagree = 0;
            //string db_cat = string.Empty;

            //if (row.Cells["remark"].Value != null)
            //    remark = row.Cells["remark"].Value.ToString();

            //if (row.Cells["isagreed"].Value != null)
            //{
            //    string agreed = row.Cells["isagreed"].Value.ToString();
            //    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
            //        db_bdisagree = 1;
            //}

            //int db_backcheck = 0;
            //if (row.Cells["backcheck"].Value != null)
            //{
            //    string agreed = row.Cells["backcheck"].Value.ToString();
            //    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
            //        db_backcheck = 1;
            //}




            string update_cond = " SET ";
            //check observation
            if (observation != "" && observation != null && observation.Trim().Length >= 1)
            {
                update_cond = update_cond + ", obser = '" + observation + "'";
            }
            //check agree or disagree
            if (db_bdisagree != -1)
            {
                update_cond = update_cond + ", bd_isagree = " + db_bdisagree;
            }

            //check agree or disagree
            if (db_backcheck != -1)
            {
                update_cond = update_cond + ", backcheck = " + db_backcheck;
            }

            //check agree or disagree
            if (category != -1)
            {
                update_cond = update_cond + ", id_cat = " + category;
            }

            //check agree or disagree
            if (severity != -1)
            {
                update_cond = update_cond + ", id_sev = " + severity;
            }

            //check remark
            if (remark != "" && remark != null && remark.Trim().Length >= 1)
            {
                if (id_process < (int)SKProcess.ModelChecking_Completed)
                    update_cond = update_cond + ", chk_rmk = '" + remark + "'";
                else if (id_process < (int)SKProcess.ModelBackDraft_Completed)
                    update_cond = update_cond + ", bd_rmk = '" + remark + "'";
                else if (id_process < (int)SKProcess.ModelBackCheck_Completed)
                    update_cond = update_cond + ", backcheck_rmk = '" + remark + "'";
            }
            if (update_cond != " SET ")
            {
                update_cond = update_cond.Replace(",,", ",");
                update_cond = update_cond.Replace("SET ,", "SET ");

                mysql = "UPDATE `ncm` " + update_cond + " WHERE idqc=" + i_idqc + " and id_ncm_progress = " + id_ncm_progress;
                mycmd.CommandText = mysql;
                //update record
                return mycmd.ExecuteNonQuery();
            }
            else
                return 0;
        }


        private string Update_ModifiedData(ArrayList dbColumnName, ArrayList MyUpdatedData, int id_ncm_progress, long timestamp)
        {
            string returnmsg = "";
            int err_ct = 0;
            int updt_ct = 0;


            for (int ct = 0; ct < MyUpdatedData.Count; ct++)
            {
                string mydata = MyUpdatedData[ct].ToString();
                mydata = mydata.Replace(sk_seperator, "|");
                string[] splitdata = mydata.Split(new Char[] { '|' });
                mysql = "";
                string update_cond = " SET ";
                int idqc = -1;
                for (int splitdatact = 0; splitdatact < dbColumnName.Count; splitdatact++)
                {
                    string mysqlcolumn = dbColumnName[splitdatact].ToString();
                    string[] splitmysqlcolumn = mysqlcolumn.Split(new Char[] { '|' });
                    if (splitmysqlcolumn[0].ToUpper() == "IN" || splitmysqlcolumn[0].ToUpper() == "DO")
                    {
                        if (splitmysqlcolumn[1].ToUpper() == "IDQC")
                        {
                            idqc = Convert.ToInt32(splitdata[splitdatact]);
                        }
                        else
                        {
                            string db_columname = splitmysqlcolumn[1];
                            string s_db_columvalue = splitdata[splitdatact];
                            int db_columvalue = 0;
                            //for checkbox type
                            if (s_db_columvalue.ToUpper().Trim() == "1" || s_db_columvalue.ToUpper().Trim() == "TRUE")
                                db_columvalue = 1;
                            //for combobox type
                            if (db_columname == "id_cat")
                                db_columvalue = getListIndex(qmsCategory, s_db_columvalue);
                            if (db_columname == "id_sev")
                                db_columvalue = getListIndex(qmsSeverity, s_db_columvalue);
                            update_cond = update_cond + ", " + db_columname + " = " + db_columvalue;
                        }

                    }
                    else if (splitmysqlcolumn[0].ToUpper() == "ST")
                    {
                        update_cond = update_cond + ", " + splitmysqlcolumn[1] + " = '" + splitdata[splitdatact] + "'";
                    }


                }
                if (idqc != -1)
                {
                    if (update_cond != " SET ")
                    {
                        update_cond = update_cond.Replace(",,", ",");
                        update_cond = update_cond.Replace("SET ,", "SET ");

                        // update user_id, date 
                        update_cond = update_cond + GetCurrentProcess_UpdateQuery_Userid_Date(SK_id_process, timestamp);

                        mysql = "UPDATE `ncm` " + update_cond + " WHERE idqc = " + idqc + " and status is null";
                        mycmd.CommandText = mysql;
                        //update record
                        int flag = mycmd.ExecuteNonQuery();
                        if (flag != -1 && flag != 0)
                            updt_ct++;
                        else
                            err_ct++;

                        //mysql = "UPDATE `ncm` " +
                    }

                    /*
   UPDATE table SET Col1 = CASE id 
                          WHEN 1 THEN 1 
                          WHEN 2 THEN 2 
                          WHEN 4 THEN 10 
                          ELSE Col1 
                        END, 
                 Col2 = CASE id 
                          WHEN 3 THEN 3 
                          WHEN 4 THEN 12 
                          ELSE Col2 
                        END
             WHERE id IN (1, 2, 3, 4);
                    example 1
                    ysql> update UpdateAllDemo
   −> set BookName = (CASE BookId WHEN 1000 THEN 'C in Depth'
   −> when 1001 THEN 'Java in Depth'
   −> END)
   −> Where BookId IN(1000,1001);
    example twow

                    UPDATE table
SET column2 = (CASE column1 WHEN 1 THEN 'val1'
                 WHEN 2 THEN 'val2'
                 WHEN 3 THEN 'val3'
         END)
WHERE column1 IN(1, 2 ,3);
                    */
                }


            }


            if (updt_ct >= 1)
                returnmsg = returnmsg + updt_ct.ToString() + " Updated ";
            if (err_ct >= 1)
                returnmsg = returnmsg + err_ct.ToString() + " Error ";
            return returnmsg;
        }

        private string Update_ModifiedData(ArrayList dbColumnName, ArrayList MyUpdatedData, long timestamp)
        {
            string returnmsg = "";
            int err_ct = 0;
            int updt_ct = 0;

            string whereidqc = ",";
            ArrayList whencolumnthen = new ArrayList(dbColumnName.Count);


            for (int ct = 0; ct < MyUpdatedData.Count; ct++)
            {
                string mydata = MyUpdatedData[ct].ToString();
                mydata = mydata.Replace(sk_seperator, "|");
                string[] splitdata = mydata.Split(new Char[] { '|' });
                string idqc = splitdata[0];
                whereidqc = whereidqc + "," + idqc;
                for (int splitdatact = 1; splitdatact < dbColumnName.Count; splitdatact++)
                {
                    string mysqlcolumn = dbColumnName[splitdatact].ToString();
                    string[] splitmysqlcolumn = mysqlcolumn.Split(new Char[] { '|' });
                    if (splitmysqlcolumn[0].ToUpper() == "IN" || splitmysqlcolumn[0].ToUpper() == "DO")
                    {
                        string db_columname = splitmysqlcolumn[1];
                        string s_db_columvalue = splitdata[splitdatact];
                        int db_columvalue = 0;
                        //for checkbox type
                        if (s_db_columvalue.ToUpper().Trim() == "1" || s_db_columvalue.ToUpper().Trim() == "TRUE")
                            db_columvalue = 1;
                        //for combobox type
                        if (db_columname == "id_cat")
                            db_columvalue = getListIndex(qmsCategory, s_db_columvalue);
                        if (db_columname == "id_sev")
                            db_columvalue = getListIndex(qmsSeverity, s_db_columvalue);

                        whencolumnthen[splitdatact] = whencolumnthen[splitdatact] + " WHEN " + idqc + " THEN " + db_columvalue;
                    }
                    else if (splitmysqlcolumn[0].ToUpper() == "ST")
                    {
                        whencolumnthen[splitdatact] = whencolumnthen[splitdatact] + " WHEN " + idqc + " THEN '" + splitdata[splitdatact] + "'";
                    }


                }
            }
            for (int splitdatact = 1; splitdatact < dbColumnName.Count; splitdatact++)
            {
                string mysqlcolumn = dbColumnName[splitdatact].ToString();
                string[] splitmysqlcolumn = mysqlcolumn.Split(new Char[] { '|' });
                string db_columname = splitmysqlcolumn[1];
                //update_cond = update_cond + GetCurrentProcess_UpdateQuery_Userid_Date(SK_id_process, timestamp);

                //mysql = "UPDATE `ncm` " + update_cond + " WHERE idqc = " + idqc + " and status is null";

                mysql = "UPDATE `ncm` SET " + db_columname + " = (CASE Column1 " + whencolumnthen[splitdatact] + "  END) WHERE Column1 IN(" + whereidqc + ") and status is null";
                mycmd.CommandText = mysql;
                //update record
                int flag = mycmd.ExecuteNonQuery();
                if (flag != -1 && flag != 0)
                    updt_ct++;
                else
                    err_ct++;
            }


                //string update_cond = " SET ";

//                UPDATE table
//SET column2 = (CASE column1 WHEN 1 THEN 'val1'
//                 WHEN 2 THEN 'val2'
//                 WHEN 3 THEN 'val3'
//         END)
//WHERE column1 IN(1, 2, 3)


                
                //if (idqc != -1)
                //{
                    //if (update_cond != " SET ")
                    //{
                    //    update_cond = update_cond.Replace(",,", ",");
                    //    update_cond = update_cond.Replace("SET ,", "SET ");

                    //    // update user_id, date 
                    //    update_cond = update_cond + GetCurrentProcess_UpdateQuery_Userid_Date(SK_id_process, timestamp);

                    //    mysql = "UPDATE `ncm` " + update_cond + " WHERE idqc = " + idqc + " and status is null";
                    //    mycmd.CommandText = mysql;
                    //    //update record
                    //    int flag = mycmd.ExecuteNonQuery();
                    //    if (flag != -1 && flag != 0)
                    //        updt_ct++;
                    //    else
                    //        err_ct++;

                    //    //mysql = "UPDATE `ncm` " +
                    //}

                    /*
   UPDATE table SET Col1 = CASE id 
                          WHEN 1 THEN 1 
                          WHEN 2 THEN 2 
                          WHEN 4 THEN 10 
                          ELSE Col1 
                        END, 
                 Col2 = CASE id 
                          WHEN 3 THEN 3 
                          WHEN 4 THEN 12 
                          ELSE Col2 
                        END
             WHERE id IN (1, 2, 3, 4);
                    example 1
                    ysql> update UpdateAllDemo
   −> set BookName = (CASE BookId WHEN 1000 THEN 'C in Depth'
   −> when 1001 THEN 'Java in Depth'
   −> END)
   −> Where BookId IN(1000,1001);
    example twow

                    UPDATE table
SET column2 = (CASE column1 WHEN 1 THEN 'val1'
                 WHEN 2 THEN 'val2'
                 WHEN 3 THEN 'val3'
         END)
WHERE column1 IN(1, 2 ,3);
                    */
                //}


            //}


            if (updt_ct >= 1)
                returnmsg = returnmsg + updt_ct.ToString() + " Updated ";
            if (err_ct >= 1)
                returnmsg = returnmsg + err_ct.ToString() + " Error ";
            return returnmsg;
        }

        private string GetCurrentProcess_UpdateQuery_Userid_Date(int current_id_process, long current_timestamp)
        {

            //get before remarks column id
            if (current_id_process == (int)SKProcess.Modelling)
            {
                //modelling
                return ", mod_user_id = " + id_emp + ", mod_dt = " + current_timestamp;
            }
            else if (current_id_process == (int)SKProcess.Modelling_PM_Approved)
            {
                //checking 
                return ", chk_user_id = " + id_emp + ", chk_dt = " + current_timestamp;

            }
            else if (current_id_process == (int)SKProcess.ModelChecking_PM_Approved)
            {
                //back drafting 
                return ", bd_user_id = " + id_emp + ", bd_dt = " + current_timestamp;

            }
            else if (current_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
            {
                //back checking
                return ", backcheck_user_id = " + id_emp + ", backcheck_dt = " + current_timestamp;

            }
            else if (current_id_process == (int)SKProcess.ModelBackCheck_PM_Approved)
            {
                //severity process
                return ", sev_user_id = " + id_emp + ", sev_dt = " + current_timestamp;

            }
            else
            {
                return "";
            }
        }

        private string GetCurrentProcess_SelectQuery_Userid(int current_id_process)
        {
            //based on process get the data for the current_user and null entry also (no one performed task earlier hence free current user can commence the work but before saving ensure null entry in user_id else conflict between user entry)
            if (SK_id_process == (int)SKProcess.Modelling)
            {
                return " and id_sk_process = " + current_id_process.ToString() + " and mod_user_id = " + id_emp;
            }
            else if (SK_id_process == (int)SKProcess.Modelling_PM_Approved)
            {
                return " and id_sk_process = " + current_id_process.ToString() + " and (chk_user_id = " + id_emp + " or chk_user_id is null)";

            }
            else if (SK_id_process == (int)SKProcess.ModelChecking_PM_Approved)
            {

                return " and id_sk_process = " + current_id_process.ToString() + " and obser NOT LIKE '%" + sk_nocomment + "%' and obser is not null and (bd_user_id = " + id_emp + " or bd_user_id is null )";

            }
            else if (SK_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
            {

                return " and id_sk_process = " + current_id_process.ToString() + " and obser NOT LIKE '%" + sk_nocomment + "%' and obser is not null and (backcheck_user_id = " + id_emp + " or backcheck_user_id is null)";

            }
            else if (SK_id_process == (int)SKProcess.ModelBackCheck_PM_Approved)
            {

                return " and id_sk_process = " + current_id_process.ToString() + " and obser NOT LIKE '%" + sk_nocomment + "%' and obser is not null and (sev_user_id = " + id_emp + " or sev_user_id is null)";

            }
            else if (SK_id_process == (int)SKProcess.ModelSeverity_PM_Approved)
            {

                return " and id_sk_process = " + current_id_process.ToString() + " and obser NOT LIKE '%" + sk_nocomment + "%' and obser is not null and (sev_user_id = " + id_emp + " or sev_user_id is null)";

            }
            else
            {
                return "";// " and id_sk_process = " + current_id_process.ToString();
            }
        }

        private string GetCurrentProcess_BeforeRemark(int current_id_process)
        {
            //get before remarks column id
            if (SK_id_process == (int)SKProcess.Modelling)
            {
                //return " and mod_user_id = " + id_emp;
                return "chk_rmk";
            }
            else if (SK_id_process == (int)SKProcess.Modelling_PM_Approved)
            {
                return "mod_rmk, bd_rmk";
            }
            else if (SK_id_process == (int)SKProcess.ModelChecking_PM_Approved)
            {
                return "mod_rmk, chk_rmk, backcheck_rmk";
            }
            else if (SK_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
            {
                return "mod_rmk, chk_rmk, bd_rmk, sev_rmk";
            }
            //else if (SK_id_process == (int)SKProcess.ModelBackCheck_PM_Approved)
            //{
            //    return "mod_rmk, chk_rmk, bd_rmk, backcheck_rmk, pm_rmk";
            //}
            //else if (SK_id_process == (int)SKProcess.ModelSeverity_PM_Approved)
            //{
            //    return "mod_rmk, chk_rmk, bd_rmk, backcheck_rmk, pm_rmk";
            //}
            else
            {
                return "mod_rmk, chk_rmk, bd_rmk, backcheck_rmk, sev_rmk, pm_rmk";
            }
        }
        private string Update_DataGridViewRowData(DataGridView myDataGridView)
        {
            string returnmsg = "";
            int err_ct = 0;
            int updt_ct = 0;

            string guid = string.Empty;
            int idqc = -1;
            string obser = string.Empty;
            string remark = string.Empty;
            string guid_type = string.Empty;

            for (int rowindex = 0; rowindex < myDataGridView.Rows.Count; rowindex++)
            {
                DataGridViewRow row = myDataGridView.Rows[rowindex];
                //check observation is not null or empty space
                obser = string.Empty;

                if (row.Cells["observation"].Value != null)
                    obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

                guid = string.Empty;
                //get variable value from data grid
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();


                guid_type = string.Empty;
                if (row.Cells["TeklaSKtype"].Value != null)
                    guid_type = row.Cells["TeklaSKtype"].Value.ToString();

                idqc = -1;
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                        idqc = Convert.ToInt32(tmp);
                }
                remark = string.Empty;
                if (row.Cells["remark"].Value != null)
                    remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

                int db_bdisagree = 0;
                if (row.Cells["isagreed"].Value != null)
                {
                    string agreed = row.Cells["isagreed"].Value.ToString();
                    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
                        db_bdisagree = 1;
                    else
                        db_bdisagree = 0;

                }

                int db_bdcat = -1;
                if (row.Cells["category"].Value != null)
                {
                    string category = row.Cells["category"].Value.ToString();
                    db_bdcat = getListIndex(qmsCategory, category);

                }


                int db_dbseverity = -1;
                if (row.Cells["severity"].Value != null)
                {
                    string severity = row.Cells["severity"].Value.ToString();
                    db_dbseverity = getListIndex(qmsSeverity, severity);
                }

                int db_backcheck = -1;
                if (row.Cells["backcheck"].Value != null)
                {
                    string agreed = row.Cells["backcheck"].Value.ToString();
                    if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
                        db_backcheck = 1;
                    else
                        db_backcheck = 0;
                }

                //already record exist need to update changed field value
                //for better performance check with earlier stored datagrid against idqc and if difference found update those column only                    
                string db_cat = string.Empty;
                int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
                //int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
                //for status log
                if (flag != -1 && flag != 0)
                    updt_ct++;
                else
                    err_ct++;
            }

            if (updt_ct >= 1)
                returnmsg = returnmsg + updt_ct.ToString() + " Updated ";
            if (err_ct >= 1)
                returnmsg = returnmsg + err_ct.ToString() + " Error ";
            return returnmsg;
        }
        //private string ncm_save()
        //{
        //    string returnmsg = "";

        //    int ins_ct = 0;
        //    int err_ct = 0;
        //    int updt_ct = 0;

        //    string guid = string.Empty;
        //    int idqc = -1;
        //    string obser = string.Empty;
        //    string remark = string.Empty;
        //    string guid_type = string.Empty;


        //    DateTime chk_dt = DateTime.Now;
        //    DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
        //    long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();
        //    {
        //        //check whether already report created
        //        if (id_stage == -1)
        //        {
        //            id_stage = CreateNewReport();
        //        }
        //        //insert data to my sql
        //        if (dgqc.RowCount >= 1)
        //        {
        //            try
        //            {
        //                if (myconn.State == ConnectionState.Closed)
        //                    myconn.Open();
        //                bool paraflag = false;
        //                int ob_null_ct = 0;
        //                for (int rowindex = 0; rowindex < dgqc.Rows.Count; rowindex++)
        //                {
        //                    DataGridViewRow row = this.dgqc.Rows[rowindex];
        //                    //check observation is not null or empty space
        //                    obser = string.Empty;

        //                    if (row.Cells["observation"].Value != null)
        //                        obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

        //                    guid = string.Empty;
        //                    //get variable value from data grid
        //                    if (row.Cells["guid"].Value != null)
        //                        guid = row.Cells["guid"].Value.ToString();


        //                    guid_type = string.Empty;
        //                    if (row.Cells["TeklaSKtype"].Value != null)
        //                        guid_type = row.Cells["TeklaSKtype"].Value.ToString();

        //                    idqc = -1;
        //                    if (row.Cells["idqc"].Value != null)
        //                    {
        //                        string tmp = row.Cells["idqc"].Value.ToString();
        //                        if (skWinLib.IsNumeric(tmp) == true)
        //                            idqc = Convert.ToInt32(tmp);
        //                    }
        //                    remark = string.Empty;
        //                    if (row.Cells["remark"].Value != null)
        //                        remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

        //                    int db_bdisagree = 0;
        //                    if (row.Cells["isagreed"].Value != null)
        //                    {
        //                        string agreed = row.Cells["isagreed"].Value.ToString();
        //                        if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
        //                            db_bdisagree = 1;
        //                        else
        //                            db_bdisagree = 0;

        //                    }




        //                    int db_bdcat = -1;
        //                    if (row.Cells["category"].Value != null)
        //                    {
        //                        string category = row.Cells["category"].Value.ToString();
        //                        db_bdcat = getListIndex(qmsCategory, category);
        //                        //DataGridViewComboBoxCell mycat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        //                        ////string chk = mycat.Items.IndexOf(category);
        //                        //if (skWinLib.IsNumeric(category) == true)
        //                        //    db_bdcat = 1;
        //                    }


        //                    int db_dbseverity = -1;
        //                    if (row.Cells["severity"].Value != null)
        //                    {
        //                        string severity = row.Cells["severity"].Value.ToString();
        //                        db_dbseverity = getListIndex(qmsSeverity, severity);
        //                    }

        //                    int db_backcheck = -1;
        //                    if (row.Cells["backcheck"].Value != null)
        //                    {
        //                        string agreed = row.Cells["backcheck"].Value.ToString();
        //                        if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
        //                            db_backcheck = 1;
        //                        else
        //                            db_backcheck = 0;
        //                    }
        //                    //if (idqc == 1211)
        //                    //{
        //                    //    string cc = "";
        //                    //}
        //                    if (idqc == -1)
        //                    {
        //                        //insert fresh record into my sql
        //                        int flag = ncm_insert(guid, guid_type, obser, remark, id_stage);
        //                        //int flag = ncm_insert(paraflag, guid, guid_type, obser, remark, skWinLib.systemname,  id_stage, timestamp); 
        //                        //for status log
        //                        if (flag != -1)
        //                        {
        //                            ins_ct++;
        //                            paraflag = true;
        //                        }
        //                        else
        //                            err_ct++;
        //                    }
        //                    else
        //                    {

        //                        //already record exist need to update changed field value
        //                        //for better performance check with earlier stored datagrid against idqc and if difference found update those column only                    
        //                        string db_cat = string.Empty;
        //                        int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
        //                        //int flag = ncm_update(id_process, idqc, id_stage, obser, db_bdisagree, db_bdcat, db_backcheck, db_dbseverity, remark);
        //                        //for status log
        //                        if (flag != -1 && flag != 0)
        //                            updt_ct++;
        //                        else
        //                            err_ct++;

        //                    }




        //                }
        //                myconn.Close();
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show("SQL Error: " + ex.Message);
        //            }
        //        }
        //    }

        //    //Fetching saved
        //    if (ins_ct >= 1)
        //    {
        //        returnmsg = ins_ct.ToString() + " Added ";
        //        //initally idqc won't be there for insert hence to map idqc with user interface datagrid fetch the data and map idqc to corresponding row
        //        //mysql = "select guid,idqc, obser, chk_rmk,guid_type from epmatdb.ncm where project_id=" + cmbprojid.Tag + " and chk_user_id =" + lblxsuser.Tag + " and id_ncm_progress = " + Convert.ToInt32(lblstage.Tag) + " and chk_sys ='" + skWinLib.systemname + "' and chk_dt =" + l_chk_dt + " order by idqc";
        //        //mysql = "select guid,idqc, obser, chk_rmk,guid_type from epmatdb.ncm where chg_rmk is null and project_id=" + id_project + " and chk_user_id =" + id_emp + " and id_ncm_progress = " + id_stage + " and chk_sys ='" + skWinLib.systemname + "' order by idqc";

        //        //function update idqc for the newly added records;
        //        Update_idqc_NewRecord(id_stage, timestamp);


        //    }
        //    if (updt_ct >= 1)
        //        returnmsg = returnmsg + updt_ct.ToString() + " Updated ";
        //    if (err_ct >= 1)
        //        returnmsg = returnmsg + err_ct.ToString() + " Error ";

        //    if ((ins_ct + updt_ct + err_ct) >= 1)
        //    {
        //        //refresh
        //        Update_DashBoard_TreeView();
        //    }

        //    return returnmsg;
        //}

        private void Update_idqc_NewRecord(int id_report, int id_emp, long timestamp)
        {
            mysql = "select guid, idqc, mod_rmk, guid_type from epmatdb.ncm where status is null and id_report = " + id_report + " and mod_user_id = " + id_emp + " and mod_dt = " + timestamp + " order by idqc";
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            dgqc.Rows.Clear();



            //dgqc.RecreatingHandle 
            int dg_rowid = 0;
            while (myrdr.Read())
            {
                string db_guid = myrdr[0].ToString();
                string db_idqc = myrdr[1].ToString();
                string db_mod_rmk = skWinLib.mysql_set_quotecomma(myrdr[2].ToString());
                string db_guid_type = myrdr[3].ToString();

                //add new row need to better the performance

                dg_rowid = dgqc.Rows.Add(db_guid);
                DataGridViewRow row = this.dgqc.Rows[dg_rowid];

                row.Cells["guid"].Value = db_guid;
                row.Cells["TeklaSKtype"].Value = db_guid_type;
                row.Cells["idqc"].Value = db_idqc;
                row.Cells["remark"].Value = db_mod_rmk;

            }
            myrdr.Close();
            myconn.Close();
            mycmd.Dispose();
            //dgqc.Refresh();
        }
        //private void Update_idqc_NewRecord(int id_ncm_progress)
        //{
        //    mysql = "select guid,idqc, obser, chk_rmk,guid_type from epmatdb.ncm where status is null and id_ncm_progress = " + id_ncm_progress + " order by idqc";
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    int dg_rowid = 0;
        //    while (myrdr.Read())
        //    {

        //        //check idqc is not null or empty space

        //        string db_guid = myrdr[0].ToString();
        //        string db_idqc = myrdr[1].ToString();
        //        string db_obser = skWinLib.mysql_set_quotecomma(myrdr[2].ToString());
        //        string db_chk_rmk = skWinLib.mysql_set_quotecomma(myrdr[3].ToString());
        //        string db_guid_type = myrdr[4].ToString();

        //    CheckNextRow_idqc:
        //        string guid = string.Empty;
        //        int idqc = -1;
        //        string obser = string.Empty;
        //        string remark = string.Empty;
        //        string guid_type = string.Empty;

        //        DataGridViewRow row = this.dgqc.Rows[dg_rowid];
        //        if (row.Cells["guid"].Value != null)
        //            guid = row.Cells["guid"].Value.ToString();
        //        if (row.Cells["idqc"].Value != null)
        //        {
        //            string tmp = row.Cells["idqc"].Value.ToString();
        //            if (skWinLib.IsNumeric(tmp) == true)
        //                idqc = Convert.ToInt32(tmp);
        //        }

        //        if (row.Cells["observation"].Value != null)
        //            obser = row.Cells["observation"].Value.ToString();
        //        if (row.Cells["remark"].Value != null)
        //            remark = row.Cells["remark"].Value.ToString();
        //        if (row.Cells["TeklaSKtype"].Value != null)
        //            guid_type = row.Cells["TeklaSKtype"].Value.ToString();


        //        if ((guid == db_guid) && (idqc == -1) && (obser == db_obser) && (remark == db_chk_rmk))
        //        {
        //            row.Cells["idqc"].Value = db_idqc;
        //            row.Cells["observation"].Style.BackColor = System.Drawing.Color.LightGreen;
        //            dg_rowid++;
        //            dgqc.Refresh();
        //        }
        //        else
        //        {
        //            //since all datagridrow and mysql records does not match check next row in datagrid
        //            dg_rowid++;
        //            goto CheckNextRow_idqc;
        //        }
        //    }


        //    myrdr.Close();
        //    myconn.Close();
        //    mycmd.Dispose();
        //    dgqc.Refresh();
        //}

        private void Update_idqc_NewRecord(int id_ncm_progress, long timestamp)
        {
            mysql = "select guid,idqc, obser, chk_rmk,guid_type from epmatdb.ncm where status is null and id_ncm_progress = " + id_ncm_progress + " order by idqc";
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            int dg_rowid = 0;
            while (myrdr.Read())
            {

                //check idqc is not null or empty space

                string db_guid = myrdr[0].ToString();
                string db_idqc = myrdr[1].ToString();
                string db_obser = myrdr[2].ToString();
                string db_chk_rmk = myrdr[3].ToString();
                string db_guid_type = myrdr[4].ToString();

            CheckNextRow_idqc:
                string guid = string.Empty;
                int idqc = -1;
                string obser = string.Empty;
                string remark = string.Empty;
                string guid_type = string.Empty;

                DataGridViewRow row = this.dgqc.Rows[dg_rowid];
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                        idqc = Convert.ToInt32(tmp);

                }

                if (row.Cells["observation"].Value != null)
                    obser = row.Cells["observation"].Value.ToString();
                if (row.Cells["remark"].Value != null)
                    remark = row.Cells["remark"].Value.ToString();
                if (row.Cells["TeklaSKtype"].Value != null)
                    guid_type = row.Cells["TeklaSKtype"].Value.ToString();


                if ((guid == db_guid) && (idqc == -1) && (obser == db_obser) && (remark == db_chk_rmk))
                {
                    row.Cells["idqc"].Value = db_idqc;
                    row.Cells["observation"].Style.BackColor = System.Drawing.Color.LightGreen;
                    dg_rowid++;
                    //dgqc.Refresh();
                }
                else
                {
                    //since all datagridrow and mysql records does not match check next row in datagrid
                    dg_rowid++;
                    goto CheckNextRow_idqc;
                }
            }


            myrdr.Close();
            myconn.Close();
            mycmd.Dispose();
            //dgqc.Refresh();
        }

        //private void Update_idqc_NewRecord(int id_ncm_progress, long timestamp)
        //{
        //    mysql = "select guid,idqc, obser, chk_rmk,guid_type from epmatdb.ncm where chg_rmk is null and project_id=" + id_project + " and chk_user_id =" + id_emp + " and id_ncm_progress = " + id_ncm_progress + " and chk_sys ='" + skWinLib.systemname + "' and entry_date =" + timestamp + " order by idqc";
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    int dg_rowid = 0;
        //    while (myrdr.Read())
        //    {

        //        //check idqc is not null or empty space

        //        string db_guid = myrdr[0].ToString();
        //        string db_idqc = myrdr[1].ToString();
        //        string db_obser = myrdr[2].ToString();
        //        string db_chk_rmk = myrdr[3].ToString();
        //        string db_guid_type = myrdr[4].ToString();

        //    CheckNextRow_idqc:
        //        string guid = string.Empty;
        //        int idqc = -1;
        //        string obser = string.Empty;
        //        string remark = string.Empty;
        //        string guid_type = string.Empty;

        //        DataGridViewRow row = this.dgqc.Rows[dg_rowid];
        //        if (row.Cells["guid"].Value != null)
        //            guid = row.Cells["guid"].Value.ToString();
        //        if (row.Cells["idqc"].Value != null)
        //        {
        //            string tmp = row.Cells["idqc"].Value.ToString();
        //            if (skWinLib.IsNumeric(tmp) == true)
        //                idqc = Convert.ToInt32(tmp);

        //        }

        //        if (row.Cells["observation"].Value != null)
        //            obser = row.Cells["observation"].Value.ToString();
        //        if (row.Cells["remark"].Value != null)
        //            remark = row.Cells["remark"].Value.ToString();
        //        if (row.Cells["TeklaSKtype"].Value != null)
        //            guid_type = row.Cells["TeklaSKtype"].Value.ToString();


        //        if ((guid == db_guid) && (idqc == -1) && (obser == db_obser) && (remark == db_chk_rmk))
        //        {
        //            row.Cells["idqc"].Value = db_idqc;
        //            row.Cells["observation"].Style.BackColor = System.Drawing.Color.LightGreen;
        //            dg_rowid++;
        //            dgqc.Refresh();
        //        }
        //        else
        //        {
        //            //since all datagridrow and mysql records does not match check next row in datagrid
        //            dg_rowid++;
        //            goto CheckNextRow_idqc;
        //        }
        //    }


        //    myrdr.Close();
        //    myconn.Close();
        //    mycmd.Dispose();
        //    dgqc.Refresh();
        //}

        private ArrayList Update_Dashboard_ReportNumber()
        {
            ArrayList myreturn1 = new ArrayList();
            ArrayList myreturn = new ArrayList();

            //check whether any report created for this project and get the report id
            string id_report_list = ",";
            mysql = "select distinct id_report, name, number from epmatdb.ncm_report where status is null and project_id = " + id_project + " order by id_report";
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            //clear all rows
            dgreportid.Rows.Clear();
            while (myrdr.Read())
            {
                id_report_list = id_report_list + "," + myrdr[0].ToString();
                myreturn1.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString());
                //add to datagrid report so that all user know the list of report available
                string rptno = myrdr[2].ToString();
                if (rptno.Length == 1)
                {
                    rptno = "00" + rptno;
                }
                else if (rptno.Length == 2)
                {
                    rptno = "0" + rptno;
                }
                dgreportid.Rows.Add(myrdr[0].ToString(), myrdr[1].ToString(), myrdr[2].ToString());
                dgreportid.Rows[dgreportid.Rows.Count - 1].HeaderCell.Value = rptno;
            }
            //auto size 
            //dgreportid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgreportid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders);
            //dgreportid.AutoResizeColumns();
            dgreportid.AutoResizeRows();
            dgreportid.Refresh();

            myrdr.Close();
            myconn.Close();

            myreturn.Add(id_report_list);
            myreturn.Add(myreturn1);
            return myreturn;
        }

        private ArrayList GetReportStage_Based_Project(int id_project)
        {
            //select id_ncm_progress from epmatdb.ncm where project_id=389 and chk_user_id =297
            ArrayList myreturn1 = new ArrayList();
            ArrayList myreturn2 = new ArrayList();
            ArrayList myreturn = new ArrayList();

            string id_report_list = ",";

            ArrayList unq_id_sk_process_id_report = new ArrayList();
            ArrayList tmp = Update_Dashboard_ReportNumber();
            if (tmp.Count >= 1)
            {
                id_report_list = tmp[0].ToString();
                myreturn1 = tmp[1] as ArrayList;

                //check report list for records
                if (id_report_list != ",")
                {

                    id_report_list = id_report_list.Replace(",,", "");
                    //select distinct id_report, id_sk_process from epmatdb.ncm where (id_sk_process = 0 and mod_user_id = 297) or (id_sk_process = 20 and (chk_user_id = 297 or chk_user_id is null)) or (id_sk_process = 40 and (bd_user_id = 297 or bd_user_id is null)) or (id_sk_process = 60 and (backcheck_user_id = 297 or backcheck_user_id is null)) or (id_sk_process = 80 and (sev_user_id = 297 or sev_user_id is null))   and status is null and id_report in (91,92,93,94,95,96,97,98,99,100,101,102,103,104,119,120,122,125,126,127,128,129) order by id_report, id_sk_process desc
                    
                    string wherecond = "((id_sk_process = " + (int)SKProcess.Modelling  + " and mod_user_id =  " + id_emp + " ) or (id_sk_process = " + (int)SKProcess.Modelling_PM_Approved + " and (chk_user_id =  " + id_emp + "  or chk_user_id is null)) or (id_sk_process = " + (int)SKProcess.ModelChecking_PM_Approved + " or (bd_user_id =  " + id_emp + "  or bd_user_id is null)) or (id_sk_process =" + (int)SKProcess.ModelBackDraft_PM_Approved + " and (backcheck_user_id =  " + id_emp + "  or backcheck_user_id is null)) or (id_sk_process = " + (int)SKProcess.ModelBackCheck_PM_Approved + " and (sev_user_id =  " + id_emp + "  or sev_user_id is null)))";

                    //string wherecond = " mod_user_id = " + id_emp + " or (chk_user_id = " + id_emp + " or chk_user_id is null) or (bd_user_id = " + id_emp + " or bd_user_id is null) or (backcheck_user_id = " + id_emp + " or backcheck_user_id is null)  or (sev_user_id = " + id_emp + " or sev_user_id is null) ";
                    //get process like modelling, checking, back draft, back check
                    
                    mysql = "select distinct id_report, id_sk_process from epmatdb.ncm where " + wherecond + " and status is null and id_report in (" + id_report_list + ") order by id_report, id_sk_process desc";
                    //mysql = "select distinct id_report, id_sk_process from epmatdb.ncm where " + wherecond + " and status is null and id_report in (" + id_report_list + ") order by id_report, id_sk_process desc";
                    //mysql = "select distinct id_sk_process,id_report from epmatdb.ncm where status is null and id_report in (" + id_report_list + ") order by id_report, id_sk_process";
                    
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    mycmd.CommandText = mysql;
                    mycmd.Connection = myconn;
                    myrdr = mycmd.ExecuteReader();
                    while (myrdr.Read())
                    {
                        string s_unq_id_sk_process_id_report = myrdr[0].ToString() + sk_seperator + myrdr[1].ToString();
                        if (!unq_id_sk_process_id_report.Contains(s_unq_id_sk_process_id_report))
                        {
                            unq_id_sk_process_id_report.Add(s_unq_id_sk_process_id_report);
                            myreturn2.Add(s_unq_id_sk_process_id_report);
                            //myreturn2.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString() + sk_seperator + myrdr[2].ToString() + sk_seperator + myrdr[3].ToString() + sk_seperator + myrdr[4].ToString() + sk_seperator + myrdr[5].ToString() + sk_seperator + myrdr[6].ToString());
                        }

                    }
                    myrdr.Close();
                    myconn.Close();
                }
            }

            myreturn.Add(myreturn1);
            myreturn.Add(myreturn2);
            return myreturn;
        }
        //private ArrayList GetReportStage_Based_Project(int id_project)
        //{
        //    //select id_ncm_progress from epmatdb.ncm where project_id=389 and chk_user_id =297
        //    ArrayList myreturn1 = new ArrayList();
        //    ArrayList myreturn2 = new ArrayList();
        //    ArrayList unqreport = new ArrayList();


        //    mysql = "select distinct id_stage, id_ncm_progress,report,user_id from epmatdb.ncm_progress where status is null and project_id=" + id_project + "  order by project_id, report, id_stage desc";
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    while (myrdr.Read())
        //    {
        //        if (!unqreport.Contains(myrdr[2].ToString()))
        //        {
        //            //only maximum stage allowed 
        //            unqreport.Add(myrdr[2].ToString());
        //            myreturn1.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString() + sk_seperator + myrdr[2].ToString() + sk_seperator + myrdr[3].ToString());
        //        }

        //    }
        //    myrdr.Close();
        //    myconn.Close();

        //    //get additional information like no comments, pending and other details
        //    //mysql = "select obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.id_ncm_progress=" + id_ncm_progress;// + " and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0";

        //    //for (int i = 0; i < myreturn1.Count; i++)
        //    //{
        //    //    string[] chkdata = myreturn1[i].ToString().Replace(sk_seperator, "|").Split(new Char[] { '|' });
        //    //    int id_ncm_progress = Convert.ToInt32(chkdata[1]);
        //    //    int id_progress = Convert.ToInt32(chkdata[0]);
        //    //    mysql = "";

        //    //    if (id_progress < (int)SKProcess.ModelChecking_PM_Approved)
        //    //        mysql = "select obser,obser from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " order by obser";
        //    //    else if (id_progress < (int)SKProcess.ModelBackDraft_PM_Approved)
        //    //        mysql = "select bd_isagree,id_cat from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by bd_isagree, id_cat";
        //    //    else if (id_progress < (int)SKProcess.ModelBackCheck_PM_Approved)
        //    //        mysql = "select backcheck,backcheck_rmk from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by backcheck";

        //    //    if (mysql != "")
        //    //    {
        //    //        //select id_ncm_progress,obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.project_id=360 and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0
        //    //        if (myconn.State == ConnectionState.Closed)
        //    //            myconn.Open();
        //    //        mycmd.CommandText = mysql;
        //    //        mycmd.Connection = myconn;
        //    //        myrdr = mycmd.ExecuteReader();
        //    //        string val1 = "";
        //    //        string val2 = "";

        //    //        int val1nullct = 0;
        //    //        int val2nullct = 0;
        //    //        int nocomment_ct = 0;

        //    //        int total = 0;

        //    //        while (myrdr.Read())
        //    //        {
        //    //            total++;
        //    //            val1 = myrdr[0].ToString();
        //    //            val2 = myrdr[1].ToString();

        //    //            //check if null value for obser,bd_isagree,id_cat,backcheck,id_sev
        //    //            if (val1.ToUpper().Trim().Length == 0)
        //    //                val1nullct++;
        //    //            else
        //    //            {
        //    //                if (val1.ToUpper().Trim() == sk_nocomment.ToUpper().Trim())
        //    //                    nocomment_ct++;
        //    //                else if (val1 == "0")
        //    //                {
        //    //                    val1nullct++;
        //    //                }
        //    //            }


        //    //            if (val2.ToUpper().Trim().Length == 0)
        //    //                val2nullct++;
        //    //        }
        //    //        myrdr.Close();

        //    //        if (id_progress < (int)SKProcess.ModelChecking_PM_Approved)
        //    //            myreturn2.Add(myreturn1[i].ToString() + sk_seperator + total.ToString() + sk_seperator + val1nullct.ToString() + sk_seperator + nocomment_ct.ToString());
        //    //        else if (id_progress < (int)SKProcess.ModelBackCheck_PM_Approved)
        //    //            myreturn2.Add(myreturn1[i] + sk_seperator + total.ToString() + sk_seperator + val1nullct + sk_seperator + val2nullct);
        //    //    }
        //    //    else
        //    //        myreturn2.Add(myreturn1[i] + sk_seperator + sk_seperator + sk_seperator);

        //    //}

        //    return myreturn2;
        //}

        //private ArrayList GetReportStage_Based_Project(int id_project)
        //{
        //    //select id_ncm_progress from epmatdb.ncm where project_id=389 and chk_user_id =297
        //    ArrayList myreturn1 = new ArrayList();
        //    ArrayList myreturn2 = new ArrayList();
        //    ArrayList unqreport = new ArrayList();


        //    mysql = "select distinct id_stage, id_ncm_progress,report,user_id from epmatdb.ncm_progress where status is null and project_id=" + id_project + "  order by project_id, report, id_stage desc";
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    while (myrdr.Read())
        //    {
        //        if (!unqreport.Contains(myrdr[2].ToString()))
        //        {
        //            //only maximum stage allowed 
        //            unqreport.Add(myrdr[2].ToString());
        //            myreturn1.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString() + sk_seperator + myrdr[2].ToString() + sk_seperator + myrdr[3].ToString());
        //        }

        //    }
        //    myrdr.Close();
        //    myconn.Close();

        //    //get additional information like no comments, pending and other details
        //    //mysql = "select obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.id_ncm_progress=" + id_ncm_progress;// + " and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0";

        //    for (int i = 0; i < myreturn1.Count; i++)
        //    {
        //        string[] chkdata = myreturn1[i].ToString().Replace(sk_seperator, "|").Split(new Char[] { '|' });
        //        int id_ncm_progress = Convert.ToInt32(chkdata[1]);
        //        int id_progress = Convert.ToInt32(chkdata[0]);
        //        mysql = "";

        //        if (id_progress < (int)SKProcess.ModelChecking_PM_Approved)
        //            mysql = "select obser,obser from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " order by obser";
        //        else if (id_progress < (int)SKProcess.ModelBackDraft_PM_Approved)
        //            mysql = "select bd_isagree,id_cat from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by bd_isagree, id_cat";
        //        else if (id_progress < (int)SKProcess.ModelBackCheck_PM_Approved)
        //            mysql = "select backcheck,backcheck_rmk from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by backcheck";

        //        if (mysql != "")
        //        {
        //            //select id_ncm_progress,obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.project_id=360 and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0
        //            if (myconn.State == ConnectionState.Closed)
        //                myconn.Open();
        //            mycmd.CommandText = mysql;
        //            mycmd.Connection = myconn;
        //            myrdr = mycmd.ExecuteReader();
        //            string val1 = "";
        //            string val2 = "";

        //            int val1nullct = 0;
        //            int val2nullct = 0;
        //            int nocomment_ct = 0;

        //            int total = 0;

        //            while (myrdr.Read())
        //            {
        //                total++;
        //                val1 = myrdr[0].ToString();
        //                val2 = myrdr[1].ToString();

        //                //check if null value for obser,bd_isagree,id_cat,backcheck,id_sev
        //                if (val1.ToUpper().Trim().Length == 0)
        //                    val1nullct++;
        //                else
        //                {
        //                    if (val1.ToUpper().Trim() == sk_nocomment.ToUpper().Trim())
        //                        nocomment_ct++;
        //                    else if (val1 == "0")
        //                    {
        //                        val1nullct++;
        //                    }
        //                }


        //                if (val2.ToUpper().Trim().Length == 0)
        //                    val2nullct++;
        //            }
        //            myrdr.Close();

        //            if (id_progress < (int)SKProcess.ModelChecking_PM_Approved)
        //                myreturn2.Add(myreturn1[i].ToString() + sk_seperator + total.ToString() + sk_seperator + val1nullct.ToString() + sk_seperator + nocomment_ct.ToString());
        //            else if (id_progress < (int)SKProcess.ModelBackCheck_PM_Approved)
        //                myreturn2.Add(myreturn1[i] + sk_seperator + total.ToString() + sk_seperator + val1nullct + sk_seperator + val2nullct);
        //        }
        //        else
        //            myreturn2.Add(myreturn1[i] + sk_seperator + sk_seperator + sk_seperator);

        //    }

        //    return myreturn2;
        //}
        //private ArrayList GetReportStage_Based_Project(int id_project)
        //{
        //    //select id_ncm_progress from epmatdb.ncm where project_id=389 and chk_user_id =297
        //    ArrayList myreturn1 = new ArrayList();
        //    ArrayList myreturn2 = new ArrayList();
        //    ArrayList unqreport = new ArrayList();


        //    mysql = "select distinct id_ncm_progress from epmatdb.ncm where status is null and epmatdb.ncm.project_id=" + id_project;
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();

        //    string dis_id_ncm_progress = "(";
        //    int uct = 0;
        //    while (myrdr.Read())
        //    {
        //        dis_id_ncm_progress = dis_id_ncm_progress + "," + myrdr[0].ToString();
        //        //tvw_data.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString());
        //        uct++;
        //    }
        //    myrdr.Close();
        //    string where_cond = "";
        //    if (dis_id_ncm_progress != "(")
        //    {
        //        where_cond = dis_id_ncm_progress.Replace("(,", "(");
        //        where_cond = " where id_ncm_progress in " + where_cond + ") order by  report, id_stage desc";

        //        //in future or next version change the condition so that both data records are pulled at one shot 
        //        mysql = "select id_stage, id_ncm_progress,report,user_id from epmatdb.ncm_progress " + where_cond;
        //        if (myconn.State == ConnectionState.Closed)
        //            myconn.Open();
        //        mycmd.CommandText = mysql;
        //        mycmd.Connection = myconn;
        //        myrdr = mycmd.ExecuteReader();
        //        while (myrdr.Read())
        //        {
        //            if (!unqreport.Contains(myrdr[2].ToString()))
        //            {
        //                //only maximum stage allowed 
        //                unqreport.Add(myrdr[2].ToString());
        //                myreturn1.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString() + sk_seperator + myrdr[2].ToString() + sk_seperator + myrdr[3].ToString());
        //            }

        //        }
        //        myrdr.Close();
        //        myconn.Close();
        //        //get additional information like no comments, pending and other details
        //        //mysql = "select obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.id_ncm_progress=" + id_ncm_progress;// + " and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0";

        //        for (int i = 0; i < myreturn1.Count; i++)
        //        {
        //            string[] chkdata = myreturn1[i].ToString().Replace(sk_seperator, "|").Split(new Char[] { '|' });
        //            int id_ncm_progress = Convert.ToInt32(chkdata[1]);
        //            int id_progress = Convert.ToInt32(chkdata[0]);
        //            mysql = "";

        //            if (id_progress < (int)SKProcess.ModelChecking_Completed)
        //                mysql = "select obser,obser from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " order by obser";
        //            else if (id_progress < (int)SKProcess.ModelBackDraft_Completed)
        //                mysql = "select bd_isagree,id_cat from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by bd_isagree, id_cat";

        //            else if (id_progress < (int)SKProcess.ModelBackCheck_Completed)
        //                mysql = "select backcheck,backcheck_rmk from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_ncm_progress + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by backcheck";

        //            if (mysql != "")
        //            {
        //                //select id_ncm_progress,obser,bd_isagree,id_cat,backcheck,id_sev from epmatdb.ncm where epmatdb.ncm.project_id=360 and character_length(obser) = 0  or character_length(bd_isagree) = 0  or character_length(id_cat) = 0  or character_length(backcheck) = 0 or character_length(id_sev) = 0
        //                if (myconn.State == ConnectionState.Closed)
        //                    myconn.Open();
        //                mycmd.CommandText = mysql;
        //                mycmd.Connection = myconn;
        //                myrdr = mycmd.ExecuteReader();
        //                string val1 = "";
        //                string val2 = "";

        //                int val1nullct = 0;
        //                int val2nullct = 0;
        //                int nocomment_ct = 0;

        //                int total = 0;

        //                while (myrdr.Read())
        //                {
        //                    total++;
        //                    val1 = myrdr[0].ToString();
        //                    val2 = myrdr[1].ToString();

        //                    //check if null value for obser,bd_isagree,id_cat,backcheck,id_sev
        //                    if (val1.ToUpper().Trim().Length == 0)
        //                        val1nullct++;
        //                    else
        //                    {
        //                        if (val1.ToUpper().Trim() == sk_nocomment.ToUpper().Trim())
        //                            nocomment_ct++;
        //                        else if (val1 == "0")
        //                        {
        //                            val1nullct++;
        //                        }
        //                    }


        //                    if (val2.ToUpper().Trim().Length == 0)
        //                        val2nullct++;
        //                }
        //                myrdr.Close();

        //                if (id_progress < (int)SKProcess.ModelChecking_Completed)
        //                    myreturn2.Add(myreturn1[i].ToString() + sk_seperator + total.ToString() + sk_seperator + val1nullct.ToString() + sk_seperator + nocomment_ct.ToString());
        //                else if (id_progress < (int)SKProcess.ModelBackDraft_Completed)
        //                    myreturn2.Add(myreturn1[i] + sk_seperator + total.ToString() + sk_seperator + val1nullct + sk_seperator + val2nullct);
        //            }
        //            else
        //                myreturn2.Add(myreturn1[i] + sk_seperator + sk_seperator + sk_seperator);

        //        }



        //    }



        //    return myreturn2;
        //}

        private ArrayList GetReportStage_Based_Project_User(int id_project, int id_user)
        {
            //select id_ncm_progress from epmatdb.ncm where project_id=389 and chk_user_id =297
            ArrayList myreturn = new ArrayList();
            ArrayList unqreport = new ArrayList();




            mysql = "select distinct id_ncm_progress from epmatdb.ncm where status is null and epmatdb.ncm.project_id=" + id_project;// + " and epmatdb.ncm.chk_user_id =" + id_user;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();

            string dis_id_ncm_progress = "(";

            int uct = 0;
            while (myrdr.Read())
            {
                dis_id_ncm_progress = dis_id_ncm_progress + "," + myrdr[0].ToString();
                uct++;
            }
            myrdr.Close();
            string where_cond = "";
            if (dis_id_ncm_progress != "(")
            {
                where_cond = dis_id_ncm_progress.Replace("(,", "(");
                where_cond = " where id_ncm_progress in " + where_cond + ") order by  report, id_stage desc";

                //in future or next version change the condition so that both data records are pulled at one shot 
                mysql = "select id_stage, id_ncm_progress,report,user_id from epmatdb.ncm_progress " + where_cond;
                if (myconn.State == ConnectionState.Closed)
                    myconn.Open();
                mycmd.CommandText = mysql;
                mycmd.Connection = myconn;
                myrdr = mycmd.ExecuteReader();
                while (myrdr.Read())
                {
                    if (!unqreport.Contains(myrdr[2].ToString()))
                    {
                        //only maximum stage allowed 
                        unqreport.Add(myrdr[2].ToString());
                        myreturn.Add(myrdr[0].ToString() + sk_seperator + myrdr[1].ToString() + sk_seperator + myrdr[2].ToString() + sk_seperator + myrdr[3].ToString());
                    }

                }
                myrdr.Close();
                myconn.Close();
            }



            return myreturn;
        }

        private int GetID_Based_Project_User_Report_Date_System_Stage(int id_project, int id_user, string report, long record_date, string system_name, int id_stage)
        {
            lblsbar2.Text = "GetID_Based_Project_User_Report_Date_System_Stage,id_project:" + id_project.ToString() + ",id_user:" + id_user.ToString() + ",id_user:" + id_user.ToString() + ",report:" + report + ",record_date:" + record_date.ToString() + ",id_stage:" + id_stage.ToString();
            int myreturn = -1;
            mysql = "select id_ncm_progress from epmatdb.ncm_progress  where project_id=" + id_project + " and user_id =" + id_user + " and report ='" + report + "' and user_dt =" + record_date + " and sys_id ='" + system_name + "' and id_stage = " + id_stage + " and status is null";

            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {
                myreturn = Convert.ToInt32(myrdr[0].ToString());
            }
            myrdr.Close();
            myconn.Close();
            return myreturn;
        }



        private int Create_NewProcess(int new_ProcessNumber, long new_processtimestamp)
        {
            lblsbar2.Text = "Create_NewProcess,new_ProcessNumber:" + new_ProcessNumber.ToString() + ",new_processtimestamp:" + new_processtimestamp.ToString();

            int myreturn = -1;

            //Check and create Connection
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();

            //insert fresh record into my sql
            mycmd.CommandText = "INSERT INTO epmatdb.ncm_progress(project_id,user_id,report,user_dt,sys_id,id_stage) VALUES(@project_id, @user_id, @report, @user_dt, @sys_id, @id_stage)";
            mycmd.Parameters.Clear();
            mycmd.Parameters.Add("@project_id", MySqlDbType.Int32);
            mycmd.Parameters.Add("@user_id", MySqlDbType.Int32);
            mycmd.Parameters.Add("@report", MySqlDbType.LongText);
            mycmd.Parameters.Add("@user_dt", MySqlDbType.Int64);
            mycmd.Parameters.Add("@sys_id", MySqlDbType.VarChar);
            mycmd.Parameters.Add("@id_stage", MySqlDbType.Int32);
            mycmd.Parameters["@project_id"].Value = id_project;
            mycmd.Parameters["@user_id"].Value = id_emp;
            mycmd.Parameters["@report"].Value = lblreport.Text;
            mycmd.Parameters["@user_dt"].Value = new_processtimestamp;
            mycmd.Parameters["@sys_id"].Value = skWinLib.systemname;
            mycmd.Parameters["@id_stage"].Value = new_ProcessNumber;



            //insert new process 
            return myreturn = mycmd.ExecuteNonQuery();
        }

        private int Getid_NewProcess(int new_ProcessNumber, long new_processtimestamp)
        {
            lblsbar2.Text = "Getid_NewProcess,new_ProcessNumber:" + new_ProcessNumber.ToString() + ",new_processtimestamp:" + new_processtimestamp.ToString();
            int myreturn = -1;


            //getid for the newly created process
            int status = Create_NewProcess(new_ProcessNumber, new_processtimestamp);

            //check status = 1
            if (status == 1)
            {
                myreturn = GetID_Based_Project_User_Report_Date_System_Stage(id_project, id_emp, lblreport.Text, new_processtimestamp, skWinLib.systemname, new_ProcessNumber);
            }
            else
            {
                skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Aborted");
                MessageBox.Show("error @ Getid_NewProcess \n\n\n Contact Support Team along with a screen shot", skApplicationName + " Ver. " + skApplicationVersion + "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                System.Windows.Forms.Application.ExitThread();
            }
            return myreturn;
        }
        private int CreateNewReport(int ui_stage = (int)SKProcess.Modelling_PM_Approved)
        {
            lblsbar2.Text = "CreateNewReport,ui_stage:" + ui_stage.ToString();

            int myreturn = -1;

            //create a new variable( timestamp) for lateruse
            long l_insert_dt = new DateTimeOffset(DateTime.Now).ToUnixTimeMilliseconds();

            //getid for the newly created process
            myreturn = Getid_NewProcess(ui_stage, l_insert_dt);


            if (myreturn != -1)
            {
                //change new node to next process (in progress) and update node.tag with my return value
                //check whether same item exists in new of treeview node
                //Update_DashBoard_TreeView(tvwnew, tvwchecking, myreturn, (int)SKProcess.Modelling_PM_Approved);

            }
            else
            {
                skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Aborted");
                MessageBox.Show("error @ CreateNewReport \n\n\n Contact Support Team along with a screen shot", skApplicationName + " Ver. " + skApplicationVersion + "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                System.Windows.Forms.Application.ExitThread();
            }

            return myreturn;
        }


        private void btnLoad_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "btnLoad_Click";

            ////vijay mysql = "select guid, observation from ncm.qms where idproject = '" + cmbprojid.Text + "' and chkempid = '" + cmbchkid.Text + "'"; // and sequence = '" + cmbseqid.Text + "'";
            ////trichy
            //mysql = "select guid, observation from epmat.qms where idproject = '" + cmbprojid.Text + "' and chkempid = '" + cmbchkid.Text + "'"; // and sequence = '" + cmbseqid.Text + "'";


            //if (myconn.State == ConnectionState.Closed)
            //    myconn.Open();
            //mycmd.CommandText = mysql;
            //mycmd.Connection = myconn;
            //myrdr = mycmd.ExecuteReader();
            //dgQMS.Rows.Clear();
            //while (myrdr.Read())
            //{
            //    dgQMS.Rows.Add(myrdr[0], myrdr[1]);
            //    CategorySeveritySetting(dgQMS.Rows.Count - 1);
            //}
            //myrdr.Close();


            //myconn.Close();
        }







        private void UpdateDataGridSelectedRow_String(DataGridView myDataGridView, string columnname, string cellvalue, bool replaceflag = false)
        {

            lblsbar2.Text = "UpdateDataGridSelectedRow_String,columnname:" + columnname + ",cellvalue:" + cellvalue.ToString() + ",replaceflag:" + replaceflag.ToString();
            skWinLib.DataGridView_Setting_Before(myDataGridView);

            foreach (DataGridViewRow row in myDataGridView.SelectedRows)
            {

                if (replaceflag == true)
                    row.Cells[columnname].Value = cellvalue;
                else
                {
                    if (row.Cells[columnname].Value != null)
                    {
                        if (cellvalue != "")
                            row.Cells[columnname].Value = row.Cells[columnname].Value + "\n" + cellvalue;
                        else
                            row.Cells[columnname].Value = cellvalue;
                    }

                    else
                        row.Cells[columnname].Value = cellvalue;
                }

            }

            skWinLib.DataGridView_Setting_After(myDataGridView);
        }

        private void UpdateDataGridSelectedRow_Bool(DataGridView myDataGridView, string columnname, int cellvalue)
        {
            lblsbar2.Text = "UpdateDataGridSelectedRow_Bool,columnname:" + columnname + ",cellvalue:" + cellvalue.ToString();
            skWinLib.DataGridView_Setting_Before(myDataGridView);
            foreach (DataGridViewRow row in myDataGridView.SelectedRows)
            {
                if (row.Cells[columnname].Value != null)
                    row.Cells[columnname].Value = cellvalue;

            }
            skWinLib.DataGridView_Setting_After(myDataGridView);
        }

        private void UpdateDataGridSelectedRow_ComboIndex(DataGridView myDataGridView, string columnname, int cellvalue)
        {
            lblsbar2.Text = "UpdateDataGridSelectedRow_ComboIndex,columnname:" + columnname + ",cellvalue:" + cellvalue.ToString();
            skWinLib.DataGridView_Setting_Before(myDataGridView);
            foreach (DataGridViewRow row in myDataGridView.SelectedRows)
            {
                if (columnname == "Category")
                {
                    DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
                    cellcat.DataSource = qmsCategory;
                    if (cellvalue != -1)
                        cellcat.Value = qmsCategory[cellvalue].ToString();
                }
                else if (columnname == "Severity")
                {
                    DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Severity"]);
                    cellcat.DataSource = qmsSeverity;
                    if (cellvalue != -1)
                        cellcat.Value = qmsSeverity[cellvalue].ToString();
                }

            }
            skWinLib.DataGridView_Setting_After(myDataGridView);
        }

        private void dgQMS_KeyUp(object sender, KeyEventArgs e)
        {
            string mykey = e.KeyCode.ToString().ToUpper();
            lblsbar2.Text = "dgQMS_KeyUp,key:" + mykey + ",Alt:" + e.Alt + ",process:" + SK_id_process;
            //check if the datagridview row is selected and key N pressed
            int ct = dgqc.SelectedRows.Count;
            if (ct >= 1)
            {
                //for all process remarks
                if ((mykey == "V" && e.Alt == true) || (mykey == "R" && e.Alt == true))
                {
                    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                }
                else if (mykey == "S" && e.Alt == true)
                {
                    UpdateDataGridSelectedRow_String(dgqc, "remark", "", chkreplace.Checked);
                }
                if (SK_id_process < (int)SKProcess.ModelChecking_Completed) //SK_id_process >= (int)SKProcess.Modelling_PM_Approved && 
                {
                    if (mykey == "Z" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_String(dgqc, "Observation", sk_nocomment, chkreplace.Checked);
                    }
                    else if (mykey == "X" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_String(dgqc, "Observation", "", chkreplace.Checked);
                    }
                    else if ((mykey == "C" && e.Alt == true) || (mykey == "O" && e.Alt == true))
                    {
                        UpdateDataGridSelectedRow_String(dgqc, "Observation", cmbqmsdblist.Text, chkreplace.Checked);
                    }
                    //else if ((mykey == "V" && e.Alt == true) || (mykey == "R" && e.Alt == true))
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    //else if (mykey == "S" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", "", chkreplace.Checked);
                    //}

                }
                else if (SK_id_process < (int)SKProcess.ModelBackDraft_Completed) //SK_id_process >= (int)SKProcess.Modelling_PM_Approved && 
                {
                    if (mykey == "Z" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "IsAgreed", 1);
                    }
                    else if (mykey == "X" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "IsAgreed", 0);
                    }
                    if (mykey == "Q" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "ModelUpdated", 1);
                    }
                    else if (mykey == "W" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "ModelUpdated", 0);
                    }
                    //else if (mykey == "C" && e.Alt == true)
                    //{
                    //    //UpdateDataGridSelectedRow_String(dgqc, "Observation", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    //else if ((mykey == "V" && e.Alt == true) || (mykey == "R" && e.Alt == true))
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    else if (mykey == "D1" && e.Alt == true)
                    {

                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 0);
                    }
                    else if (mykey == "D2" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 1);
                    }
                    else if (mykey == "D3" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 2);
                    }
                    else if (mykey == "D4" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 3);
                    }
                    else if (mykey == "D5" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 4);
                    }
                    else if (mykey == "D6" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 5);
                    }
                    else if (mykey == "D7" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 6);
                    }
                    else if (mykey == "D8" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 7);
                    }
                    else if (mykey == "D9" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Category", 8);
                    }

                }

                else if (SK_id_process < (int)SKProcess.ModelBackCheck_Completed) //SK_id_process >= (int)SKProcess.Modelling_PM_Approved && 
                {
                    if (mykey == "Z" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "backcheck", 1);
                    }
                    else if (mykey == "X" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_Bool(dgqc, "backcheck", 0);
                    }
                    //else if (mykey == "V" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    //else if (mykey == "R" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                }
                else if (SK_id_process < (int)SKProcess.ModelSeverity_Completed) //SK_id_process >= (int)SKProcess.Modelling_PM_Approved && 
                {
                    //if (mykey == "X" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", "", chkreplace.Checked);
                    //}
                    //else if (mykey == "V" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    //else if (mykey == "R" && e.Alt == true)
                    //{
                    //    UpdateDataGridSelectedRow_String(dgqc, "remark", cmbqmsdblist.Text, chkreplace.Checked);
                    //}
                    if (mykey == "D1" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Severity", 0);
                    }
                    else if (mykey == "D2" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Severity", 1);
                    }
                    else if (mykey == "D3" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Severity", 2);
                    }
                    else if (mykey == "D4" && e.Alt == true)
                    {
                        UpdateDataGridSelectedRow_ComboIndex(dgqc, "Severity", 3);
                    }
                }
            }


        }
        //data grid combobox free entry add
        //private void dgQMS_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        //{

        //    DataGridViewRow row = dgQMS.CurrentRow;
        //    DataGridViewCell cell = dgQMS.CurrentCell;
        //    if (e.Control.GetType() == typeof(DataGridViewComboBoxEditingControl))
        //    {

        //        if (cell == row.Cells["Observation"])
        //        {
        //            DataGridViewComboBoxEditingControl cbo = e.Control as DataGridViewComboBoxEditingControl;
        //            cbo.DropDownStyle = ComboBoxStyle.DropDown;
        //            cbo.Validating += new CancelEventHandler(cbo_Validating);
        //        }
        //    }

        //}



        void cbo_Validating(object sender, CancelEventArgs e)
        {

            DataGridViewComboBoxEditingControl cbo = sender as DataGridViewComboBoxEditingControl;
            DataGridView grid = cbo.EditingControlDataGridView;
            object value = cbo.Text;
            // Add value to list if not there
            if (cbo.Items.IndexOf(value) == -1)
            {

                DataGridViewComboBoxCell cboCol = (DataGridViewComboBoxCell)grid.CurrentCell;
                // Must add to both the current combobox as well as the template, to avoid duplicate entries...
                cbo.Items.Add(value);
                cboCol.Items.Add(value);
                dgObservation.Items.Add(value);
                grid.CurrentCell.Value = value;
            }

        }

        private void dgQMS_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //if (e.ColumnIndex == comboBoxColumn.DisplayIndex)
            //{
            //    if (!this.comboBoxColumn.Items.Contains(e.FormattedValue))
            //    {
            //        this.comboBoxColumn.Items.Add(e.FormattedValue);
            //    }
            //}


            //DataGridViewRow row = dgQMS.CurrentRow;
            //DataGridViewCell cell = dgQMS.CurrentCell;


            //if (cell == row.Cells["Observation"])
            //{
            //    DataGridViewComboBoxEditingControl cbo = sender as DataGridViewComboBoxEditingControl;
            //    DataGridView grid = cbo.EditingControlDataGridView;
            //    object value = cbo.Text;
            //    // Add value to list if not there
            //    if (cbo.Items.IndexOf(value) == -1)
            //    {

            //        DataGridViewComboBoxCell cboCol = (DataGridViewComboBoxCell)grid.CurrentCell;
            //        // Must add to both the current combobox as well as the template, to avoid duplicate entries...
            //        cbo.Items.Add(value);
            //        cboCol.Items.Add(value);
            //        grid.CurrentCell.Value = value;


            //    }
            //}
        }



        private void dgQMS_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (this.dgQMS.CurrentCellAddress.X ==   comboBoxColumn.DisplayIndex)
            //{

            //    ComboBox cb = e.Control as ComboBox;
            //    if (cb != null)
            //    {
            //        cb.DropDownStyle = ComboBoxStyle.DropDown;
            //    }
            //}


            //DataGridViewRow row = dgQMS.CurrentRow;
            //DataGridViewCell cell = dgQMS.CurrentCell;
            //if (e.Control.GetType() == typeof(DataGridViewComboBoxEditingControl))
            //{

            //    if (cell == row.Cells["Observation"])
            //    {
            //        DataGridViewComboBoxEditingControl cbo = e.Control as DataGridViewComboBoxEditingControl;
            //        cbo.DropDownStyle = ComboBoxStyle.DropDown;
            //        cbo.Validating += new CancelEventHandler(cbo_Validating);
            //    }
            //}
        }

        private void dgQMS_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            skTSLib.selectDataGridselectedrows(dgqc, 0, "GUID", chkzoom.Checked);
        }

        private void chktop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chktop.Checked;
        }

        private void btntsfetch_Click(object sender, EventArgs e)
        {

            //DateTime Today = DateTime.Now;
            //DateTime StartDate = Convert.ToDateTime(txttsfromdate.Text);
            //DateTime EndDate = Convert.ToDateTime(txttstodate.Text);
            //FetchTSData(StartDate, EndDate, "");
            //FetchTSData();
        }

        private static readonly DateTime UnixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        public static long GetCurrentUnixTimestampMillis()
        {
            return (long)(DateTime.UtcNow - UnixEpoch).TotalMilliseconds;
        }

        public static DateTime DateTimeFromUnixTimestampMillis(long millis)
        {
            return UnixEpoch.AddMilliseconds(millis);
        }

        public static long GetCurrentUnixTimestampSeconds()
        {
            return (long)(DateTime.UtcNow - UnixEpoch).TotalSeconds;
        }

        public static DateTime DateTimeFromUnixTimestampSeconds(long seconds)
        {
            return UnixEpoch.AddSeconds(seconds);
        }

        private int getEmployeeIDbySKSID(ArrayList myEmpData, string SKSID)
        {
            foreach (string mydata in myEmpData)
            {
                string[] emp_data = mydata.Split(new Char[] { '|' });
                if (SKSID.ToUpper().Trim() == emp_data[2].ToUpper().Trim())
                    return Convert.ToInt32(emp_data[0]);
            }
            return -1;

        }

        private string getEmployeeNameFromEmployeeData(ArrayList myEmpData, int Index)
        {
            int ct = myEmpData.Count;
            if (ct > Index)
            {
                string[] emp_data = myEmpData[Index].ToString().Split(new Char[] { '|' });
                return emp_data[2].ToString();
            }
            else
            {
                return "";
            }

        }

        private string getmyValue(ArrayList mylist, string SearchId)
        {
            for (int i = 0; i < mylist.Count; i++)
            {
                string[] splt_data = mylist[i].ToString().Split(new Char[] { '|' });

                if (splt_data[0].ToString() == SearchId)
                {
                    return splt_data[1].ToString();
                }
            }
            return "";
        }
        //private void FetchTSData(DateTime StartDate, DateTime EndDate, string centre, bool HolidayFlag = false)

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            

            //string current_tabpagename = TabControl1.
            if (TabControl1.SelectedIndex == 0)
            {
                //export is part dash board hence enable
                btnexport.Enabled = true;
                btnrefresh.Enabled = true;
                btnTSAll.Enabled = true;
                btnTSClear.Enabled = true;
                //btnpreview.Enabled = true;

                //btnclear.Enabled = Convert.ToBoolean(btnclear.Enabled.ToString());
                //btnfetch.Enabled = Convert.ToBoolean(btnfetch.Enabled.ToString());
                //btnSave.Enabled = Convert.ToBoolean(btnSave.Enabled.ToString());
                //btnAdd.Enabled = Convert.ToBoolean(btnAdd.Enabled.ToString());
                //btnDelete.Enabled = Convert.ToBoolean(btnDelete.Enabled.ToString());
                //btnUnique.Enabled = Convert.ToBoolean(btnUnique.Enabled.ToString());
                //btntspreview1.Enabled = Convert.ToBoolean(btntspreview1.Enabled.ToString());
                //btntspreview2.Enabled = Convert.ToBoolean(btntspreview2.Enabled.ToString());
                //btnfilter.Enabled = Convert.ToBoolean(btnfilter.Enabled.ToString());
                //btnexport.Enabled = Convert.ToBoolean(btnexport.Enabled.ToString());
                //btncopy.Enabled = Convert.ToBoolean(btncopy.Enabled.ToString());
                //btnsplitsave.Enabled = Convert.ToBoolean(btnsplitsave.Enabled.ToString());
                //chkzoom.Enabled = Convert.ToBoolean(chkzoom.Enabled.ToString());
                //chkreplace.Enabled = Convert.ToBoolean(chkreplace.Enabled.ToString());

                //EnableDisableControl(false);
            }
            else
            {
                SetUserInterface_QC(SK_id_process);
                //EnableDisableControl(false);
                //btnclear.Enabled = Convert.ToBoolean(btnclear.Tag.ToString());
                //btnfetch.Enabled = Convert.ToBoolean(btnfetch.Tag.ToString());
                //btnSave.Enabled = Convert.ToBoolean(btnSave.Tag.ToString());
                //btnAdd.Enabled = Convert.ToBoolean(btnAdd.Tag.ToString());
                //btnDelete.Enabled = Convert.ToBoolean(btnDelete.Tag.ToString());
                //btnUnique.Enabled = Convert.ToBoolean(btnUnique.Tag.ToString());
                //btntspreview1.Enabled = Convert.ToBoolean(btntspreview1.Tag.ToString());
                //btntspreview2.Enabled = Convert.ToBoolean(btntspreview2.Tag.ToString());
                //btnfilter.Enabled = Convert.ToBoolean(btnfilter.Tag.ToString());
                //btnexport.Enabled = Convert.ToBoolean(btnexport.Tag.ToString());
                //btncopy.Enabled = Convert.ToBoolean(btncopy.Tag.ToString());
                //btnsplitsave.Enabled = Convert.ToBoolean(btnsplitsave.Tag.ToString());
                //chkzoom.Enabled = Convert.ToBoolean(chkzoom.Tag.ToString());
                //chkreplace.Enabled = Convert.ToBoolean(chkreplace.Tag.ToString());
            }

            
            //else if (TabControl1.SelectedIndex == 3)
            //{
            //    EnableDisableControl(false);
            //}
            //else if (TabControl1.SelectedIndex == 2)
            //{
            //    txttstodate.Text = DateTime.Now.ToString("dd/MMM/yy");
            //    DateTime mydate = Convert.ToDateTime(txttstodate.Text);
            //    txttsfromdate.Text = mydate.AddDays(-7).ToString("dd/MMM/yy");
            //}
        }


        //int ob_null_ct = 0;
        //                if (row.Cells["observation"].Value != null)
        //                {
        //                    string obser = string.Empty;
        //obser = row.Cells["observation"].Value.ToString();
        //                }
        //                else
        //                {
        //                    ob_null_ct++;
        //                }
        //private void Update_DB_MovetoNextProcess(string employee_role, int current_process)
        //{
        //    int updtct = 0;
        //    int next_process = -1;
        //    string mymessage = string.Empty;
        //    //            if (employee_role.ToUpper().Trim() == "CHECKER" && current_process == 10)
        //    //0 Modelling complete,5 DPM , 10 Checking, 20 PM, 30 Back Drafting, 40 PM, 50 Back Checking, 60 PM
        //    if (current_process <= 10)
        //    {
        //        //checking process (10) completed and sent for dpm approval (20) and dpm assign for back drafting (30) 
        //        next_process = 30;
        //    }
        //    else if ( current_process <= 30)
        //    {
        //        //back drafting (30) completed and sent for dpm approval (40) and dpm assign for back checking (50)
        //        next_process = 50;
        //    }
        //    else if (current_process <= 50)
        //    {
        //        //back checking (50) completed and sent for dpm approval (60) and dpm assign for back checking (70)
        //        next_process = 60;
        //    }
        //    if ( next_process != -1)
        //    {
        //        //check row count is greater than or equal to one
        //        if (dgqc.Rows.Count >= 1)
        //        {

        //            //if process is back checking then ensure all back value is set as true
        //            if (next_process == 50)
        //            {
        //                if (IsBackCheckingCompleted() == false)
        //                {
        //                    //inform user at back checking stage all should be back checked to move for next stage

        //                    MessageBox.Show("Few items not back checked hence aborting now!!!\n\nComplete all back checking to proceed futher.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);

        //                    return;
        //                }
        //            }
        //            //else
        //            //{

        //            //}

        //            //myconn = new MySqlConnection(connString);


        //            //getid for the newly created process

        //            int next_id_stage = Getid_NewProcess(next_process, new DateTimeOffset(DateTime.Now).ToUnixTimeMilliseconds());
        //            if (id_stage != -1)
        //            {                        
        //                foreach (DataGridViewRow row in dgqc.Rows)
        //                { 
        //                    int i_idqc = -1;
        //                    string tmp = row.Cells["idqc"].Value.ToString();
        //                    if (skWinLib.IsNumeric(tmp) == true)
        //                        i_idqc = Convert.ToInt32(tmp);

        //                    mysql = "UPDATE `ncm` SET id_ncm_progress = " + next_id_stage + " WHERE idqc=" + i_idqc;

        //                    mycmd.CommandText = mysql;
        //                    mycmd.Connection = myconn;


        //                    if (myconn.State == ConnectionState.Closed)
        //                        myconn.Open();

        //                    //update record
        //                    if (mycmd.ExecuteNonQuery() == 1)
        //                        updtct++;


        //                }
        //                //clear all row's now since process moved
        //                dgqc.Rows.Clear();

        //                //0 Modelling complete,5 DPM , 10 Checking, 20 PM, 30 Back Drafting, 40 PM, 50 Back Checking, 60 PM
        //                if (current_process <= 10)
        //                {
        //                    Update_DashBoard_TreeView(tvwchecking, tvwbackdraft, next_id_stage, next_process);
        //                    mymessage = " guid(s) checking completed.";
        //                }
        //                else if (current_process <= 30)
        //                {
        //                    Update_DashBoard_TreeView(tvwbackdraft, tvwbackcheck, next_id_stage, next_process);
        //                    mymessage = " guid(s) back drafting completed";
        //                }
        //                else if (current_process <= 50)
        //                {
        //                    Update_DashBoard_TreeView(tvwbackcheck, tvwcompleted, next_id_stage, next_process);
        //                    mymessage = " guid(s) back checking completed";
        //                }

        //            }
        //        }
        //        else
        //            mymessage = "No guid(s) to change process.";
        //    }
        //    else
        //    {
        //        mymessage = "Employee_role: " + employee_role + "\nCurrent_process: " + current_process.ToString() + "\n\n\n  Could not get next process!!!";
        //    }
        //    if (updtct >= 1)
        //        mymessage = updtct.ToString() + mymessage;// " guid(s) updated to next process.";
        //    else
        //        mymessage = "There is nothing to save at this time.";


        //    MessageBox.Show(mymessage, msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        //}

        //mysql = "select distinct id_ncm_progress,id_stage from epmatdb.ncm_progress where project_id =" + project + " id_stage ="  + current_process + " and report  = '" + report + "'";
        //myconn = new MySqlConnection(connString);
        //if (myconn.State == ConnectionState.Closed)
        //    myconn.Open();
        //mycmd.CommandText = mysql;
        //mycmd.Connection = myconn;
        //myrdr = mycmd.ExecuteReader();
        //int id_ncm_progress = id_stage;
        //while (myrdr.Read())
        //{
        //    id_ncm_progress = Convert.ToInt32(myrdr[0].ToString());
        //}
        //myrdr.Close();


        private void MoveReport_NextProcess(int id_sk_report, int current_process, string report, string selected_dblinkid = "")
        {
            string mymessage = string.Empty;
            MessageBoxIcon myicon = MessageBoxIcon.Information;
            //get the next process number
            int next_process = -1;
            string processlog = string.Empty;
            if (current_process < 100)
            {
                ////based on current process check the user id and report id to move for next process where no null entry or pending work
                string wherecond = " WHERE id_sk_process = " + current_process + " and id_report = " + id_sk_report;
                if (current_process == (int)SKProcess.Modelling)
                {
                    //from modelling to checking
                    next_process = (int)SKProcess.Modelling_PM_Approved;
                    wherecond = wherecond + " and mod_user_id = " + id_emp;
                    processlog = "Modelling completed and ready for checking.";
                }
                else if (current_process == (int)SKProcess.Modelling_PM_Approved)
                {
                    //from checking to back draft
                    next_process = (int)SKProcess.ModelChecking_PM_Approved;
                    wherecond = wherecond + " and chk_user_id = " + id_emp + " and LENGTH(obser) != 0 ";
                    processlog = "Checking completed and ready for Backdrafting.";
                }
                else if (current_process == (int)SKProcess.ModelChecking_PM_Approved)
                {
                    //from back draft to back checking
                    next_process = (int)SKProcess.ModelBackDraft_PM_Approved;
                    wherecond = wherecond + " and bd_user_id = " + id_emp + " and (bd_isagree is not null and bd_isagree != 0)  and (bd_ismodelupdated is not null and bd_ismodelupdated != 0) and (id_cat is not null and id_cat != 0)";
                    processlog = "Backdraft completed and ready for Backchecking.";
                }
                else if (current_process == (int)SKProcess.ModelBackDraft_PM_Approved)
                {
                    //from back checking to severity process
                    next_process = (int)SKProcess.ModelBackCheck_PM_Approved;
                    wherecond = wherecond + " and backcheck_user_id = " + id_emp + " and (backcheck is not null and backcheck != 0)";
                    processlog = "Backcheck completed and ready for Severity Entry.";
                }
                else if (current_process == (int)SKProcess.ModelBackCheck_PM_Approved)
                {
                    //from severity to back checking completed
                    next_process = (int)SKProcess.ModelSeverity_PM_Approved;
                    wherecond = wherecond + " and sev_user_id = " + id_emp + " and (id_sev is not null and id_sev != 0)";
                    processlog = "Severity entry completed and ready for QA Analysis.";
                }
                if (next_process != -1)
                {
                    mysql = "UPDATE `ncm` SET id_sk_process = " + next_process + wherecond + selected_dblinkid;

                    mycmd.CommandText = mysql;
                    mycmd.Connection = myconn;


                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();

                    //update record
                    int updtct = mycmd.ExecuteNonQuery();

                    if (updtct >= 1)
                    {
                        int totalct = 100;
                        int diffct = totalct - updtct;
                        if (diffct == 0)
                            mymessage = "All " + updtct.ToString() + " records moved to next process.";
                        else
                            mymessage = diffct.ToString() + " guid(s) have pending task. " + updtct.ToString() + " records moved to next process.";

                        mymessage = updtct.ToString() + " records moved to next process.";

                        //send email for DPM
                        skWinLib.SendEmail(skApplicationVersion , skTSLib.ModelName,  report, lblxsuser.Text , updtct.ToString(), processlog);

                        //refresh
                        Update_DashBoard_TreeView();
                        myicon = MessageBoxIcon.Information;
                    }
                    else
                    {
                        mymessage = "Nothing to move.\n\nCheck your inputs and contact support team.";
                        myicon = MessageBoxIcon.Error;
                    }


                }
                else
                {
                    myicon = MessageBoxIcon.Error;
                    mymessage = "Current process can't moved to next process. Check your inputs and contact support team.";
                }



            }
            else
            {
                myicon = MessageBoxIcon.Stop;
                //(int project, string report, int current_process)
                mymessage = "Report: " + tpgqc.Text.ToString() + " is not the current selection. Check your inputs and contact support team.";
            }
        skipMoveReport_NextProcess:;
            MessageBox.Show(mymessage, msgCaption, MessageBoxButtons.OK, myicon);
        }

        //private void MoveReport_NextProcess(int id_sk_report, int current_process, string report)
        //{
        //    string mymessage = string.Empty;
        //    MessageBoxIcon myicon = MessageBoxIcon.Information;
        //    int updtct = 0;
        //    TreeNode myselectednode = tvwncm.SelectedNode;
        //    if (myselectednode.Text == tpgqc.Text)
        //    {
        //        //get the next process number
        //        int next_process = -1;
        //        if (current_process == 0)
        //        {
        //            string wherecond = " WHERE id_report = " + id_sk_report + " and mod_user_id = " + id_emp + " and status is null";
        //            ////based on process set user_id
        //            //if (SK_id_process == (int)SKProcess.Modelling)
        //            //{
        //            //    wherecond = wherecond + " and mod_user_id = " + id_emp;
        //            //}
        //            //else if (SK_id_process == (int)SKProcess.Modelling_PM_Approved)
        //            //{
        //            //    wherecond = wherecond + " and chk_user_id = " + id_emp + " and obser is not null";
        //            //}
        //            //else if (SK_id_process == (int)SKProcess.ModelChecking_PM_Approved)
        //            //{
        //            //    wherecond = wherecond + " and bd_user_id = " + id_emp;
        //            //}
        //            //else if (SK_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
        //            //{
        //            //    wherecond = wherecond + " and backcheck_user_id = " + id_emp;
        //            //}
        //            next_process = (int)SKProcess.Modelling_PM_Approved;
        //            mysql = "UPDATE `ncm` SET id_sk_process = " + next_process + wherecond ;

        //            mycmd.CommandText = mysql;
        //            mycmd.Connection = myconn;


        //            if (myconn.State == ConnectionState.Closed)
        //                myconn.Open();

        //            //update record
        //            updtct = mycmd.ExecuteNonQuery();


        //            //send email for DPM
        //            skWinLib.SendEmail(report);

        //            //refresh
        //            Update_DashBoard_TreeView();

        //            myicon = MessageBoxIcon.Information;
        //            mymessage = updtct.ToString() + " guid(s) moved to next process.";
        //        }
        //        else
        //        {
        //            //check the report name whether all activity completed if not inform except parital complete all other are moved to next process            int updtct = 0;



        //            if (current_process == (int)SKProcess.Modelling_PM_Approved)
        //            {
        //                foreach (TreeNode child in myselectednode.Nodes)
        //                {
        //                    //get child node here
        //                    if (child.Text.ToString().ToUpper().IndexOf("PENDING") >= 0)
        //                    {
        //                        myicon = MessageBoxIcon.Error;
        //                        //(int project, string report, int current_process)
        //                        mymessage = "Observation " + child.Text + ".\nComplete all observation or split observation of completed to move for next process.";
        //                        goto skipMoveReport_NextProcess;
        //                    }
        //                }
        //                //first check whether all observation is no comment to sent next_process as completed
        //                mysql = "select distinct obser from epmatdb.ncm where status is null and id_report = " + id_sk_report;

        //                ArrayList myobser = new ArrayList();
        //                myconn = new MySqlConnection(connString);
        //                if (myconn.State == ConnectionState.Closed)
        //                    myconn.Open();
        //                mycmd.CommandText = mysql;
        //                mycmd.Connection = myconn;
        //                myrdr = mycmd.ExecuteReader();
        //                while (myrdr.Read())
        //                {
        //                    myobser.Add(myrdr[0].ToString());

        //                }
        //                myrdr.Close();
        //                if (myobser == null)
        //                {
        //                    //no records inform user that it is empty report hence can't move to next process
        //                    myicon = MessageBoxIcon.Hand;
        //                    //(int project, string report, int current_process)
        //                    mymessage = "No record's found.\nDelete this reports or add checking observation to move for next process.";
        //                    goto skipMoveReport_NextProcess;
        //                }
        //                else
        //                {
        //                    //change next_process to complete if all are no comments
        //                    next_process = (int)SKProcess.ModelChecking_PM_Approved;
        //                    if (myobser.Count == 1 && myobser[0].ToString().ToUpper().Trim() == sk_nocomment.ToUpper().Trim())
        //                        next_process = (int)SKProcess.Completed;
        //                }

        //                //checking process (10) completed and sent for dpm approval (20) and dpm assign for back drafting (30) 
        //                //mysql = "select idqc,obser from epmatdb.ncm where chg_rmk is null and id_ncm_progress = " + id_ncm_progress;// + " and character_length(obser) = 0";



        //            }
        //            else if (current_process < (int)SKProcess.ModelBackDraft_Completed)
        //            {
        //                //back drafting (30) completed and sent for dpm approval (40) and dpm assign for back checking (50)
        //                next_process = (int)SKProcess.ModelBackDraft_PM_Approved;
        //            }
        //            else if (current_process < (int)SKProcess.ModelBackCheck_Completed)
        //            {
        //                //back checking (50) completed and sent for dpm approval (60) and dpm assign for back checking (70)
        //                next_process = (int)SKProcess.ModelBackCheck_PM_Approved;
        //            }
        //            mysql = "select idqc,obser from epmatdb.ncm where status is null and id_report = " + id_sk_report;
        //            if (next_process != -1)
        //            {
        //                string update_cond = ",";
        //                int errct = 0;
        //                myconn = new MySqlConnection(connString);
        //                if (myconn.State == ConnectionState.Closed)
        //                    myconn.Open();
        //                mycmd.CommandText = mysql;
        //                mycmd.Connection = myconn;
        //                myrdr = mycmd.ExecuteReader();
        //                while (myrdr.Read())
        //                {
        //                    if (myrdr[1].ToString().Trim().Length == 0)
        //                    {
        //                        errct++;
        //                    }
        //                    else
        //                    {
        //                        update_cond = update_cond + "," + myrdr[0].ToString();
        //                    }

        //                }
        //                myrdr.Close();
        //                if (update_cond != ",")
        //                {
        //                    update_cond = update_cond.Replace(",,", "");

        //                    //getid for the newly created process
        //                    //int next_id_process = Getid_NewProcess(next_process, new DateTimeOffset(DateTime.Now).ToUnixTimeMilliseconds());

        //                    //id_ncm_progress in " + update_cond + ")";
        //                    string wherecond_userid = GetCurrentProcess_SelectQuery_Userid(SK_id_process);
        //                    //if (wherecond_userid != "")
        //                    //    wherecond_userid = " and " + wherecond_userid;
        //                    mysql = "UPDATE `ncm` SET id_sk_process = " + next_process + " WHERE idqc in (" + update_cond + ")" + wherecond_userid;

        //                    mycmd.CommandText = mysql;
        //                    mycmd.Connection = myconn;


        //                    if (myconn.State == ConnectionState.Closed)
        //                        myconn.Open();

        //                    //update record
        //                    updtct = mycmd.ExecuteNonQuery();

        //                    //send email for DPM
        //                    skWinLib.SendEmail(report);

        //                    //refresh
        //                    Update_DashBoard_TreeView();

        //                    //if (current_process <= (int)SKProcess.ModelChecking_Completed)
        //                    //{
        //                    //    Update_DashBoard_TreeView(tvwchecking, tvwbackdraft, next_id_process, next_process);
        //                    //    mymessage = updtct.ToString() +  " guid(s) checking completed.";
        //                    //}
        //                    //else if (current_process <= (int)SKProcess.ModelBackDraft_Completed)
        //                    //{
        //                    //    Update_DashBoard_TreeView(tvwbackdraft, tvwbackcheck, next_id_process, next_process);
        //                    //    mymessage = updtct.ToString() + " guid(s) back drafting completed";
        //                    //}
        //                    //else if (current_process <= (int)SKProcess.ModelBackCheck_Completed)
        //                    //{
        //                    //    Update_DashBoard_TreeView(tvwbackcheck, tvwcompleted, next_id_process, next_process);
        //                    //    mymessage = updtct.ToString() + " guid(s) back checking completed";

        //                    //}
        //                    myicon = MessageBoxIcon.Information;
        //                    mymessage = updtct.ToString() + " guid(s) moved to next process.";
        //                }


        //            }
        //        }

        //    }
        //    else
        //    {
        //        myicon = MessageBoxIcon.Stop;
        //        //(int project, string report, int current_process)
        //        mymessage = "Report: " + tpgqc.Text.ToString() + " is not the current selection. Check your inputs and contact support team.";
        //    }
        //skipMoveReport_NextProcess:;
        //    MessageBox.Show(mymessage, msgCaption, MessageBoxButtons.OK, myicon);
        //}

        private string CopySelectedRows_To_BackDraftProcess()
        {
            int errct = 0;
            int updtct = 0;
            int newct = 0;
            string msg = "";
            string backcheckobservationlist = ",";

            DateTime chk_dt = DateTime.Now;
            DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
            long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

            //check back check additional information are there if not inform user about this

            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();

            foreach (DataGridViewRow row in dgqc.SelectedRows)
            {
                string idqc = string.Empty;
                string guid = string.Empty;
                string guidtype = string.Empty;
                string mod_user_id = string.Empty;
                string mod_rmk = string.Empty;
                string mod_dt = string.Empty;

                string obser = string.Empty;
                string chk_user_id = string.Empty;
                string chk_rmk = string.Empty;

                string backcheck_obser = string.Empty;
                string backcheck_rmk = string.Empty;
                string backcheck = string.Empty;

                //idqc
                if (row.Cells["idqc"].Value != null)
                    idqc =  row.Cells["idqc"].Value.ToString();

                //GUID
                if (row.Cells["guid"].Value != null)
                    guid = row.Cells["guid"].Value.ToString();

                //guidtype
                if (row.Cells["TeklaSKtype"].Value != null)
                    guidtype = row.Cells["TeklaSKtype"].Value.ToString();

                //mod_user_id
                if (row.Cells["mod_user_id"].Value != null)
                    mod_user_id = row.Cells["mod_user_id"].Value.ToString();

                //Observation
                if (row.Cells["observation"].Value != null)
                    obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());


                //Back Check Observation
                if (row.Cells["backcheck_obser"].Value != null)
                    backcheck_obser = skWinLib.mysql_replace_quotecomma(row.Cells["backcheck_obser"].Value.ToString());


                //Back Check Observation
                if (row.Cells["backcheck"].Value != null)
                    backcheck = row.Cells["backcheck"].Value.ToString();

                int isbackcheck = 0;
                if (backcheck == "1" || backcheck.ToUpper() == "TRUE")
                    isbackcheck = 1;

                //Back Check Remark
                if (row.Cells["Remark"].Value != null)
                    backcheck_rmk = skWinLib.mysql_replace_quotecomma(row.Cells["Remark"].Value.ToString());

                //get idqc
                if (backcheck_obser.Length >= 1)
                {
                    mysql = "UPDATE `ncm` SET backcheck = " + isbackcheck + ", backcheck_obser = '" + backcheck_obser + "', backcheck_user_id = " + id_emp + ", backcheck_rmk= '" + backcheck_rmk + "', status = 'REBACKDRAFT', status_dt = " + timestamp + ", status_user_id = " + id_emp + " WHERE idqc = " + idqc;
                    mycmd.CommandText = mysql;
                    //update record
                    updtct = updtct + mycmd.ExecuteNonQuery();
                    //id_report,id_sk_process,guid,guid_type,mod_user_id,mod_rmk,mod_dt, obser, chk_user_id, chk_rmk, chk_dt
                    backcheckobservationlist = backcheckobservationlist + ", (" + SK_id_report + ", " + (int)SKProcess.ModelChecking_PM_Approved + ", '" + guid + "', '" + guidtype + "', " + mod_user_id + ", 'BACKCHECK-BackDraft', " + timestamp + ", '" + backcheck_obser + "', " + id_emp + ", '" + backcheck_rmk + "', " + timestamp + ")";

                }
                else
                    errct++;
  
            }
            //insert new record for back draft
            if (backcheckobservationlist != ",")
            {
                backcheckobservationlist = backcheckobservationlist.Replace(",,", "");
                newct = bulk_copy_backcheck_backdraft(backcheckobservationlist);
            }


            msg = "Process Completed.\n" + newct.ToString() + " records added.\n" + updtct.ToString() + " records updated.";
            if (errct >=1)
                msg = msg + "\n" + errct.ToString() + " records error.";
            return msg;
        }

        private int bulk_copy_backcheck_backdraft(string backcheckobservationlist)
        {
            string insert_statement = "INSERT INTO epmatdb.ncm(id_report,id_sk_process,guid,guid_type,mod_user_id,mod_rmk,mod_dt, obser, chk_user_id, chk_rmk, chk_dt) VALUES ";
            mycmd.CommandText = insert_statement + backcheckobservationlist;
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            int insct = mycmd.ExecuteNonQuery();
            return insct;
        }

        private int Move_To_BeforeProcess(int id_sk_report, int current_process)
        {
            //get the previous process number
            int previous_process = -1;
            string wherecond = "";
            if (current_process != 0)
            {
                wherecond = " and id_report = " + id_sk_report + " and status is null";
                //based on process set user_id
                if (SK_id_process == (int)SKProcess.Modelling_PM_Approved)
                {
                    //checking process
                    previous_process = (int)SKProcess.Modelling;
                    wherecond = wherecond + " and chk_user_id = " + id_emp;
                }
                else if (SK_id_process == (int)SKProcess.ModelChecking_PM_Approved)
                {
                    //back draft process
                    previous_process = (int)SKProcess.Modelling_PM_Approved;
                    wherecond = wherecond + " and bd_user_id = " + id_emp;
                }
                else if (SK_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
                {
                    //back checking process
                    previous_process = (int)SKProcess.ModelChecking_PM_Approved;
                    wherecond = wherecond + " and backcheck_user_id = " + id_emp;
                }
                else if (SK_id_process == (int)SKProcess.ModelBackCheck_PM_Approved)
                {
                    //severity process
                    previous_process = (int)SKProcess.ModelBackDraft_PM_Approved;
                    wherecond = wherecond + " and sev_user_id = " + id_emp;
                }

            }
            int updtct = 0;
            string update_cond = ",";
            foreach (DataGridViewRow row in dgqc.SelectedRows)
            {
                //get idqc
                if (row.Cells["idqc"].Value != null)
                {
                    string tmp = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(tmp) == true)
                    {

                        update_cond = update_cond + "," + tmp;
                    }

                }
            }
            if (update_cond != ",")
            {
                update_cond = update_cond.Replace(",,", "");
                mysql = "UPDATE `ncm` SET id_sk_process = " + previous_process + " WHERE idqc in (" + update_cond + ")" + wherecond;
                mycmd.CommandText = mysql;
                mycmd.Connection = myconn;


                if (myconn.State == ConnectionState.Closed)
                    myconn.Open();

                //update record
                updtct = updtct + mycmd.ExecuteNonQuery();
            }
            return updtct;
        }
        private void Update_Stage_ReportName(string report, string stagename, int stage)
        {
            //string mycaption = this.Text;
            //user have selected new/inprogress/completed/held or cancelled node
            //int newindx = mycaption.IndexOf(" [[[");
            //if (newindx >= 0)
            //{
            //    //mycaption = mycaption.Substring(0, newindx);
            //}
            //if ((stagename + report).Length == 0)
            if (report.Length == 0)
            {
                //this.Text = mycaption;
                this.tpgqc.Parent = null; //hide
            }
            else
            {
                //this.Text = mycaption + " [[[ " + stage + " : " + report + " ]]]";
                this.tpgqc.Parent = this.TabControl1; //show
                if (report != cmbqmsdblist.Text)
                {
                    //clear all existing row values
                    dgqc.Rows.Clear();
                }
                if (cmbqmsdblist.Text.Trim().Length == 0 || cmbqmsdblist.Text.ToUpper().Trim() != report.ToUpper().Trim())
                {
                    cmbqmsdblist.Text = report;
                    //tpgdash.Tag = report;
                }


            }
            lblstage.Text = stagename;
            lblreport.Text = report;
            tpgqc.Text = report;
            //lblstage.Tag = stage.ToString();
            id_stage = stage;


        }
        private void tvwncm_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
        private void tvwncm_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            ToolTip mytooltip = new ToolTip();
            //isowner = false;
            //lblxsuser.Tag = "";

            //get the selected node from tree view
            TreeNode mynode = tvwncm.SelectedNode;
            if (mynode != null)
            {
                //get node color
                System.Drawing.Color nodecolor = mynode.ForeColor;

                //get selected node full path 
                string tvwpath = mynode.FullPath;
                string[] tvwdata = tvwpath.Split(new Char[] { '+' });

                string stage_displayname = string.Empty;
                if (tvwdata.Length >= 3)
                    stage_displayname = tvwdata[2];

                string report = "";

                //if length is 4 then user has selected one of the process (i.e checking, back draft, back check,.....)
                if (tvwdata.Length == 4)
                {
                    //get the report name
                    report = tvwdata[3].ToString();

                    //set owner information
                    if (stage_displayname == "New")
                        isowner = true;
                    else if (stage_displayname == "QA")
                        isowner = false;
                    else
                    {
                        if (nodecolor != System.Drawing.Color.Red)
                        {
                            isowner = true;
                            //lblxsuser.Tag = "Owner";
                        }
                        else
                        {
                            isowner = false;
                            //lblxsuser.Tag = "";
                        }
                    }

                    //if stage is back checking then enable recheck




                    //if (stage_displayname == "Checking")
                    //{
                    //    //get node color
                    //    System.Drawing.Color nodecolor = mynode.ForeColor;
                    //    if (nodecolor != System.Drawing.Color.Red)
                    //    {
                    //        isowner = true;
                    //        lblxsuser.Tag = "Owner";
                    //    }
                    //}
                    //else
                    //{
                    //    isowner = true;
                    //    lblxsuser.Tag = "Owner";
                    //}

                    //check node tag and modify id stage if null
                    if (mynode.Tag == null)
                    {
                        id_stage = -1;
                    }
                    else
                    {
                        string[] tmp = mynode.Tag.ToString().Split(new Char[] { '|' });
                        //id_process = Convert.ToInt32(tmp[0]);
                        //id_stage = Convert.ToInt32(tmp[1]);

                        //tvwnewnode.Tag = SK_id_report.ToString() + "|" + SK_id_process;


                        SK_id_report = Convert.ToInt32(tmp[0]);
                        SK_id_process = Convert.ToInt32(tmp[1]);

                    }

                    btncomplete.Text = "Completed";
                    mytooltip.SetToolTip(this.btncomplete, "Completed");

                    btn1.Enabled = false;
                    btn2.Enabled = false;
                    btn3.Enabled = false;
                    lblqmsdblist.Enabled = false;
                    cmbqmsdblist.Enabled = false;
                    btncomplete.Enabled = false;
                    btncomplete.Enabled = false;
                    if (isowner == true)
                    {
                        if (id_process < (int)SKProcess.Modelling_PM_Approved)
                        {
                            //new
                            lblqmsdblist.Text = "Report";
                            btn1.Text = "+";
                            mytooltip.SetToolTip(this.btn1, "Add");
                            btn1.Enabled = false;
                            btn2.Text = "-";
                            mytooltip.SetToolTip(this.btn2, "Delete");
                            btn2.Enabled = true;
                            btn3.Text = "REN";
                            mytooltip.SetToolTip(this.btn3, "Rename");
                            btn3.Enabled = true;
                            lblqmsdblist.Enabled = true;
                            cmbqmsdblist.Enabled = true;
                            cmbqmsdblist.Text = report;
                            btncomplete.Enabled = true;
                        }
                        else if (id_process < (int)SKProcess.ModelChecking_Completed)
                        {
                            //checking in progress
                            btn2.Text = "-";
                            mytooltip.SetToolTip(this.btn2, "Delete");
                            btn2.Enabled = true;
                            btn3.Text = "REN";
                            mytooltip.SetToolTip(this.btn3, "Rename");
                            btn3.Enabled = true;
                            lblqmsdblist.Enabled = true;
                            cmbqmsdblist.Enabled = true;
                            cmbqmsdblist.Text = report;
                            //cmbqmsdblist.Enabled = true;
                            btncomplete.Enabled = true;
                            //btnfetch.Enabled = true;
                        }
                        else if (id_process >= (int)SKProcess.ModelChecking_Completed)
                        {
                            //checking in progress
                            btn2.Text = "";
                            mytooltip.SetToolTip(this.btn2, "");
                            btn2.Enabled = false;
                            btn3.Text = "";
                            mytooltip.SetToolTip(this.btn3, "");
                            btn3.Enabled = false;
                            lblqmsdblist.Enabled = true;
                            cmbqmsdblist.Enabled = true;
                            cmbqmsdblist.Text = report;
                            //cmbqmsdblist.Enabled = true;
                            btncomplete.Enabled = true;
                        }
                    }
                }
                else
                {
                    btn1.Enabled = false;
                    btn2.Enabled = false;
                    btn3.Enabled = false;
                    lblqmsdblist.Enabled = false;
                    cmbqmsdblist.Enabled = false;
                    btncomplete.Enabled = false;

                    if (stage_displayname == "New")
                    {
                        btn1.Text = "+";
                        mytooltip.SetToolTip(this.btn1, "Add");
                        btn1.Enabled = true;
                        lblqmsdblist.Enabled = true;
                        cmbqmsdblist.Enabled = true;
                        cmbqmsdblist.Text = "";
                        isowner = true;
                    }
                }
                //if selected node text differs with tab page qc name then clear all rows
                if ((mynode.Text != tpgqc.Text) || (lblstage.Text != stage_displayname))
                {
                    dgqc.Rows.Clear();
                    //SetControlVisible_QC(false);
                    btnclear.Tag = false;
                    btnSave.Tag = false;
                    btnAdd.Tag = false;
                    btnDelete.Tag = false;
                    btnUnique.Tag = false;
                    btnnextprocess.Tag = false;
                    btntspreview2.Tag = false;
                    btnfilter.Tag = false;
                    btnexport.Tag = false;
                    btnrefresh.Tag = false;
                    btnTSAll.Tag = true;
                    btnTSClear.Tag = true;
                   // btnpreview.Tag = true;
                    btncopy.Tag = false;
                    btnsplitsave.Tag = false;
                    chkzoom.Tag = false;
                    chkreplace.Tag = false;
                    btnpreviousprocess.Tag = false;
                    btnnextprocess.Tag = false;
                    btnrebackdraft.Tag = false;
                    btnview.Tag = false;

                }

                Update_Stage_ReportName(report, stage_displayname, id_stage);



            }
        }
        //private void tvwncm_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        //{
        //    isowner = false;

        //    //get the selected node from tree view
        //    TreeNode mynode = tvwncm.SelectedNode;
        //    if (mynode != null)
        //    {
        //        //get selected node full path 
        //        string tvwpath = mynode.FullPath;
        //        string[] tvwdata = tvwpath.Split(new Char[] { '+' });

        //        string stage_displayname = string.Empty;
        //        if (tvwdata.Length >= 3)
        //            stage_displayname = tvwdata[2];

        //        string report = "";

        //        //if length is 4 then user has selected one of the process (i.e checking, back draft, back check,.....)
        //        if (tvwdata.Length == 4)
        //        {
        //                report = tvwdata[3].ToString();
        //            //btn2.Enabled = true;
        //            //btn3.Enabled = true;
        //            //btncomplete.Enabled = true; 
        //            //get node color
        //            System.Drawing.Color nodecolor = mynode.ForeColor;



        //            //check node tag and modify id stage if null
        //            if (mynode.Tag == null)
        //            {
        //                id_stage = -1;
        //            }
        //            else
        //            {
        //                string[] tmp = mynode.Tag.ToString().Split(new Char[] { '|' });
        //                id_process = Convert.ToInt32(tmp[0]);
        //                id_stage = Convert.ToInt32(tmp[1]);
        //            }



        //            //Update_Stage_ReportName(tvwdata[tvwdata.Length - 1].ToString(), tvwdata[tvwdata.Length - 2].ToString(), Convert.ToInt32(mynode.Tag));

        //            //if node fore color is blue then current user is the owner set user interface accordinlgly
        //            if (nodecolor == System.Drawing.Color.Red)
        //            {
        //                lblxsuser.Tag = "";
        //            }
        //            else
        //            {
        //                isowner = true;
        //                lblxsuser.Tag = "Owner";
        //                btncomplete.Text = "Completed";


        //                btn1.Enabled = false;
        //                btn2.Enabled = false;
        //                btn3.Enabled = false;
        //                lblqmsdblist.Enabled = false;
        //                cmbqmsdblist.Enabled = false;
        //                btncomplete.Enabled = false;
        //                btncomplete.Enabled = false;



        //                if (id_process < (int)SKProcess.Modelling_PM_Approved)
        //                {
        //                    //new
        //                    btn1.Text = "+";
        //                    btn1.Enabled = false;
        //                    btn2.Text = "-";
        //                    btn2.Enabled = true;
        //                    btn3.Text = "REN";
        //                    btn3.Enabled = true;
        //                    lblqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Text = report;
        //                }
        //                else if (id_process < (int)SKProcess.ModelChecking_Completed)
        //                {
        //                    //checking in progress
        //                    btn2.Text = "-";
        //                    btn2.Enabled = true;
        //                    btn3.Text = "REN";
        //                    btn3.Enabled = true;
        //                    lblqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Text = report;
        //                    cmbqmsdblist.Enabled = true;
        //                    btncomplete.Enabled = true;
        //                    //btnfetch.Enabled = true;
        //                }
        //                else if (id_process >= (int)SKProcess.ModelChecking_Completed)
        //                {
        //                    //checking in progress
        //                    btn2.Text = "";
        //                    btn2.Enabled = false;
        //                    btn3.Text = "";
        //                    btn3.Enabled = false;
        //                    lblqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Enabled = true;
        //                    cmbqmsdblist.Text = report;
        //                    cmbqmsdblist.Enabled = true;
        //                    btncomplete.Enabled = true;
        //                }
        //                else if (id_process > 500)
        //                {
        //                    //complete
        //                }

        //                //if (stage_displayname == "New")
        //                //{
        //                //    btn1.Text = "+";
        //                //    btn1.Enabled = false;
        //                //    btn2.Text = "-";
        //                //    btn2.Enabled = true;
        //                //    btn3.Text = "REN";
        //                //    btn3.Enabled = true;
        //                //    lblqmsdblist.Enabled = true;
        //                //    cmbqmsdblist.Enabled = true;
        //                //    cmbqmsdblist.Text = report;
        //                //}
        //                //else if (stage_displayname == "Checking")
        //                //{

        //                //    btn2.Text = "-";
        //                //    btn2.Enabled = true;
        //                //    btn3.Text = "REN";
        //                //    btn3.Enabled = true;
        //                //    lblqmsdblist.Enabled = true;
        //                //    cmbqmsdblist.Enabled = true;
        //                //    cmbqmsdblist.Text = report;
        //                //    cmbqmsdblist.Enabled = true;
        //                //    btncomplete.Enabled = true;
        //                //    //btnfetch.Enabled = true;
        //                //}
        //                //else if (stage_displayname == "Completed")
        //                //{
        //                //    //btncomplete.Visible = false;
        //                //    //lblqmsdblist.Visible = false;
        //                //    //cmbqmsdblist.Visible = false;
        //                //}
        //                //btnSave.Enabled = true;
        //                //btnsplitsave.Enabled = true;
        //                //lblqmsdblist.Visible = true;
        //                //cmbqmsdblist.Visible = true;
        //                ////btn1.Visible = true;

        //            }
        //            //else
        //            //{

        //            //    lblxsuser.Tag = "";
        //            //    //if other scope then user can have read only access and can't modify
        //            //    //btnSave.Enabled = false;
        //            //    //btnAdd.Enabled = false;
        //            //    //btnDelete.Enabled = false;
        //            //    //btnsplitsave.Enabled = false;
        //            //}
        //        }
        //        else
        //        {
        //            btn1.Enabled = false;
        //            btn2.Enabled = false;
        //            btn3.Enabled = false;
        //            lblqmsdblist.Enabled = false;
        //            cmbqmsdblist.Enabled = false;
        //            btncomplete.Enabled = false;
        //            if (stage_displayname == "New")
        //            {
        //                btn1.Text = "+";
        //                btn1.Enabled = true;
        //                lblqmsdblist.Enabled = true;
        //                cmbqmsdblist.Enabled = true;
        //                cmbqmsdblist.Text = "";
        //                isowner = true;
        //            }


        //            //report = "";
        //            //stage_displayname = "";
        //            //mynode.Tag = "";
        //            //id_stage = -1;
        //            //Update_Stage_ReportName("", "", -1);
        //        }

        //        //if selected node text differs with tab page qc name then clear all rows
        //        if (mynode.Text != tpgqc.Text)
        //        {
        //            dgqc.Rows.Clear();
        //            //SetControlVisible_QC(false);
        //            btnclear.Tag = false;
        //            btnSave.Tag = false;
        //            btnAdd.Tag = false;
        //            btnDelete.Tag = false;
        //            btnUnique.Tag = false;
        //            btntspreview1.Tag = false;
        //            btntspreview2.Tag = false;
        //            btnfilter.Tag = false;
        //            btnexport.Tag = false;
        //            btncopy.Tag = false;
        //            btnsplitsave.Tag = false;
        //            chkzoom.Tag = false;
        //            chkreplace.Tag = false;

        //        }

        //        Update_Stage_ReportName(report, stage_displayname, id_stage);



        //    }
        //}

        //private void tvwncm_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        //{

        //    //tabpage dash board
        //    //SetControlVisible_DashBoard(false);

        //    //set visible off
        //    //ControlVisible_DashBoardAndQC(false);

        //    //set owner as false
        //    isowner = false;

        //    //get the selected node from tree view
        //    TreeNode mynode = tvwncm.SelectedNode;
        //    if (mynode != null)
        //    {
        //        //get node color
        //        System.Drawing.Color nodecolor = mynode.ForeColor;

        //        //get selected node full path 
        //        string tvwpath = mynode.FullPath;
        //        string[] tvwdata = tvwpath.Split(new Char[] { '+' });

        //        //check node tag and modify id stage if null
        //        if (mynode.Tag == null)
        //        {
        //            id_stage = -1;
        //        }

        //        string stage_displayname = string.Empty;
        //        if (tvwdata.Length >= 3)
        //            stage_displayname = tvwdata[2];

        //        string report = "";




        //        //if length is 4 then user has selected one of the process (i.e checking, back draft, back check,.....)
        //        if (tvwdata.Length == 4)
        //        {
        //            report = tvwdata[3].ToString();
        //            //btn2.Enabled = true;
        //            //btn3.Enabled = true;
        //            //btncomplete.Enabled = true;                   

        //            if (mynode.Tag != null)
        //            {
        //                string[] tmp = mynode.Tag.ToString().Split(new Char[] { '|' });
        //                id_process = Convert.ToInt32(tmp[0]);
        //                id_stage = Convert.ToInt32(tmp[1]);
        //            }



        //            //Update_Stage_ReportName(tvwdata[tvwdata.Length - 1].ToString(), tvwdata[tvwdata.Length - 2].ToString(), Convert.ToInt32(mynode.Tag));

        //            //if node fore color is blue then current user is the owner set user interface accordinlgly
        //            if (nodecolor == System.Drawing.Color.Blue)
        //            {
        //                isowner = true;

        //                lblxsuser.Tag = "Owner";
        //                btncomplete.Text = "Completed";
        //                //btncomplete.Visible = true;
        //                btncomplete.Enabled = true;

        //                if (stage_displayname == "New")
        //                {
        //                    btn2.Text = "-";
        //                    //btn1.Visible = true;
        //                    btn2.Enabled = true;
        //                    btn3.Text = "REN";
        //                    //btn1.Visible = true;
        //                    btn3.Enabled = true;
        //                    //lblqmsdblist.Visible = true;
        //                    //cmbqmsdblist.Visible = true;
        //                    cmbqmsdblist.Text = "";
        //                    //btncomplete.Visible = false;
        //                }
        //                else if (stage_displayname == "Checking")
        //                {
        //                    btn2.Text = "-";
        //                    //btn1.Visible = true;
        //                    btn2.Enabled = true;
        //                    btn3.Text = "REN";
        //                    //btn1.Visible = true;
        //                    btn3.Enabled = true;
        //                    //lblqmsdblist.Visible = true;
        //                    //cmbqmsdblist.Visible = true;
        //                    cmbqmsdblist.Text = report;
        //                    cmbqmsdblist.Enabled = true;
        //                }
        //                else if (stage_displayname == "Completed")
        //                {
        //                    //btncomplete.Visible = false;
        //                    //lblqmsdblist.Visible = false;
        //                    //cmbqmsdblist.Visible = false;
        //                }
        //                //btnSave.Enabled = true;
        //                //btnsplitsave.Enabled = true;
        //                //lblqmsdblist.Visible = true;
        //                //cmbqmsdblist.Visible = true;
        //                ////btn1.Visible = true;

        //            }
        //            else
        //            {
        //                lblxsuser.Tag = "";
        //                //if other scope then user can have read only access and can't modify
        //                //btnSave.Enabled = false;
        //                //btnAdd.Enabled = false;
        //                //btnDelete.Enabled = false;
        //                //btnsplitsave.Enabled = false;
        //            }
        //        }
        //        else
        //        {
        //            if (stage_displayname == "New")
        //            {
        //                btn1.Text = "+";
        //                //btn1.Visible = true;
        //                //lblqmsdblist.Visible = true;
        //                //cmbqmsdblist.Visible = true;
        //                cmbqmsdblist.Text = "";
        //            }

        //            //btn2.Enabled = false;
        //            //btn3.Enabled = false;
        //            //btncomplete.Enabled = false;
        //            //report = "";
        //            //stage_displayname = "";
        //            mynode.Tag = "";
        //            id_stage = -1;
        //            //Update_Stage_ReportName("", "", -1);
        //        }

        //        //if selected node text differs with tab page qc name then clear all rows
        //        if (mynode.Text != tpgqc.Text)
        //        {
        //            dgqc.Rows.Clear();
        //            SetControlVisible_QC(false);
        //        }

        //        Update_Stage_ReportName(report, stage_displayname, id_stage);



        //    }
        //    else
        //    {
        //        //no node selected do coding for this
        //    }



        //}

        //private void tvwncm_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        //{
        //    //get the selected node from tree view
        //    //TreeNode mynode = e.Node;

        //    TreeNode mynode = tvwncm.SelectedNode;
        //    System.Drawing.Color nodecolor = mynode.ForeColor;

        //    if (mynode != null)
        //    {
        //        //you will get selected tree view item here 
        //        string tvwpath = mynode.FullPath;
        //        string[] tvwdata = tvwpath.Split(new Char[] { '+' });
        //        string report = "";
        //        string stage_displayname = string.Empty;
        //        if (mynode.Tag == null)
        //        {
        //            id_stage = -1;
        //        }
        //        if (tvwdata.Length == 4)
        //        {
        //            btn2.Enabled = true;
        //            btn3.Enabled = true;
        //            btncomplete.Enabled = true;
        //            stage_displayname = tvwdata[2];
        //            if (mynode.Tag != null)
        //            {
        //                string[] tmp = mynode.Tag.ToString().Split(new Char[] { '|' });
        //                id_process = Convert.ToInt32(tmp[0]);
        //                id_stage = Convert.ToInt32(tmp[1]);
        //            }
        //            report = tvwdata[3].ToString();
        //            //Update_Stage_ReportName(tvwdata[tvwdata.Length - 1].ToString(), tvwdata[tvwdata.Length - 2].ToString(), Convert.ToInt32(mynode.Tag));

        //            //if node fore color is blue then current user is the owner set user interface accordinlgly
        //            if (nodecolor == System.Drawing.Color.Blue)
        //            {
        //                lblxsuser.Tag = "Owner";
        //                btnSave.Enabled = true;
        //                if (id_process < (int)SKProcess.ModelChecking_Completed)
        //                {
        //                    btnAdd.Enabled = true;
        //                    btnDelete.Enabled = true;
        //                }
        //                else
        //                {
        //                    btnAdd.Enabled = false;
        //                    btnDelete.Enabled = false;
        //                }

        //                btnsplitsave.Enabled = true;
        //            }
        //            else
        //            {
        //                lblxsuser.Tag = "";
        //                //if other scope then user can have read only access and can't modify
        //                btnSave.Enabled = false;
        //                btnAdd.Enabled = false;
        //                btnDelete.Enabled = false;
        //                btnsplitsave.Enabled = false;
        //            }

        //            //for completed process there is no requirement for save, select, delete and split hence disabled
        //            if (id_process == 500)
        //            {
        //                btnSave.Enabled = false;
        //                btnAdd.Enabled = false;
        //                btnDelete.Enabled = false;
        //                btnsplitsave.Enabled = false;
        //            }







        //        }
        //        else
        //        {
        //            btn2.Enabled = false;
        //            btn3.Enabled = false;
        //            btncomplete.Enabled = false;
        //            //report = "";
        //            stage_displayname = "";
        //            mynode.Tag = "";
        //            id_stage = -1;
        //            //Update_Stage_ReportName("", "", -1);
        //        }

        //        Update_Stage_ReportName(report, stage_displayname, id_stage);
        //    }

        //}




        private void btncomplete_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "CompleteTask", "", "Pl.Wait,saving data's", "btncomplete_Click", lblsbar1, lblsbar2);

            //tpgdash.Tag = cmbqmsdblist.Text;
            //move one stage to next stage based on role and current process
            //Update_DB_MovetoNextProcess(emp_role, id_process);
            MoveReport_NextProcess(SK_id_report, SK_id_process, tvwncm.SelectedNode.Text);

            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "CompleteTask", "", "", "btncomplete_Click", lblsbar1, lblsbar2, s_tm);



        }

        private void tpgQC_Enter(object sender, EventArgs e)
        {
            //set all control visible as false
            //SetControlVisible_QC(false);

            //if (tpgqc.Text.ToUpper().Trim() != skreportname.ToUpper().Trim() )
            //{
            //    dgqc.Rows.Clear();
            //}


            //if active tab page is ncm set the user inferface accordingly
            //SetUserInterface_QC(id_process);

            //////set enable or disable of button based on row count in datagrid qc
            //setpreview();

        }

        private void setControlVisible(System.Windows.Forms.Control mycontrol, bool myVisible = true)
        {
            mycontrol.Visible = myVisible;
        }
        private void Update_Control()
        {
            setControlVisible(btn1);
            setControlVisible(btn2);
            setControlVisible(btn3);

            setControlVisible(btnclear);
            setControlVisible(btnsplitsave);
            setControlVisible(btnpreviousprocess);
            setControlVisible(btnnextprocess);
            setControlVisible(btnrebackdraft);


            setControlVisible(btnnextprocess);
            setControlVisible(btntspreview2);
            setControlVisible(btncomplete);
            setControlVisible(btncopy);
            setControlVisible(btnDelete);
            //setControlVisible(btnexport);
            setControlVisible(btnfilter);
            setControlVisible(btnfetch);
            //setControlVisible(btnhlp);
            setControlVisible(btnLoad);
            setControlVisible(btnSave);
            setControlVisible(btnAdd);
            setControlVisible(btnview);


            setControlVisible(btnUnique);
            setControlVisible(chkzoom);
            setControlVisible(chkzoom);
        }

        private void SetControlVisible_DashBoard(bool myVisible)
        {
            //excel export is part of dashboard hence enable
            btnexport.Enabled = true;
            btnrefresh.Enabled = true;
            btnTSAll.Enabled = true;
            btnTSClear.Enabled = true;
            //btnpreview.Enabled = true;
            //lblqmsdblist.Visible = myVisible;
            //cmbqmsdblist.Visible = myVisible;
            //btn1.Visible = myVisible;
            //btn2.Visible = myVisible;
            //btn3.Visible = myVisible;
            //btncomplete.Visible = myVisible;
        }

        private void SetControlVisible_QC(bool myVisible)
        {
            ////btnclear.Visible = myVisible;
            //btnfetch.Visible = myVisible;
            //btnSave.Visible = myVisible;
            //btnAdd.Visible = myVisible;
            //btnDelete.Visible = myVisible;
            //btnUnique.Visible = myVisible;
            //btntspreview1.Visible = myVisible;
            //btntspreview2.Visible = myVisible;
            //btnfilter.Visible = myVisible;
            //btnexport.Visible = myVisible;
            //btncopy.Visible = myVisible;
            //btnsplitsave.Visible = myVisible;
            //chkzoom.Visible = myVisible;
            //chkreplace.Visible = myVisible;
        }





        private void SetControlVisible_QC_DataGridRowCount(int process = -1)
        {
            //fetch records will always be true
            //btnfetch.Visible = true;

            //excel export is part of dashboard hence disable during qc task
            btnexport.Enabled = false;
            btnrefresh.Enabled = false;
            btnTSAll.Enabled = false;
            btnTSClear.Enabled = false;
            btnview.Enabled = true;
            //btnpreview.Enabled = false;

            //set visible
            bool flag = false;
            if (dgqc.Rows.Count >= 1)
            {
                //since there is one or more row's set flag as visible
                flag = true;

            }

            btnclear.Enabled = flag;
            btnUnique.Enabled = flag;
            btnnextprocess.Enabled = flag;
            btntspreview2.Enabled = flag;
            btnfilter.Enabled = flag;
            //btnexport.Enabled = flag;
            btncopy.Enabled = flag;
            chkzoom.Enabled = flag;
            btnpreviousprocess.Enabled = flag;
            btnnextprocess.Enabled = flag;
            chkzoom.Enabled = flag;

            //if stage is new or checking then ensure selection of tekla objects is enabled.
            if (lblstage.Text == "New")
            {
                btnAdd.Enabled = isowner;
                btnpreviousprocess.Enabled = false;
            }
            else
            {
                btnAdd.Enabled = false;
                if (lblstage.Text == "QA" || lblstage.Text == "COMPLETED" || lblstage.Text == "HELD")
                {
                    btnpreviousprocess.Enabled = false;
                    btnnextprocess.Enabled = false;
                }
            }

            if (dgqc.Rows.Count >= 1)
            {
                //set delete based on btnadd setting
                btnDelete.Enabled = btnAdd.Enabled;
            }
            else
                btnDelete.Enabled = false;

            //if row count is more than 1 ensure save button is enabled except for completed
            if (dgqc.Rows.Count >= 1)
            {
                if (id_process >= (int)SKProcess.QA)
                    btnSave.Enabled = false;
                else
                    btnSave.Enabled = isowner;
            }
            else
            {
                btnSave.Enabled = false;
            }



            //if row count is more than 2 split save can be enabled based on process owner
            if (dgqc.Rows.Count >= 2)
            {
                if (id_process >= (int)SKProcess.Completed)
                    btnsplitsave.Enabled = false;
                else
                    btnsplitsave.Enabled = isowner;
            }
            else
            {
                btnsplitsave.Enabled = false;
            }
            //based on owner condition set btn1,2,3,4
            if (isowner == false)
            {
                btn1.Enabled = isowner;
                btn2.Enabled = isowner;
                btn3.Enabled = isowner;
                btncomplete.Enabled = isowner;
            }

            //set previous stage
            //btnpreviousprocess.Enabled = btnSave.Enabled;
            //btnnextprocess.Enabled = btnSave.Enabled;
            btnrebackdraft.Enabled = btnSave.Enabled;

        }

        private void ControlVisible_DashBoardAndQC(bool myVisible)
        {
            //tabpage dash board
            SetControlVisible_DashBoard(myVisible);

            //tabpage qc
            SetControlVisible_QC(myVisible);
        }



        private void btnget_Click(object sender, EventArgs e)
        {
            //check whether null entry before fetching records
            //if (cmbqmsdblist.Text.Trim().Length >= 1)
            if (tpgqc.Text.Trim().Length >= 1)
            {
                //skreportname = cmbqmsdblist.Text;
                skreportname = tpgqc.Text;
                //tpgqc.Tag = cmbqmsdblist.Text;;
                //start log
                DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "Fetch", "", "Pl.Wait,fetching...", "btnget_Click", lblsbar1, lblsbar2);

                //fetch data from mysql
                string msg = Fetch_Model_Data();


                //dgqc.
                //store mirror data for future comparision
                MirrorTabPageData();

                //based on row count set control visiblity
                SetControlVisible_QC_DataGridRowCount();


                //dgqc.Controls[0].Enabled = true; // Index zero is the horizontal scrollbar
                //dgqc.Controls[1].Enabled = true; // Index one is the vertical scrollbar

                //for back checking stage alone enable back check to back draft process button
                if (lblstage.Text  == "Back Checking")
                    btnrebackdraft.Enabled = isowner;
                else
                    btnrebackdraft.Enabled = false;

                dgqc.Columns["mod_user_id"].Width = 50;

                skTSLib.setCompletion(skApplicationName, skApplicationVersion, "Fetch", "", msg, "btnget_Click", lblsbar1, lblsbar2, s_tm);
            }
        }


        //for future performance increase
        //new table for fetching records
        //DataTable mytable = new DataTable();

        ////add fixed column
        //mytable.Columns.Add("guid");
        //mytable.Columns.Add("guid_type");
        //mytable.Columns.Add("idqc");
        //mytable.Columns.Add("id_sk_process");

        //mysql = ",";
        //foreach (DataGridViewColumn mycolumn in dgqc.Columns)
        //{

        //if (mycolumn.Visible == true)
        //{
        //mytable.Columns.Add(mycolumn.Name);
        ////based on column name set sql name
        //mysql = mysql + ", " + mycolumn.Tag.ToString();
        //}

        //}
        //mysql = mysql.Replace(",,", "");
        //mysql = "select guid, guid_type, idqc, id_sk_process, " + mysql + " from epmatdb.ncm where status is null and id_report =" + id_stage + " order by idqc";


        ////get database values
        //if (myconn.State == ConnectionState.Closed)
        //myconn.Open();
        //mycmd.CommandText = mysql;
        //mycmd.Connection = myconn;
        //myrdr = mycmd.ExecuteReader();

        //while (myrdr.Read())
        //{
        //DataRow mydr = mytable.NewRow();
        //for (int i = 0; i < mytable.Columns.Count; i++)
        //{
        ////DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        ////cellcat.DataSource = qmsCategory;
        ////if (db_idcat != -1)
        ////    cellcat.Value = qmsCategory[db_idcat].ToString();


        ////int db_idsev = -1;
        ////if (db_sev != "" && db_sev.Trim().Length >= 1)
        ////    db_idsev = Convert.ToInt32(db_sev);

        ////DataGridViewComboBoxCell cellsev = (DataGridViewComboBoxCell)(row.Cells["severity"]);
        ////cellsev.DataSource = qmsSeverity;
        ////if (db_idsev != -1)
        ////    cellsev.Value = qmsSeverity[db_idsev].ToString();

        ////owner
        ////if (db_owner.Length >= 1)
        ////{
        ////    ArrayList tmp = getDBUserID(Convert.ToInt32(db_owner));
        ////    if (tmp.Count >= 1)
        ////    {
        ////        string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
        ////        row.Cells["owner"].Value = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]";
        ////    }
        ////}

        //mydr[mytable.Columns[i].ColumnName] = skWinLib.mysql_set_quotecomma(myrdr[i].ToString());
        //}
        //mytable.Rows.Add(mydr);
        //}

        //dgqc.DataSource = mytable;
        private string Fetch_Model_Data(bool viewmode = false)
        {
            skWinLib.DataGridView_Setting_Before(dgqc);


            //First Clear
            dgqc.Rows.Clear();

            //based on process get the data for the current_user and null entry also (no one performed task earlier hence free current user can commence the work but before saving ensure null entry in user_id else conflict between user entry)
            string wherecond = GetCurrentProcess_SelectQuery_Userid(SK_id_process);


            ////get before remarks column id
            string selectcond = GetCurrentProcess_BeforeRemark(SK_id_process);

            if (viewmode == true)
            {
                selectcond = "mod_rmk, chk_rmk, bd_rmk, backcheck_rmk, sev_rmk, pm_rmk";
                wherecond = "";
            }


            //get data's only for visible column apart from idqc, guid, guid_type, id_sk_process, beforeremark
            mysql = ",";
            foreach (DataGridViewColumn mycolumn in dgqc.Columns)
            {
                if (mycolumn.Visible == true)
                {
                    //check for teklask column
                    if (mycolumn.Name.Contains("TeklaSK") == false)
                    {
                        string ui_colname = mycolumn.Tag.ToString();

                        if (ui_colname.ToUpper() == "ST|BEFOREREMARK")
                        {
                            ui_colname = selectcond;
                        }
                        else
                        {
                            string sql_type_colname = mycolumn.Tag.ToString();
                            if (sql_type_colname.Contains("|") == true)
                            {
                                string[] tmp = sql_type_colname.Split(new Char[] { '|' });
                                //string sqlcoltype = tmp[0].ToString();
                                ui_colname = tmp[1].ToString();
                            }
                        }
                        if (ui_colname.Trim().Length >= 1)
                        {                        //based on column name set sql name
                            mysql = mysql + ", " + ui_colname;
                        }
                    }

                }

            }
            mysql = mysql.Replace(",,", "");
            mysql = "select guid, guid_type, idqc, id_sk_process," + mysql + " from epmatdb.ncm where status is null" + wherecond + " and id_report = " + SK_id_report + " order by idqc";

            //get database values
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            int dg_rowid = 1;

            TeklaAdditionalColumnFlag = IsTeklaAdditionalEnabled();
            while (myrdr.Read())
            {
                //add new row
                dg_rowid = dgqc.Rows.Add(myrdr[0]);
                DataGridViewRow row = this.dgqc.Rows[dg_rowid];
                if (chkPerformance.Checked == false)
                {
                    Identifier ID = new Identifier(myrdr[0].ToString());
                    if (ID.IsValid() == true)
                    {
                        TSM.ModelObject MOB = MyModel.SelectModelObject(ID);
                        row.Cells["TeklaSKtype"].Value = MOB.GetType().Name.ToString();
                        if (TeklaAdditionalColumnFlag == true)
                        {
                            update_Tekla_AdditionalAttributes(MOB, row);
                        }
                    }
                }
                else
                    row.Cells["TeklaSKtype"].Value = myrdr[1].ToString();

                row.Cells["idqc"].Value = myrdr[2].ToString();
                row.Cells["stage"].Value = myrdr[3].ToString();
                //break;
                int dbcolct = 4;
                foreach (DataGridViewColumn mycolumn in dgqc.Columns)
                {
                    if (mycolumn.Visible == true)
                    {
                        if (mycolumn.Name.Contains("TeklaSK") == false)
                        {
                            string ui_colname = mycolumn.Tag.ToString();
                            string db_colname = myrdr.GetName(dbcolct);
                            string before_remark = "";
                            if (ui_colname.ToUpper() == "ST|BEFOREREMARK")
                            {
                                if (SK_id_process == (int)SKProcess.Modelling_PM_Approved)
                                {
                                    before_remark = myrdr[6].ToString() + ", " + myrdr[7].ToString();
                                }
                                else if (SK_id_process == (int)SKProcess.ModelChecking_PM_Approved)
                                {
                                    before_remark = myrdr[9].ToString() + ", " + myrdr[10].ToString() + ", " + myrdr[11].ToString();
                                }
                                else if (SK_id_process == (int)SKProcess.ModelBackDraft_PM_Approved)
                                {
                                    before_remark = myrdr[11].ToString() + ", " + myrdr[12].ToString() + ", " + myrdr[13].ToString() + ", " + myrdr[14].ToString();
                                }
                                else if (SK_id_process == (int)SKProcess.ModelBackCheck_PM_Approved)
                                {
                                    before_remark = myrdr[8].ToString() + ", " + myrdr[9].ToString() + ", " + myrdr[10].ToString() + ", " + myrdr[11].ToString(); ;
                                }
                                else
                                {
                                    before_remark = myrdr[10].ToString() + ", " + myrdr[11].ToString() + ", " + myrdr[12].ToString() + ", " + myrdr[13].ToString() + ", " + myrdr[14].ToString();
                                }
                                before_remark = before_remark.Replace(", ,", ",,");
                                before_remark = before_remark.Replace(", ,", ",,");
                                before_remark = before_remark.Replace(", ,", ",,");
                                before_remark = before_remark.Replace(", ,", ",,");
                                before_remark = before_remark.Replace(", ,", ",,");
                                before_remark = before_remark.Replace(",,", ",");
                                before_remark = before_remark.Replace(",,", ",");
                                before_remark = before_remark.Replace(",,", ",");
                                before_remark = before_remark.Replace(",,", ",");
                                before_remark = before_remark.Replace(",,", ",");
                                before_remark = before_remark.Replace(",,", ",");

                                if (before_remark != ",")
                                    row.Cells[mycolumn.Name].Value = skWinLib.mysql_set_quotecomma(before_remark);
                            }
                            else if (ui_colname.ToUpper() == "IN|BD_USER_ID")
                            {
                                string empid = myrdr["bd_user_id"].ToString();
                                if (empid.Trim().Length >= 1)
                                {
                                    ArrayList tmp = getDBUserID(Convert.ToInt32(empid));
                                    if (tmp.Count >= 1)
                                    {
                                        string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                                        row.Cells[mycolumn.Name].Value = empsplit[2];
                                    }
                                }

                            }
                            else if (ui_colname.ToUpper() == "IN|MOD_USER_ID") 
                            {
                                string empid = myrdr["mod_user_id"].ToString();
                                if (empid.Trim().Length >= 1)
                                {
                                    row.Cells[mycolumn.Name].Value = empid;
                                }

                            }
                            else if (ui_colname.ToUpper() == "IN|BD_ISAGREE")
                            {
                                string db_bdisagree = myrdr["bd_isagree"].ToString();
                                //string db_bdisagree = myrdr[dbcolct].ToString();
                                int db_isagreed = 0;
                                if (db_bdisagree == "1")
                                    db_isagreed = Convert.ToInt32(db_bdisagree);

                                row.Cells["IsAgreed"].Value = db_isagreed;


                            }
                            else if (ui_colname.ToUpper() == "IN|BD_ISMODELUPDATED")
                            {
                                string db_bdismodelupdated = myrdr["bd_ismodelupdated"].ToString();
                                //string db_bdismodelupdated = myrdr[dbcolct].ToString();
                                int ismodelupdated = 0;
                                if (db_bdismodelupdated == "1")
                                    ismodelupdated = Convert.ToInt32(db_bdismodelupdated);

                                row.Cells["MODELUPDATED"].Value = ismodelupdated;


                            }
                            else if (ui_colname.ToUpper() == "IN|BACKCHECK")
                            {
                                string db_bdismodelupdated = myrdr["backcheck"].ToString();
                                //string db_bdismodelupdated = myrdr[dbcolct].ToString();
                                int ismodelupdated = 0;
                                if (db_bdismodelupdated == "1")
                                    ismodelupdated = Convert.ToInt32(db_bdismodelupdated);

                                row.Cells["BACKCHECK"].Value = ismodelupdated;


                            }
                            else if (ui_colname.ToUpper() == "IN|ID_CAT")
                            {
                                string db_cat = myrdr["id_cat"].ToString();
                                // string db_cat = myrdr[dbcolct].ToString();
                                int db_idcat = -1;
                                if (db_cat != "" && db_cat.Trim().Length >= 1)
                                    db_idcat = Convert.ToInt32(db_cat);

                                DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
                                cellcat.DataSource = qmsCategory;
                                if (db_idcat != -1)
                                    cellcat.Value = qmsCategory[db_idcat].ToString();



                            }
                            else if (ui_colname.ToUpper() == "IN|ID_SEV")
                            {
                                string db_sev = myrdr["id_sev"].ToString();
                                //string db_sev = myrdr[dbcolct].ToString();
                                int db_idsev = -1;
                                if (db_sev != "" && db_sev.Trim().Length >= 1)
                                    db_idsev = Convert.ToInt32(db_sev);

                                DataGridViewComboBoxCell cellsev = (DataGridViewComboBoxCell)(row.Cells["severity"]);
                                cellsev.DataSource = qmsSeverity;
                                if (db_idsev != -1)
                                    cellsev.Value = qmsSeverity[db_idsev].ToString();
                            }
                            else
                            {
                                //ui = myrdr[dbcolct].ToString();
                                row.Cells[mycolumn.Name].Value = skWinLib.mysql_set_quotecomma(myrdr[db_colname].ToString());
                                //row.Cells[mycolumn.Name].Value = skWinLib.mysql_set_quotecomma(myrdr[dbcolct].ToString());
                            }

                        }
                        dbcolct++;
                    }

                }
                dg_rowid++;
            }
            myrdr.Close();
            myconn.Close();

            //update row header
            skWinLib.updaterowheader(dgqc);

            skWinLib.DataGridView_Setting_After(dgqc);

            return dgqc.RowCount.ToString() + " record(s) found.";
        }

        //private string Fetch_Model_Data()
        //{
        //    //First Clear
        //    dgqc.Rows.Clear();

        //    //new table for fetching records
        //    DataTable mytable = new DataTable();

        //    //if new then no action on this
        //    if (lblstage.Text.ToUpper() != "NEW")
        //    {
        //        mysql = "";
        //        if (id_process <= (int)SKProcess.Modelling_PM_Approved)
        //        {
        //            //if (id_stage != -1)
        //            //{
        //            //    mysql = "select guid, guid_type, idqc, obser, chk_rmk from epmatdb.ncm where chg_rmk is null and project_id=" + id_project + " and chk_user_id =" + id_emp + " and id_ncm_progress =" + id_stage + " order by idqc";

        //            //}
        //            mysql = "select guid, guid_type, idqc, obser, chk_rmk from epmatdb.ncm where status is null and project_id=" + id_project + " and id_ncm_progress =" + id_stage + " order by idqc";

        //        }
        //        else if (id_process <= (int)SKProcess.ModelChecking_PM_Approved)
        //        {
        //            //back drafting
        //            mysql = "select guid, guid_type, idqc, obser, bd_rmk, bd_isagree, id_cat from epmatdb.ncm where status is null and project_id=" + id_project + " and id_ncm_progress =" + id_stage + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by idqc";

        //        }
        //        else if (id_process <= (int)SKProcess.ModelBackDraft_PM_Approved)
        //        {
        //            //back checking
        //            mysql = "select guid, guid_type, idqc, obser, chk_rmk, bd_isagree, id_cat, backcheck, backcheck_rmk from epmatdb.ncm where status is null and project_id=" + id_project + " and id_ncm_progress =" + id_stage + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by idqc";

        //        }
        //        else if (id_process >= (int)SKProcess.Completed)
        //        {
        //            //completed
        //            //mysql = "select guid, guid_type, idqc, obser, chk_rmk, bd_isagree, id_cat, id_sev, bd_rmk, backcheck, backcheck_rmk from epmatdb.ncm where chg_rmk is null and project_id=" + id_project + " and id_ncm_progress =" + id_stage + " and epmatdb.ncm.obser NOT LIKE '%" + sk_nocomment + "%' and epmatdb.ncm.obser is not null order by idqc";
        //            mysql = "select guid, guid_type, idqc, obser, chk_rmk, bd_isagree, id_cat, id_sev, bd_rmk, backcheck, backcheck_rmk from epmatdb.ncm where status is null and project_id=" + id_project + " and id_ncm_progress =" + id_stage + " order by idqc";

        //        }

        //        if (mysql != "")
        //        {
        //            if (myconn.State == ConnectionState.Closed)
        //                myconn.Open();
        //            mycmd.CommandText = mysql;
        //            mycmd.Connection = myconn;
        //            myrdr = mycmd.ExecuteReader();
        //            int dg_rowid = 0;
        //            sk_db_performance = null;

        //            //data table


        //            //data table




        //            while (myrdr.Read())
        //            {
        //                DataRow mydr = mytable.NewRow();


        //                string myrowdata = null;
        //                //fetch data from database
        //                //guid,guid_type,idqc, obser,chk_rmk 
        //                //check idqc is not null or empty space
        //                string db_guid = myrdr[0].ToString();
        //                string db_guid_type = myrdr[1].ToString();
        //                string db_idqc = myrdr[2].ToString();

        //                //DataRow mydatarow =  mytable.Rows.Add(db_guid);

        //                //add new row
        //                dg_rowid = dgqc.Rows.Add(db_guid);
        //                DataGridViewRow row = this.dgqc.Rows[dg_rowid];

        //                row.Cells["guid"].Value = db_guid;
        //                row.Cells["TeklaSKtype"].Value = db_guid_type;
        //                row.Cells["idqc"].Value = db_idqc;

        //                myrowdata = db_idqc + sk_seperator + db_guid + sk_seperator + db_guid_type + sk_seperator + db_idqc;



        //                if (id_process <= (int)SKProcess.Modelling_PM_Approved) //emp_role == "Checker" &&
        //                {
        //                    //fetch data from database for checking progress
        //                    string db_obser = myrdr[3].ToString();
        //                    string db_chk_rmk = myrdr[4].ToString();

        //                    row.Cells["observation"].Value = db_obser;
        //                    row.Cells["remark"].Value = db_chk_rmk;
        //                    row.Cells["stage"].Value = id_stage;
        //                }
        //                else if (id_process <= (int)SKProcess.ModelChecking_PM_Approved)
        //                {
        //                    //fetch data from database for back draft progress
        //                    //guid, guid_type, idqc, obser, bd_rmk, bd_isagree, id_cat 
        //                    string db_obser = myrdr[3].ToString();
        //                    string db_bd_rmk = myrdr[4].ToString();
        //                    string db_bdisagree = myrdr[5].ToString();
        //                    string db_cat = myrdr[6].ToString();


        //                    int db_isagreed = 0;
        //                    if (db_bdisagree == "1")
        //                        db_isagreed = Convert.ToInt32(db_bdisagree);

        //                    int db_idcat = -1;
        //                    if (db_cat != "" && db_cat.Trim().Length >= 1)
        //                        db_idcat = Convert.ToInt32(db_cat);


        //                    row.Cells["observation"].Value = db_obser;
        //                    row.Cells["remark"].Value = db_bd_rmk;
        //                    row.Cells["IsAgreed"].Value = db_isagreed;

        //                    DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        //                    cellcat.DataSource = qmsCategory;
        //                    if (db_idcat != -1)
        //                        cellcat.Value = qmsCategory[db_idcat].ToString();





        //                    row.Cells["stage"].Value = id_stage;
        //                }
        //                else if (id_process <= (int)SKProcess.ModelBackDraft_PM_Approved)
        //                {
        //                    //fetch data from database for back check progress
        //                    //select guid, guid_type, idqc, obser, chk_rmk, bd_isagree, id_cat, backcheck, backcheck_rmk
        //                    string db_obser = myrdr[3].ToString();
        //                    string db_chk_rmk = myrdr[4].ToString();
        //                    string db_bdisagree = myrdr[5].ToString();
        //                    string db_cat = myrdr[6].ToString();
        //                    string db_backcheck = myrdr[7].ToString();
        //                    string db_backcheck_rmk = myrdr[8].ToString();

        //                    int db_isagreed = 0;
        //                    if (db_bdisagree == "1")
        //                        db_isagreed = Convert.ToInt32(db_bdisagree);

        //                    int db_bc = 0;
        //                    if (db_backcheck == "1")
        //                        db_bc = Convert.ToInt32(db_backcheck);

        //                    int db_idcat = -1;
        //                    if (db_cat != "" && db_cat.Trim().Length >= 1)
        //                        db_idcat = Convert.ToInt32(db_cat);



        //                    row.Cells["observation"].Value = db_obser;
        //                    row.Cells["beforeremark"].Value = db_chk_rmk;
        //                    row.Cells["IsAgreed"].Value = db_isagreed;
        //                    row.Cells["backcheck"].Value = db_bc;
        //                    row.Cells["stage"].Value = id_stage;
        //                    row.Cells["remark"].Value = db_backcheck_rmk;

        //                    DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        //                    cellcat.DataSource = qmsCategory;
        //                    if (db_idcat != -1)
        //                        cellcat.Value = qmsCategory[db_idcat].ToString();

        //                }
        //                else if (id_process >= (int)SKProcess.Completed)
        //                {
        //                    //fetch data from database for review 
        //                    //guid, guid_type, idqc, obser, chk_rmk, bd_isagree, id_cat, id_sev, bd_rmk, backcheck, backcheck_rmk
        //                    string db_obser = myrdr[3].ToString();
        //                    string db_chk_rmk = myrdr[4].ToString();
        //                    string db_bdisagree = myrdr[5].ToString();
        //                    string db_cat = myrdr[6].ToString();
        //                    string db_sev = myrdr[7].ToString();
        //                    string db_bdrmk = myrdr[8].ToString();
        //                    string db_backcheck = myrdr[9].ToString();
        //                    string db_backcheck_rmk = myrdr[10].ToString();

        //                    int db_isagreed = 0;
        //                    if (db_bdisagree == "1")
        //                        db_isagreed = Convert.ToInt32(db_bdisagree);

        //                    int db_bc = 0;
        //                    if (db_backcheck == "1")
        //                        db_bc = Convert.ToInt32(db_backcheck);

        //                    int db_idcat = -1;
        //                    if (db_cat != "" && db_cat.Trim().Length >= 1)
        //                        db_idcat = Convert.ToInt32(db_cat);

        //                    int db_idsev = -1;
        //                    if (db_sev != "" && db_sev.Trim().Length >= 1)
        //                        db_idsev = Convert.ToInt32(db_sev);

        //                    row.Cells["observation"].Value = db_obser;
        //                    row.Cells["beforeremark"].Value = db_chk_rmk + " << " + db_bdrmk + " >>" + " << " + db_backcheck_rmk + " >>";
        //                    row.Cells["IsAgreed"].Value = db_isagreed;
        //                    row.Cells["backcheck"].Value = db_bc;
        //                    row.Cells["stage"].Value = id_stage;
        //                    row.Cells["remark"].Value = "";

        //                    DataGridViewComboBoxCell cellcat = (DataGridViewComboBoxCell)(row.Cells["Category"]);
        //                    cellcat.DataSource = qmsCategory;
        //                    if (db_idcat != -1)
        //                        cellcat.Value = qmsCategory[db_idcat].ToString();


        //                    DataGridViewComboBoxCell cellsev = (DataGridViewComboBoxCell)(row.Cells["severity"]);
        //                    cellsev.DataSource = qmsSeverity;
        //                    if (db_idsev != -1)
        //                        cellsev.Value = qmsSeverity[db_idsev].ToString();
        //                }
        //                //add all row value
        //                //sk_db_performance.Add(myrowdata);


        //                //string db_id_ncm_progress = lblstage.Tag;



        //                dgqc.Rows[dg_rowid].HeaderCell.Value = dgqc.Rows.Count.ToString();
        //            }


        //            myrdr.Close();
        //            myconn.Close();
        //        }
        //    }

        //    //dgqc.DataSource = mytable;
        //    dgqc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        //    dgqc.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders);
        //    dgqc.AutoResizeColumns();
        //    dgqc.AutoResizeRows();
        //    dgqc.Refresh();
        //    return dgqc.RowCount.ToString() + " record(s) found.";
        //}



        public static void UpdateDataGridselectedrows_Tekla_AdditionalField(DataGridView mydatagrid, int colid = 0, string idtype = "GUID")
        {
            try
            {
                //int rowindex = e.RowIndex;
                //DataGridViewRow row = mydatagrid.Rows[rowindex];
                ArrayList ObjectsToSelect = new ArrayList();
                int i = 0;
                foreach (DataGridViewRow row in mydatagrid.SelectedRows)
                {
                    //DataGridViewRow row = mydatagrid.SelectedRows[i]; 
                    Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier();
                    if (idtype.ToUpper().Trim() == "GUID")
                    {
                        string myguid = row.Cells[colid].Value.ToString();
                        myguid = myguid.ToUpper().Replace("GUID:", "").Replace("GUID", "").Trim();
                        myIdentifier = new Tekla.Structures.Identifier(myguid);
                    }
                    else if (idtype.ToUpper().Trim() == "ID")
                    {
                        string s_id = row.Cells[colid].Value.ToString();
                        s_id = s_id.ToUpper().Replace("ID:", "").Trim();
                        int id = Convert.ToInt32(s_id);
                        myIdentifier = new Tekla.Structures.Identifier(id);
                    }

                    bool flag = myIdentifier.IsValid();

                    if (flag)
                    {
                        Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
                        if (ModelSideObject.Identifier.IsValid())
                        {
                            // cal function for other tekla fields
                        }
                    }
                    i++;

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : " + ex.Message);
            }


        }

        private bool IsBackCheckingCompleted()
        {

            foreach (DataGridViewRow bcrow in dgqc.Rows)
            {
                if (bcrow.Cells["backcheck"].Value != null)
                {
                    string backcheck = bcrow.Cells["backcheck"].Value.ToString();
                    //(backcheck.ToUpper().Trim() != "FALSE") || (
                    if (backcheck.ToUpper().Trim() != "1")
                        return false;
                }

            }
            return true;
        }

        private void btnUnique_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "TS_Unique", "", "Pl.Wait,Unique...", "btnUnique_Click", lblsbar1, lblsbar2);
            //tpgqc.Tag = cmbqmsdblist.Text;;

            //process all rows and get unquie observation
            int unqct = 0;
            //int chkcolid = 2;
            int grpct = 0;
            //check whether save button to be disabled
            bool flag = true;

            //sort by guid
            dgqc.Sort(dgqc.Columns["guid"], ListSortDirection.Ascending);
            skWinLib.DataGridView_Setting_Before(dgqc);

            for (int i = 0; i < (dgqc.Rows.Count - 1); i++)
            {
                DataGridViewRow aboverow = dgqc.Rows[i];
                string aboverowguid = aboverow.Cells["guid"].Value.ToString();
                int skipct = 0;
                for (int j = i + 1; j < dgqc.Rows.Count; j++)
                {
                    DataGridViewRow belowrow = dgqc.Rows[j];
                    //string aboverowguid = string.Empty;
                    //string belowrowguid = string.Empty;
                    //if (aboverow.Cells[chkcolid].Value != null)
                    //string aboverowguid = aboverow.Cells["guid"].Value.ToString();
                    //if (belowrow.Cells[chkcolid].Value != null)
                    string belowrowguid = belowrow.Cells["guid"].Value.ToString();

                    //compare both 
                    if (aboverowguid.ToUpper().Trim() == belowrowguid.ToUpper().Trim())
                    {
                        string aboverowcellvalue = string.Empty;
                        string belowrowcellvalue = string.Empty;

                        if (aboverow.Cells["observation"].Value != null)
                            aboverowcellvalue = aboverow.Cells["observation"].Value.ToString();
                        if (belowrow.Cells["observation"].Value != null)
                            belowrowcellvalue = belowrow.Cells["observation"].Value.ToString();

                        //append value
                        if (aboverowcellvalue != string.Empty)
                        {
                            if (belowrowcellvalue != string.Empty)
                                aboverow.Cells["observation"].Value = aboverowcellvalue + " , \r\n" + belowrowcellvalue;
                        }
                        else
                        {
                            if (belowrowcellvalue != string.Empty)
                                aboverow.Cells["observation"].Value = belowrowcellvalue;
                        }

                        belowrow.Cells["observation"].Value = string.Empty;
                        flag = false;
                        skipct++;
                        unqct++;
                    }
                    else
                        break;




                }
                if (skipct >= 1)
                {
                    i = i + skipct;
                    grpct++;
                }

            }

            //if (btnSave.Enabled == true && flag == false)
            //{
            //    btnSave.Enabled = flag;
            //}
            //btnSave.Enabled = isowner;

            //restore sorting
            //sort by idqc
            dgqc.Sort(dgqc.Columns["idqc"], ListSortDirection.Ascending);

            //message
            string mymessage = string.Empty;
            if (unqct >= 1)
            {
                mymessage = grpct.ToString() + " group(s).\n\n\n" + (grpct + unqct).ToString() + " observation(s) grouped.";
            }
            else

            {
                mymessage = "Nothing to group!!!";
            }

            MessageBox.Show(mymessage, msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);


            skWinLib.DataGridView_Setting_After(dgqc);
            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "TS_Unique", "", "", "btnUnique_Click", lblsbar1, lblsbar2, s_tm);
        }

        private void btnexport_Click(object sender, EventArgs e)
        {

            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Export", "", "Pl.Wait,processing model check report... ", "btnexport_Click", lblsbar1, lblsbar2);
            string msg = "";
            if (dgreportid.SelectedRows.Count >= 1)
            {
                string xl_qaf01_report_List = ",";
                string xl_qaf01_report_name = "; ";
                //get report_id for export
                foreach (DataGridViewRow row in dgreportid.SelectedRows)
                {
                    if (row.Cells["rptid"].Value != null)
                    {
                        xl_qaf01_report_List = xl_qaf01_report_List + "," + row.Cells["rptid"].Value.ToString();
                        xl_qaf01_report_name = xl_qaf01_report_name + "; " + row.HeaderCell.Value.ToString() + " - " + row.Cells["rptname"].Value.ToString();
                    }
                }

                xl_qaf01_report_List = xl_qaf01_report_List.Replace(",,", "");
                xl_qaf01_report_name = xl_qaf01_report_name.Replace("; ; ", "");
                //export data result
                ArrayList export_data = Export_SK_QAF01(xl_qaf01_report_List, xl_qaf01_report_name);

                if (export_data.Count >= 1)
                {
                    msg = export_data[0].ToString();

                }

            }


            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "QAF01", "", msg, "btnexport_Click", lblsbar1, lblsbar2, s_tm);

            if (msg.Length  >= 1)
            {
                MessageBox.Show("Report created and kept in model folder.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //private void btnexport_Click(object sender, EventArgs e)
        //{
        //    DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Export", "", "Pl.Wait,processing model check report... ", "btnexport_Click" , lblsbar1, lblsbar2);

        //    ArrayList TransmittalData = new ArrayList();
        //    ArrayList Tran_Info = new ArrayList();
        //    TransmittalData.Add("EsskayQAF01");
        //    TransmittalData.Add(skTSLib.ModelPath);
        //    TransmittalData.Add(skTSLib.ModelName);

        //    //2623
        //    //0 Tran_Info.Add(esskayjn);
        //    //1 Tran_Info.Add(esskayjnName);
        //    //2 Tran_Info.Add(SK_ClientName);
        //    //3 Tran_Info.Add(SK_ReportName);
        //    //4 Tran_Info.Add(SK_ReportNameAreaSequence);
        //    //5 Tran_Info.Add(checked_by);
        //    //6 Tran_Info.Add(checked_date);
        //    //7 Tran_Info.Add(approved_by);
        //    //8 Tran_Info.Add(approved_date);
        //    //9 Tran_Info.Add(backdraft_by);
        //    //10 Tran_Info.Add(backdraft_by);
        //    //11 Tran_Info.Add(backchecked_by);
        //    //12 Tran_Info.Add(backchecked_by);

        //    //string[] projectinfo = projectmast[cmbprojid.SelectedIndex].ToString().Split(new Char[] { '|' });


        //    string SK_ClientName = client;
        //    string SK_ReportName = lblreport.Text;
        //    string SK_ReportNameAreaSequence = string.Empty;
        //    string checked_by = string.Empty;
        //    string checked_date = string.Empty;
        //    string approved_by = string.Empty;
        //    string approved_date = string.Empty;
        //    string backchecked_by = string.Empty;
        //    string backchecked_date = string.Empty;
        //    string backdraft_by = string.Empty;
        //    string backdraft_date = string.Empty;


        //    Tran_Info.Add(clientjn);
        //    Tran_Info.Add(projectname);
        //    Tran_Info.Add(SK_ClientName);
        //    Tran_Info.Add(SK_ReportName);
        //    Tran_Info.Add(SK_ReportNameAreaSequence);
        //    //get emp id and date from ncm_progress table
        //    mysql = "SELECT id_stage,user_id,user_dt FROM epmatdb.ncm_progress where project_id = " + id_project + " and report = '" + SK_ReportName + "' and status is null order by id_stage";
        //    myconn = new MySqlConnection(connString);
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    string userid = ",";
        //    while (myrdr.Read())
        //    {
        //        int report_stage = Convert.ToInt32(myrdr[0].ToString());
        //        //change 20 with other number once Jave front end completed
        //        if (report_stage == 20)
        //        {
        //            checked_by = myrdr[1].ToString();
        //            string s_sqldate = myrdr[2].ToString();
        //            if (s_sqldate.Length >= 1)
        //            {
        //                long l_sqldate = Convert.ToInt64(s_sqldate);
        //                checked_date = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //            }
        //        }
        //        else if (report_stage == 30 || report_stage == 50 || report_stage == 70 || report_stage == 90)
        //        {
        //            approved_by = myrdr[1].ToString();
        //            string s_sqldate = myrdr[2].ToString();
        //            if (s_sqldate.Length >= 1)
        //            {
        //                long l_sqldate = Convert.ToInt64(s_sqldate);
        //                approved_date = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //            }
        //        }
        //        else if (report_stage == 40)
        //        {
        //            backdraft_by = myrdr[1].ToString();
        //            string s_sqldate = myrdr[2].ToString();
        //            if (s_sqldate.Length >= 1)
        //            {
        //                long l_sqldate = Convert.ToInt64(s_sqldate);
        //                backdraft_date = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //            }
        //        }
        //        else if (report_stage == 60)
        //        {
        //            backchecked_by = myrdr[1].ToString();
        //            string s_sqldate = myrdr[2].ToString();
        //            if (s_sqldate.Length >= 1)
        //            {
        //                long l_sqldate = Convert.ToInt64(s_sqldate);
        //                backchecked_date = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //            }
        //        }
        //        userid = userid + "," + myrdr[1].ToString();

        //    }
        //    myrdr.Close();
        //    //get username based on user_id
        //    if (userid != ",")
        //    {
        //        userid = userid.Replace(",,", "");
        //        mysql = "SELECT distinct id,name,username FROM epmatdb.users where id in (" + userid + ")";
        //        myconn = new MySqlConnection(connString);
        //        if (myconn.State == ConnectionState.Closed)
        //            myconn.Open();
        //        mycmd.CommandText = mysql;
        //        mycmd.Connection = myconn;
        //        myrdr = mycmd.ExecuteReader();
        //        while (myrdr.Read())
        //        {
        //            if (checked_by == myrdr[0].ToString())
        //                checked_by = myrdr[1].ToString() + "[" + myrdr[2].ToString() + "]";
        //            if (backdraft_by == myrdr[0].ToString())
        //                backdraft_by = myrdr[1].ToString() + "[" + myrdr[2].ToString() + "]";
        //            if (backchecked_by == myrdr[0].ToString())
        //                backchecked_by = myrdr[1].ToString() + "[" + myrdr[2].ToString() + "]";
        //            if (approved_by == myrdr[0].ToString())
        //                approved_by = myrdr[1].ToString() + "[" + myrdr[2].ToString() + "]";

        //        }
        //        myrdr.Close();
        //    }



        //    Tran_Info.Add(checked_by);
        //    Tran_Info.Add(checked_date);
        //    Tran_Info.Add(approved_by);
        //    Tran_Info.Add(approved_date);
        //    Tran_Info.Add(backdraft_by);
        //    Tran_Info.Add(backdraft_date);
        //    Tran_Info.Add(backchecked_by);
        //    Tran_Info.Add(backchecked_date);

        //    TransmittalData.Add(Tran_Info);
        //    ArrayList export_data = QAF01(dgqc, TransmittalData);
        //    string msg = "";
        //    if (export_data.Count >= 1)
        //        msg = export_data[0].ToString();


        //    skTSLib.setCompletion(skApplicationName, skApplicationVersion, "QAF01", "", msg, "btnexport_Click", lblsbar1, lblsbar2, s_tm);


        //}

        private ArrayList Export_SK_QAF01(string id_report, string report_name)
        {
            string SK_ClientName = client;
            string SK_ProjectName = projectname;
            string SK_ReportName = lblreport.Text;
            string SK_ReportNameAreaSequence = string.Empty;
            string checked_by = ",";
            string checked_date = ",";
            string approved_by = ",";
            string approved_date = ",";
            string backchecked_by = ",";
            string backchecked_date = ",";
            string backdraft_by = ",";
            string backdraft_date = ",";

            ArrayList myreturndata = new ArrayList();
            try
            {
                //lblsbar1.Text = "Processing for Trasmittal....";
                string modelpath = skTSLib.ModelPath;
                string modelname = skTSLib.ModelName;
                string XLTemplate = "EsskayQAF01";
                string xlfilename = "d:\\esskay\\" + XLTemplate + ".xlsx";
                //MessageBox.Show("1");
                //Create xlapp and check whether excel is supported
                myExcel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    //MessageBox.Show("Excel is not properly installed!!", msgCaption);
                    myreturndata.Add("Excel is not properly installed!!!");
                    return myreturndata;
                }




                //check whether template available 
                if (File.Exists(xlfilename) == false)
                {
                    lblsbar1.Text = "Excel " + xlfilename + " file missing";
                    //MessageBox.Show(lblsbar1.Text, msgCaption);
                    myreturndata.Add(lblsbar1.Text);
                    return myreturndata;
                }
                myExcel.Workbook xlWorkBook;
                myExcel.Worksheet xlWorkSheet;
                Range cRange;
                object misValue = System.Reflection.Missing.Value;
                //create xl work book
                xlWorkBook = xlApp.Workbooks.Open(xlfilename, false, true);
                //MessageBox.Show("3");
                if ((skWinLib.username.ToUpper() == "SKS360") || (skWinLib.username.ToUpper() == "VIJAYAKUMAR") || (skWinLib.systemname.ToUpper() == "VIJAYAKUMAR-LAP"))
                    xlApp.Visible = true;
                else
                    xlApp.Visible = false;

                xlApp.Visible = true;

                xlApp.WindowState = myExcel.XlWindowState.xlNormal;
                xlWorkSheet = (myExcel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //xlWorkSheet.Name = "Transmittal";
                xlApp.CutCopyMode = myExcel.XlCutCopyMode.xlCopy;
                int rowid = 1;
                int colid = 2;
                int maxrowid = 1;
                int maxcolid = 9;
                Microsoft.Office.Interop.Excel.XlLineStyle myLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Microsoft.Office.Interop.Excel.XlBorderWeight myWeight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Microsoft.Office.Interop.Excel.XlHAlign myHorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                Microsoft.Office.Interop.Excel.XlVAlign myVerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;



                rowid = rowid + 8;
                int tr_start_rowid = rowid;
                int sno = 1;
                colid = 1;


                //get data from mysql
                mysql = "select idqc, guid, obser, id_cat, id_sev, mod_user_id, bd_isagree, bd_user_id, chk_user_id, backcheck_user_id, sev_user_id, mod_rmk, chk_rmk, bd_rmk, backcheck_rmk, pm_rmk, chk_dt, bd_dt, backcheck_dt, pm_dt from epmatdb.ncm where status is null and epmatdb.ncm.id_report in (" + id_report + ") order by id_report,idqc";
                if (mysql != "")
                {
                    if (myconn.State == ConnectionState.Closed)
                        myconn.Open();
                    mycmd.CommandText = mysql;
                    mycmd.Connection = myconn;
                    myrdr = mycmd.ExecuteReader();

                    ArrayList phasename = new ArrayList();
                    ArrayList chkby = new ArrayList();
                    ArrayList chkdt = new ArrayList();
                    ArrayList appby = new ArrayList();
                    ArrayList appdt = new ArrayList();
                    ArrayList backchkby = new ArrayList();
                    ArrayList backchkdt = new ArrayList();


                    bool fontstrike = false;
                    while (myrdr.Read())
                    {


                        //serial number
                        xlWorkSheet.Cells[rowid, colid] = sno.ToString();
                        string guid = skWinLib.mysql_set_quotecomma(myrdr[1].ToString());

                        //SK_ReportNameAreaSequence


                        ArrayList phaseData = skTSLib.GetPhaseName(guid);
                        string mophasename = phaseData[1].ToString();
                        if (mophasename.Length >= 1)
                        {
                            fontstrike = false;
                            if (!phasename.Contains(mophasename))
                                phasename.Add(mophasename);
                        }
                        else
                        {
                            string noPhaseObjects = "Tekla.Structures.Model.Grid,Tekla.Structures.Model.BooleanPart";
                            string motype = phaseData[0].ToString();
                            if (noPhaseObjects.Contains(motype) == false)
                                fontstrike = true;
                        }



                        string obser = skWinLib.mysql_set_quotecomma(myrdr[2].ToString());
                        string category = string.Empty;
                        string severity = string.Empty;
                        string agree_disagree = string.Empty;

                        //guid
                        xlWorkSheet.Cells[rowid, colid + 1] = "guid: " + guid;

                        //observation
                        xlWorkSheet.Cells[rowid, colid + 2] = obser;

                        //id category
                        if (myrdr[3].ToString().Length != 0)
                        {
                            if (myrdr[3].ToString() != "0") //blank for none
                            {
                                category = qmsCategory[Convert.ToInt32(myrdr[3].ToString())].ToString();
                                xlWorkSheet.Cells[rowid, colid + 3] = category;
                            }
                        }

                        //id severity
                        if (myrdr[4].ToString().Length != 0)
                        {
                            if (myrdr[4].ToString() != "0") //blank for none
                            {
                                severity = qmsSeverity[Convert.ToInt32(myrdr[4].ToString())].ToString();
                                xlWorkSheet.Cells[rowid, colid + 4] = severity;
                            }

                        }

                        //owner
                        if (myrdr[5].ToString().Length >= 1)
                        {
                            ArrayList tmp = getDBUserID(Convert.ToInt32(myrdr[5].ToString()));
                            if (tmp.Count >= 1)
                            {
                                string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                                xlWorkSheet.Cells[rowid, colid + 5] = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]"; ;
                            }

                        }

                        //agree or disagree
                        if (myrdr[6].ToString().Length != 0)
                        {

                            if (myrdr[6].ToString() == "1")
                                agree_disagree = "Agreed";
                            else
                                agree_disagree = "Not Agreed";
                            xlWorkSheet.Cells[rowid, colid + 6] = agree_disagree;
                        }

                        //corrected by [back draft by]
                        if (myrdr[7].ToString().Length >= 1)
                        {
                            ArrayList tmp = getDBUserID(Convert.ToInt32(myrdr[7].ToString()));
                            if (tmp.Count >= 1)
                            {
                                string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                                xlWorkSheet.Cells[rowid, colid + 7] = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]"; ;
                            }

                        }

                        //chk_user_id, backcheck_user_id, pm_id, mod_rmk, 
                        //Checked by
                        if (myrdr[8].ToString().Length >= 1)
                        {
                            if (!chkby.Contains(myrdr[8].ToString()))
                                chkby.Add(myrdr[8].ToString());
                        }
                        //BackCheck by
                        if (myrdr[9].ToString().Length >= 1)
                        {
                            if (!backchkby.Contains(myrdr[9].ToString()))
                                backchkby.Add(myrdr[9].ToString());
                        }
                        //Approved by
                        if (myrdr[10].ToString().Length >= 1)
                        {
                            if (!appby.Contains(myrdr[10].ToString()))
                                appby.Add(myrdr[10].ToString());
                        }

                        //checker remarks not other remarks
                        string remark = ",";
                        if (myrdr[11].ToString().Length >= 1)
                        {
                            remark = remark + "," + myrdr[11].ToString() + "[MOD]";
                        }
                        if (myrdr[12].ToString().Length >= 1)
                        {
                            remark = remark + "," + myrdr[12].ToString() + "[CHK]";
                        }
                        if (myrdr[13].ToString().Length >= 1)
                        {
                            remark = remark + "," + myrdr[13].ToString() + "[BD]";
                        }
                        if (myrdr[14].ToString().Length >= 1)
                        {
                            remark = remark + "," + myrdr[14].ToString() + "[BC]";
                        }
                        if (myrdr[15].ToString().Length >= 1)
                        {
                            remark = remark + "," + myrdr[15].ToString() + "[PM]";
                        }
                        if (remark != ",")
                        {
                            remark = remark.Replace(",,", "");
                            xlWorkSheet.Cells[rowid, colid + 8] = skWinLib.mysql_set_quotecomma(remark);
                        }

                        // chk_dt, bd_dt, backcheck_dt, pm_dt 
                        //Checked Date
                        if (myrdr[16].ToString().Length >= 1)
                        {
                            string s_sqldate = myrdr[16].ToString();
                            if (s_sqldate.Length >= 1)
                            {
                                long l_sqldate = Convert.ToInt64(s_sqldate);
                                s_sqldate = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
                                if (!chkdt.Contains(s_sqldate))
                                    chkdt.Add(s_sqldate);
                            }

                        }
                        //BackCheck Date
                        if (myrdr[18].ToString().Length >= 1)
                        {
                            string s_sqldate = myrdr[18].ToString();
                            if (s_sqldate.Length >= 1)
                            {
                                long l_sqldate = Convert.ToInt64(s_sqldate);
                                s_sqldate = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
                                if (!backchkdt.Contains(s_sqldate))
                                    backchkdt.Add(s_sqldate);
                            }
                        }
                        //Approved Date
                        if (myrdr[19].ToString().Length >= 1)
                        {
                            string s_sqldate = myrdr[19].ToString();
                            if (s_sqldate.Length >= 1)
                            {
                                long l_sqldate = Convert.ToInt64(s_sqldate);
                                s_sqldate = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
                                if (!appdt.Contains(s_sqldate))
                                    appdt.Add(s_sqldate);
                            }

                        }


                        if (fontstrike == true)
                        {
                            cRange = xlWorkSheet.Range[xlWorkSheet.Cells[rowid, 1], xlWorkSheet.Cells[rowid, maxcolid]];
                            cRange.Cells.Font.Strikethrough = fontstrike;
                            cRange.Cells.Font.Color = System.Drawing.Color.Red;
                        }

                        ////bulk formatting
                        //cRange = xlWorkSheet.Range[xlWorkSheet.Cells[rowid, 1], xlWorkSheet.Cells[rowid, maxcolid]];
                        ////MessageBox.Show("7");                       

                        //cRange.Borders.LineStyle = myLineStyle;
                        //cRange.Borders.Weight = myWeight;
                        //cRange.Cells.Font.Name = "Calibri";
                        //cRange.Cells.Font.Size = 11;
                        //cRange.Cells.Font.Bold = false;
                        //if (fontstrike == true)
                        //{
                        //    cRange.Cells.Font.Strikethrough = fontstrike;
                        //    cRange.Cells.Font.Color = System.Drawing.Color.Red;
                        //}

                        //cRange.WrapText = false;
                        //cRange.Orientation = 0;
                        //cRange.AddIndent = false;
                        //cRange.IndentLevel = 0;
                        //cRange.ShrinkToFit = false;
                        //cRange.MergeCells = false;
                        //cRange.HorizontalAlignment = myHorizontalAlignment;
                        //cRange.VerticalAlignment = myVerticalAlignment;
                        //cRange.Interior.Color = System.Drawing.Color.White;

                        rowid = rowid + 1;
                        sno++;
                    }
                    myrdr.Close();





                    //bulk formatting
                    cRange = xlWorkSheet.Range[xlWorkSheet.Cells[tr_start_rowid, 1], xlWorkSheet.Cells[rowid - 1, maxcolid]];
                    //MessageBox.Show("7");                       

                    cRange.Borders.LineStyle = myLineStyle;
                    cRange.Borders.Weight = myWeight;
                    cRange.Cells.Font.Name = "Calibri";
                    cRange.Cells.Font.Size = 11;
                    cRange.Cells.Font.Bold = false;
                    //cRange.Cells.Font.Strikethrough = fontstrike;
                    cRange.RowHeight = 25;
                    cRange.WrapText = false;
                    cRange.Orientation = 0;
                    cRange.AddIndent = false;
                    cRange.IndentLevel = 0;
                    cRange.ShrinkToFit = false;
                    cRange.MergeCells = false;
                    cRange.HorizontalAlignment = myHorizontalAlignment;
                    cRange.VerticalAlignment = myVerticalAlignment;
                    cRange.Interior.Color = System.Drawing.Color.White;
                    rowid++;




                    //Set Print Area
                    ////get the row and column details
                    var lastrowcol = (myExcel.Range)xlWorkSheet.Cells[rowid - 2, maxcolid];
                    //MessageBox.Show("10");
                    string lastrow = lastrowcol.Row.ToString();
                    string lastcolumn = skXLLib.ColumnAdress(lastrowcol.Column);
                    //MessageBox.Show("11");
                    xlWorkSheet.PageSetup.PrintArea = "A1:" + lastcolumn + lastrow;

                    //update transmittal information
                    rowid = 1;
                    colid = 2;
                    string fname = esskayjn + " - " + SK_ProjectName;
                    xlWorkSheet.Cells[rowid + 1, colid] = esskayjn + "     [" + clientjn + "]";
                    xlWorkSheet.Cells[rowid + 2, colid] = SK_ProjectName;
                    xlWorkSheet.Cells[rowid + 3, colid] = SK_ClientName;
                    xlWorkSheet.Cells[rowid + 4, colid] = report_name;

                    //get SK_ReportNameAreaSequence
                    //foreach (DataGridViewRow row in dgreportid.Rows)
                    //{
                    //    if (row.Cells["rptname"].Value != null)
                    //    {
                    //        if (row.Cells["rptname"].Value.ToString().ToUpper() == SK_ReportName.ToUpper())
                    //        {
                    //            xlWorkSheet.Cells[rowid + 4, colid] = row.HeaderCell.Value.ToString() + "-" + SK_ReportName;
                    //            fname = fname + " - " + row.HeaderCell.Value.ToString() +  " - " + SK_ReportName;
                    //            break;
                    //        }

                    //    }


                    //}


                    SK_ReportNameAreaSequence = ",";
                    //get SK_ReportNameAreaSequence
                    phasename.Sort();
                    foreach (string phname in phasename)
                    {
                        SK_ReportNameAreaSequence = SK_ReportNameAreaSequence + "," + phname;
                    }
                    if (SK_ReportNameAreaSequence != ",")
                    {
                        SK_ReportNameAreaSequence = SK_ReportNameAreaSequence.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 4, colid + 3] = SK_ReportNameAreaSequence;
                    }


                    //get all checker name
                    foreach (string checker in chkby)
                    {
                        ArrayList tmp = getDBUserID(Convert.ToInt32(checker));
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                            checked_by = checked_by + "," + empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]";
                        }

                    }
                    //get all checked date
                    foreach (string checkdate in chkdt)
                    {
                        checked_date = checked_date + "," + checkdate;

                    }

                    //get all approver name
                    foreach (string approver in appby)
                    {
                        ArrayList tmp = getDBUserID(Convert.ToInt32(approver));
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                            approved_by = approved_by + "," + empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]";
                        }

                    }
                    //get all approver date
                    foreach (string approverdt in chkdt)
                    {
                        approved_date = approved_date + "," + approverdt;

                    }

                    //get all backchecker name
                    foreach (string backchecker in backchkby)
                    {
                        ArrayList tmp = getDBUserID(Convert.ToInt32(backchecker));
                        if (tmp.Count >= 1)
                        {
                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
                            backchecked_by = backchecked_by + "," + empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]";
                        }

                    }
                    //get all backchecked date
                    foreach (string backcheckdate in backchkdt)
                    {
                        backchecked_date = backchecked_date + "," + backcheckdate;

                    }


                    if (checked_by != ",")
                    {
                        checked_by = checked_by.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 5, colid] = checked_by;
                    }
                    if (checked_date != ",")
                    {
                        checked_date = checked_date.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 6, colid] = checked_date;
                    }

                    if (approved_by != ",")
                    {
                        approved_by = approved_by.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 5, colid + 2] = approved_by;
                    }
                    if (approved_date != ",")
                    {
                        approved_date = approved_date.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 6, colid + 2] = approved_date;
                    }

                    if (backchecked_by != ",")
                    {
                        backchecked_by = backchecked_by.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 5, colid + 5] = backchecked_by;
                    }
                    if (backchecked_date != ",")
                    {
                        backchecked_date = backchecked_date.Replace(",,", "");
                        xlWorkSheet.Cells[rowid + 6, colid + 5] = backchecked_date;
                    }




                    //print page setup
                    string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    fname = fname + "; " + report_name + "; " + timestamp + ".xlsx";
                    xlWorkBook.SaveAs(skTSLib.ModelPath + "\\" + fname);
                    xlWorkBook.Close(true);
                    xlApp.Visible = false;
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    myreturndata.Add(fname);
                    return myreturndata;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Transmittal Error \n\n" + ex.Message + "\n\n" + ex.StackTrace + "\n\n" + ex.Source + "\n\n" + ex.Data.ToString());
            }
            return myreturndata;
        }
        //private ArrayList QAF01(DataGridView mydatagrid, ArrayList TransmittalData)
        //{
        //    string SK_ReportNameAreaSequence = string.Empty;
        //    string checked_by = string.Empty;
        //    string checked_date = string.Empty;
        //    string approved_by = string.Empty;
        //    string approved_date = string.Empty;
        //    string backchecked_by = string.Empty;
        //    string backchecked_date = string.Empty;
        //    string backdraft_by = string.Empty;
        //    string backdraft_date = string.Empty;

        //    ArrayList myreturndata = new ArrayList();
        //    try
        //    {
        //        //lblsbar1.Text = "Processing for Trasmittal....";
        //        string Tran_Type = TransmittalData[0].ToString();
        //        string modelpath = TransmittalData[1].ToString();
        //        string modelname = TransmittalData[2].ToString();
        //        ArrayList Tran_Info = TransmittalData[3] as ArrayList;
        //        string XLTemplate = string.Empty;
        //        //MessageBox.Show("1");
        //        //Create xlapp and check whether excel is supported
        //        myExcel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //        if (xlApp == null)
        //        {
        //            //MessageBox.Show("Excel is not properly installed!!", msgCaption);
        //            myreturndata.Add("Excel is not properly installed!!!");
        //            return myreturndata;
        //        }
        //        //MessageBox.Show("2");
        //        if (Tran_Type == "EsskayQAF01")
        //        {
        //            XLTemplate = "d:\\esskay\\" + Tran_Type + ".xlsx";
        //        }

        //        //check whether template available 
        //        if (File.Exists(XLTemplate) == false)
        //        {
        //            lblsbar1.Text = "Excel " + XLTemplate + " file missing";
        //            //MessageBox.Show(lblsbar1.Text, msgCaption);
        //            myreturndata.Add(lblsbar1.Text);
        //            return myreturndata;
        //        }
        //        myExcel.Workbook xlWorkBook;
        //        myExcel.Worksheet xlWorkSheet;
        //        Range cRange;
        //        object misValue = System.Reflection.Missing.Value;
        //        //create xl work book
        //        xlWorkBook = xlApp.Workbooks.Open(XLTemplate, false, true);
        //        //MessageBox.Show("3");
        //        if ((skWinLib.username.ToUpper() == "SKS360") || (skWinLib.username.ToUpper() == "VIJAYAKUMAR") || (skWinLib.systemname.ToUpper() == "VIJAYAKUMAR-LAP"))
        //            xlApp.Visible = true;
        //        else
        //            xlApp.Visible = false;

        //        xlApp.Visible = true;

        //        xlApp.WindowState = myExcel.XlWindowState.xlNormal;
        //        xlWorkSheet = (myExcel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        //        //xlWorkSheet.Name = "Transmittal";
        //        xlApp.CutCopyMode = myExcel.XlCutCopyMode.xlCopy;
        //        int rowid = 1;
        //        int colid = 2;
        //        int maxrowid = 1;
        //        int maxcolid = 8;
        //        Microsoft.Office.Interop.Excel.XlLineStyle myLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        //        Microsoft.Office.Interop.Excel.XlBorderWeight myWeight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
        //        Microsoft.Office.Interop.Excel.XlHAlign myHorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        //        Microsoft.Office.Interop.Excel.XlVAlign myVerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        //        if (Tran_Type == "EsskayQAF01")
        //        {
        //            //0 Tran_Info.Add(esskayjn);
        //            //1 Tran_Info.Add(esskayjnName);
        //            //2 Tran_Info.Add(SK_ClientName);
        //            //3 Tran_Info.Add(SK_ReportName);
        //            //4 Tran_Info.Add(SK_ReportAreaSequneceName);
        //            //5 Tran_Info.Add(checked_by);
        //            //6 Tran_Info.Add(checked_date);
        //            //7 Tran_Info.Add(approved_by);
        //            //8 Tran_Info.Add(approved_date);
        //            //9 Tran_Info.Add(backdraft_by);
        //            //10 Tran_Info.Add(backdraft_by);
        //            //11 Tran_Info.Add(backchecked_by);
        //            //12 Tran_Info.Add(backchecked_by);
        //            xlWorkSheet.Cells[rowid + 1, colid] = Tran_Info[0].ToString();
        //            xlWorkSheet.Cells[rowid + 2, colid] = Tran_Info[1].ToString();
        //            xlWorkSheet.Cells[rowid + 3, colid] = Tran_Info[2].ToString();
        //            xlWorkSheet.Cells[rowid + 4, colid] = Tran_Info[3].ToString();
        //            xlWorkSheet.Cells[rowid + 4, colid + 4] = Tran_Info[4].ToString();
        //            //Issue 1.1
        //            xlWorkSheet.Cells[rowid + 5, colid] = Tran_Info[5].ToString();
        //            xlWorkSheet.Cells[rowid + 6, colid] = Tran_Info[6].ToString();
        //            xlWorkSheet.Cells[rowid + 5, colid + 2] = Tran_Info[7].ToString();
        //            xlWorkSheet.Cells[rowid + 6, colid + 2] = Tran_Info[8].ToString();
        //            xlWorkSheet.Cells[rowid + 5, colid + 5] =  Tran_Info[11].ToString();
        //            xlWorkSheet.Cells[rowid + 6, colid + 5] =  Tran_Info[12].ToString();
        //            //Issue1.0
        //            //xlWorkSheet.Cells[rowid + 5, colid - 1] = "Checked By: " + Tran_Info[5].ToString();
        //            //xlWorkSheet.Cells[rowid + 6, colid - 1] = "Checked Date: " + Tran_Info[6].ToString();
        //            //xlWorkSheet.Cells[rowid + 5, colid + 1] = "Approved By: " + Tran_Info[7].ToString();
        //            //xlWorkSheet.Cells[rowid + 6, colid + 1] = "Approved Date: " + Tran_Info[8].ToString();
        //            //xlWorkSheet.Cells[rowid + 5, colid + 3] = "Back Checked By: " + Tran_Info[11].ToString();
        //            //xlWorkSheet.Cells[rowid + 6, colid + 3] = "Back Checked Date: " + Tran_Info[12].ToString();
        //            //MessageBox.Show("5");
        //            rowid = rowid + 8;
        //            int tr_start_rowid = rowid;
        //            //220511 Sl.No GUID     Observation Category    Severity Owner   Agreed / Not Agreed Remarks

        //            int sno = 1;
        //            colid = 1;

        //            //
        //            mysql = "select idqc, guid, obser, id_cat, id_sev, bd_user_id, bd_isagree, chk_rmk, chk_user_id, chk_dt, backcheck_user_id, backcheck_dt from epmatdb.ncm where status is null and epmatdb.ncm.id_ncm_progress=" + id_stage  + " order by idqc";
        //            if (mysql != "")
        //            {
        //                if (myconn.State == ConnectionState.Closed)
        //                    myconn.Open();
        //                mycmd.CommandText = mysql;
        //                mycmd.Connection = myconn;
        //                myrdr = mycmd.ExecuteReader();


        //                while (myrdr.Read())
        //                {
        //                    string category = string.Empty;
        //                    string severity = string.Empty;
        //                    string agree_disagree = string.Empty;

        //                    //serial number
        //                    xlWorkSheet.Cells[rowid, colid] = sno.ToString();

        //                    //guid
        //                    xlWorkSheet.Cells[rowid, colid + 1] = "guid: " + skWinLib.mysql_set_quotecomma(myrdr[1].ToString());

        //                    //observation
        //                    xlWorkSheet.Cells[rowid, colid + 2] = skWinLib.mysql_set_quotecomma(myrdr[2].ToString());

        //                    //id category
        //                    if (myrdr[3].ToString().Length != 0)
        //                    {
        //                        if (myrdr[3].ToString() != "0") //not none
        //                        {
        //                            category = qmsCategory[Convert.ToInt32(myrdr[3].ToString())].ToString();
        //                            xlWorkSheet.Cells[rowid, colid + 3] = category;
        //                        }

        //                    }



        //                    //id severity
        //                    if (myrdr[4].ToString().Length != 0)
        //                    {
        //                        severity = qmsSeverity[Convert.ToInt32(myrdr[4].ToString())].ToString();
        //                        xlWorkSheet.Cells[rowid, colid + 4] = severity;
        //                    }

        //                    //owner
        //                    if (myrdr[5].ToString().Length >=1)
        //                    {
        //                        ArrayList tmp = getDBUserID(Convert.ToInt32(myrdr[5].ToString()));
        //                        if (tmp.Count >= 1)
        //                        {
        //                            string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
        //                            xlWorkSheet.Cells[rowid + 5, colid] = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]"; ;
        //                        }

        //                    }

        //                    //agree or disagree
        //                    if (myrdr[6].ToString().Length != 0)
        //                    {

        //                        if (myrdr[6].ToString() == "1")
        //                            agree_disagree = "Agreed";
        //                        else
        //                            agree_disagree = "Not Agreed";
        //                        xlWorkSheet.Cells[rowid, colid + 6] = agree_disagree;
        //                    }

        //                    //checker remarks not other remarks
        //                    xlWorkSheet.Cells[rowid, colid + 7] = skWinLib.mysql_set_quotecomma(myrdr[7].ToString());
        //                    //in case checker name is missng due to split type
        //                    if (Tran_Info[5].ToString().Length == 0)
        //                    {
        //                        //in case split type use this for checker
        //                        if (myrdr[8].ToString().Length >= 1)
        //                        {
        //                            ArrayList tmp = getDBUserID(Convert.ToInt32(myrdr[8].ToString()));
        //                            if (tmp.Count >= 1)
        //                            {
        //                                string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
        //                                xlWorkSheet.Cells[rowid + 5, colid] = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]"; ;
        //                            }
        //                        }
        //                        string s_sqldate = myrdr[9].ToString();
        //                        if (s_sqldate.Length >= 1)
        //                        {
        //                            long l_sqldate = Convert.ToInt64(s_sqldate);
        //                            xlWorkSheet.Cells[rowid + 6, colid] = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //                        }
        //                    }
        //                    //in case back checker is missng due to split type
        //                    if (Tran_Info[12].ToString().Length == 0)
        //                    {
        //                        //in case split type use this for checker
        //                        if (myrdr[10].ToString().Length >= 1)
        //                        {
        //                            ArrayList tmp = getDBUserID(Convert.ToInt32(myrdr[10].ToString()));
        //                            if (tmp.Count >= 1)
        //                            {
        //                                string[] empsplit = tmp[0].ToString().Split(new Char[] { '|' });
        //                                xlWorkSheet.Cells[rowid + 5, colid + 5] = empsplit[2].ToString() + "[" + empsplit[1].ToString() + "]"; ;
        //                            }
        //                        }
        //                        string s_sqldate = myrdr[11].ToString();
        //                        if (s_sqldate.Length >= 1)
        //                        {
        //                            long l_sqldate = Convert.ToInt64(s_sqldate);
        //                            xlWorkSheet.Cells[rowid + 6, colid + 5] = DateTimeOffset.FromUnixTimeMilliseconds(l_sqldate).DateTime.ToString("dd/MMM/yy");
        //                        }
        //                    }
        //                    rowid = rowid + 1;
        //                    sno++;
        //                }
        //                myrdr.Close();
        //            }


        //            /*

        //            //string owner = Tran_Info[9].ToString();
        //            foreach (DataGridViewRow row in mydatagrid.Rows)
        //            {
        //                //check observation is not null or empty space
        //                string guid = string.Empty;
        //                string obser = string.Empty;
        //                string category = string.Empty;
        //                string severity = string.Empty;
        //                string owner = string.Empty;
        //                string agree_disagree = string.Empty;
        //                string remark = string.Empty;



        //                //Sl.No	GUID	 Observation	Category	Severity	Corrected by	Agreed / Not Agreed

        //                //serial number
        //                xlWorkSheet.Cells[rowid, colid] = sno.ToString();

        //                //GUID
        //                if (row.Cells["guid"].Value != null)
        //                    guid = "guid: " + row.Cells["guid"].Value.ToString();
        //                xlWorkSheet.Cells[rowid, colid + 1] = guid;

        //                //Observation
        //                if (row.Cells["observation"].Value != null)
        //                    obser = row.Cells["observation"].Value.ToString();
        //                xlWorkSheet.Cells[rowid, colid + 2] = obser;

        //                //Category
        //                if (row.Cells["category"].Value != null)
        //                    category = row.Cells["category"].Value.ToString();
        //                xlWorkSheet.Cells[rowid, colid + 3] = category;

        //                //Severity
        //                if (row.Cells["severity"].Value != null)
        //                    severity = row.Cells["severity"].Value.ToString();
        //                xlWorkSheet.Cells[rowid, colid + 4] = severity;

        //                //if (row.Cells["isagreed"].Value != null)
        //                //{
        //                //    agree_disagree = row.Cells["isagreed"].Value.ToString();
        //                //    if (agree_disagree == "1")
        //                //        agree_disagree = "Agreed";
        //                //}

        //                xlWorkSheet.Cells[rowid, colid + 5] = owner;

        //                //agree or disagree
        //                if (row.Cells["isagreed"].Value != null)
        //                {
        //                    agree_disagree = row.Cells["isagreed"].Value.ToString();
        //                    if (agree_disagree == "1")
        //                        agree_disagree = "Agreed";
        //                    else
        //                        agree_disagree = "Not Agreed";
        //                }
        //                else
        //                    agree_disagree = "";
        //                xlWorkSheet.Cells[rowid, colid + 6] = agree_disagree;

        //                rowid = rowid + 1;
        //                sno++;
        //            }
        //            */

        //            //formatting
        //            if (mydatagrid.Rows.Count >= 1)
        //            {
        //                cRange = xlWorkSheet.Range[xlWorkSheet.Cells[tr_start_rowid, 1], xlWorkSheet.Cells[rowid - 1, maxcolid]];
        //                //MessageBox.Show("7");                       

        //                cRange.Borders.LineStyle = myLineStyle;
        //                cRange.Borders.Weight = myWeight;
        //                cRange.Cells.Font.Name = "Calibri";
        //                cRange.Cells.Font.Size = 11;
        //                cRange.Cells.Font.Bold = false;
        //                cRange.WrapText = false;
        //                cRange.Orientation = 0;
        //                cRange.AddIndent = false;
        //                cRange.IndentLevel = 0;
        //                cRange.ShrinkToFit = false;
        //                cRange.MergeCells = false;
        //                cRange.HorizontalAlignment = myHorizontalAlignment;
        //                cRange.VerticalAlignment = myVerticalAlignment;
        //                cRange.Interior.Color = System.Drawing.Color.White;

        //            }
        //            //MessageBox.Show("8");

        //            rowid++;




        //            //Set Print Area
        //            //MessageBox.Show("9");
        //            ////get the row and column details
        //            var lastrowcol = (myExcel.Range)xlWorkSheet.Cells[rowid - 2, maxcolid];
        //            //MessageBox.Show("10");
        //            string lastrow = lastrowcol.Row.ToString();
        //            string lastcolumn = skXLLib.ColumnAdress(lastrowcol.Column);
        //            //MessageBox.Show("11");
        //            xlWorkSheet.PageSetup.PrintArea = "A1:" + lastcolumn + lastrow;

        //            //MessageBox.Show("12");

        //            //print page setup

        //            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        //            //string xlfilename = "EsskayQAF01_" + modelname + "_" + timestamp + ".xlsx";
        //            string xlfilename = esskayjn + "_" + projectname + "_" + skreportname  + "_" + timestamp + ".xlsx";
        //            //this.Text = this.Text + "      " + esskayjn + " / " + projectname + " / " + client;
        //            //string xlfilename = timestamp + ".xlsx";
        //            //MessageBox.Show("13");
        //            xlWorkBook.SaveAs(modelpath + "\\" + xlfilename);
        //            //xlWorkBook.SaveAs()
        //            xlWorkBook.Close(true);
        //            //MessageBox.Show("14");
        //            //, misValue, misValue, misValue, misValue, misValue,  Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
        //            //xlWorkBook.SaveAs(packpath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);



        //            xlApp.Visible = false;
        //            xlApp.Quit();
        //            //MessageBox.Show("15");
        //            //xlApp.WindowState = Excel.XlWindowState.xlNormal;
        //            Marshal.ReleaseComObject(xlWorkSheet);
        //            Marshal.ReleaseComObject(xlWorkBook);
        //            Marshal.ReleaseComObject(xlApp);
        //            //MessageBox.Show("16");
        //            myreturndata.Add(xlfilename);
        //            return myreturndata;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Transmittal Error \n\n" + ex.Message + "\n\n" + ex.StackTrace + "\n\n" + ex.Source + "\n\n" + ex.Data.ToString());
        //    }
        //    return myreturndata;
        //}
        //private bool Is_NewReportNameValid(string report)
        //{
        //    bool flag = skncm.IsNodeTextAvailable(tvwnew, report);
        //    if (flag == true)
        //        return false;
        //    flag = skncm.IsNodeTextAvailable(tvwchecking, report);
        //    if (flag == true)
        //        return false;
        //    flag = skncm.IsNodeTextAvailable(tvwbackdraft, report);
        //    if (flag == true)
        //        return false;
        //    flag = skncm.IsNodeTextAvailable(tvwbackcheck, report);
        //    if (flag == true)
        //        return false;
        //    //flag = skncm.IsNodeTextAvailable(tvwheld, report);
        //    //if (flag == true)
        //    //    return false;
        //    //flag = skncm.IsNodeTextAvailable(tvwcancelled, report);
        //    //if (flag == true)
        //    //    return false;
        //    flag = skncm.IsNodeTextAvailable(tvwcompleted, report);
        //    if (flag == true)
        //        return false;

        //    //check whether report name already exists in same project
        //    flag = IsReportExists(report, id_project);
        //    if (flag == true)
        //        return false;



        //    return true;
        //}



        //function to create new report number after validation of inputs
        private int Create_NewReportNumber(string report, int project, int newReportNumber, int id_emp, long timestamp)
        {


            lblsbar2.Text = "Create_NewReportNumber,new_ReportNumber:" + newReportNumber.ToString();
            int myreturn = -1;

            //Check and create Connection
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();



            //insert fresh record into my sql
            mycmd.CommandText = "INSERT INTO epmatdb.ncm_report(project_id,number,name,user_id,date,sys_name) VALUES(@project_id, @number, @name, @user_id, @date, @sys_name)";
            mycmd.Parameters.Clear();

            mycmd.Parameters.Add("@project_id", MySqlDbType.Int32);
            mycmd.Parameters.Add("@number", MySqlDbType.Int32);
            mycmd.Parameters.Add("@name", MySqlDbType.LongText);
            mycmd.Parameters.Add("@user_id", MySqlDbType.Int32);
            mycmd.Parameters.Add("@date", MySqlDbType.Int64);
            mycmd.Parameters.Add("@sys_name", MySqlDbType.VarChar);

            mycmd.Parameters["@project_id"].Value = id_project;
            mycmd.Parameters["@number"].Value = newReportNumber;
            mycmd.Parameters["@name"].Value = report;
            mycmd.Parameters["@user_id"].Value = id_emp;


            mycmd.Parameters["@date"].Value = timestamp;
            mycmd.Parameters["@sys_name"].Value = skWinLib.systemname;




            //insert new process 
            return myreturn = mycmd.ExecuteNonQuery();
        }

        private int Update_ReportNumber(string report, int id_report, int id_emp, long timestamp)
        {
            lblsbar2.Text = "Update_ReportNumber";
            int myreturn = -1;

            //Check and create Connection
            //update with new report name
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = "UPDATE `ncm_report` set name = '" + report + "',  user_id = " + id_emp + ",  date = " + timestamp + " where id_report = " + id_report;
            mycmd.Connection = myconn;
            //update record
            return myreturn = mycmd.ExecuteNonQuery();
        }

        private int GetLastReportNumber(int project)
        {
            mysql = "SELECT max(number) FROM epmatdb.ncm_report where project_id = " + project;
            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {
                //always number will be integer hence no integercheck not done
                string s_report_number = myrdr[0].ToString();
                if (s_report_number.Trim().Length >= 1)
                {
                    int report_number = Convert.ToInt32(s_report_number);
                    myrdr.Close();
                    return report_number;
                }
            }
            myrdr.Close();
            return -1;
        }

        private int IsReportExists(string report, int project)
        {
            mysql = "select id_report from epmatdb.ncm_report where project_id = " + project + " and name = '" + report + "'";
            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {
                //always number will be integer hence no integercheck not done
                string s_id_report = myrdr[0].ToString();
                if (s_id_report.Trim().Length >= 1)
                {
                    int id_report = Convert.ToInt32(s_id_report);
                    myrdr.Close();
                    return id_report;
                }
            }
            myrdr.Close();
            return -1;
        }



        //private bool IsReportExists(string report, int project)
        //{
        //    ArrayList myreturn = new ArrayList();
        //    mysql = "SELECT report  FROM epmatdb.ncm_progress where project_id = " + project + " and report = '" + report + "'";
        //    myconn = new MySqlConnection(connString);
        //    if (myconn.State == ConnectionState.Closed)
        //        myconn.Open();
        //    mycmd.CommandText = mysql;
        //    mycmd.Connection = myconn;
        //    myrdr = mycmd.ExecuteReader();
        //    while (myrdr.Read())
        //    {
        //        myrdr.Close();
        //        return true;
        //    }
        //    myrdr.Close();



        //    return false;
        //}

        public static class Prompt
        {
            public static string ShowDialog(string caption)
            {
                Form prompt = new Form()
                {
                    Width = 500,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedSingle,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };

                RadioButton newrep = new RadioButton() { Left = 30, Top = 10, AutoSize = true, Checked = true, Text = "New Project" };
                RadioButton earrep = new RadioButton() { Left = 320, Top = 10, AutoSize = true, Text = "Earlier Project" };
                Label textLabel = new Label() { Left = 30, Top = 40, AutoSize = true, Text = "Start Report Number" };
                //this.lblqmsdblist.AutoSize = true;
                TextBox textBox = new TextBox() { Left = 320, Top = 40, Width = 60, Text = "1" };
                Button confirmation = new Button() { Text = "OK", Left = 200, Width = 100, Top = 70, DialogResult = DialogResult.OK };
                newrep.Click += (sender, e) => { textBox.Text = "1"; };
                earrep.Click += (sender, e) => { textBox.Text = ""; };
                confirmation.Click += (sender, e) => { prompt.Close(); };

                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(newrep);
                prompt.Controls.Add(earrep);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }

        private void btn1_Click(object sender, EventArgs e)
        {
            //isowner = false;
            string option = btn1.Text.Trim().ToUpper();
            if (option == "+")
            {
                string report = cmbqmsdblist.Text.ToUpper().Trim();
                if (report.Length >= 1)
                {

                    //check whether same report number already exists
                    if (IsReportExists(report, id_project) == -1)
                    {
                        //get the last report number for that project
                        int reportnumber = GetLastReportNumber(id_project);
                        if (reportnumber == -1)
                        {

                            //this is first time for this project hence get confirmation from user whether project is new or earlier for report start number
                            bool topmostflag = this.TopMost;
                            this.TopMost = false;
                        ReportNumberConfirmation:
                            string report_number = Prompt.ShowDialog(esskayjn + " " + projectname + " Report Number Confirmation:");
                            if (report_number.Length == 0)
                            {
                                MessageBox.Show("Report number cannot be blank. Please check your input and try again.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                goto ReportNumberConfirmation;
                            }
                            else
                            {
                                if (skWinLib.IsNumeric(report_number) == false)
                                {
                                    MessageBox.Show(report_number + " in not a valid number. Please check your input and try again.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    goto ReportNumberConfirmation;
                                }

                            }
                            this.TopMost = topmostflag;
                            reportnumber = Convert.ToInt32(report_number) - 1;
                        }

                        //add new report number for that project with +1 from the reportnumber
                        id_project_report = reportnumber + 1;

                        DateTime chk_dt = DateTime.Now;
                        DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
                        long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

                        int chkflag = Create_NewReportNumber(report, id_project, id_project_report, id_emp, timestamp);
                        if (chkflag == 1)
                        {
                            //get id_report for the newly added report number for that project

                            int SK_id_report = IsReportExists(report, id_project);
                            //add new report number to treeview
                            TreeNode tvwnewnode = tvwnew.Nodes.Add(report);
                            //process is modelling since new entry
                            SK_id_process = (int)SKProcess.Modelling;
                            lblreport.Text = report;
                            lblstage.Text = "New";
                            //Update_Stage_ReportName(report, "New", id_project_report);
                            //id_stage = id_project_report;
                            tvwnewnode.Tag = SK_id_report.ToString() + "|" + SK_id_process;


                            isowner = true;
                        }
                        else
                        {
                            MessageBox.Show(report + " cannot be added.\n\nCritical Error [" + id_project_report + "].\n\nPlease contact suppor team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {

                        MessageBox.Show(report + " already exists. Please try with other name.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {

                    MessageBox.Show("Report name cannot be blank. Please try with other name.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            else if (option == "O")
            {
                //for entry into data grid cell "Observation"
                if (cmbqmsdblist.Text.Trim().Length >= 1)
                {
                    //tpgqc.Tag = cmbqmsdblist.Text;;

                    UpdateDataGridSelectedRow_String(dgqc, "Observation", cmbqmsdblist.Text, chkreplace.Checked);
                }
            }
            else if (option == "R")
            {
                //for entry into data grid cell "Remark"
                if (cmbqmsdblist.Text.Trim().Length >= 1)
                {
                    //tpgqc.Tag = cmbqmsdblist.Text;;
                    UpdateDataGridSelectedRow_String(dgqc, "Remark", cmbqmsdblist.Text, chkreplace.Checked);
                }
            }
        }

        //private void btn1_Click(object sender, EventArgs e)
        //{
        //    //isowner = false;
        //    string option = btn1.Text.Trim().ToUpper();
        //    if (option == "+")
        //    {
        //        string report = cmbqmsdblist.Text.ToUpper().Trim();
        //        if (report.Length >= 1)
        //        {
        //            if (Is_NewReportNameValid(report) == true)
        //            {
        //                TreeNode tvwnewnode = tvwnew.Nodes.Add(report);
        //                //tvwnew.ForeColor = System.Drawing.Color.Black;
        //                //tvwnewnode.ForeColor = System.Drawing.Color.Black;
        //                id_process = (int)SKProcess.Modelling_PM_Approved;
        //                Update_Stage_ReportName(report, "New", id_process);
        //                id_stage = -1;
        //                isowner = true;
        //                ////tpgqc.Tag = cmbqmsdblist.Text;;
        //                //tpgdash.Tag = cmbqmsdblist.Text;


        //            }
        //            else
        //            {

        //                MessageBox.Show(report + " already exists. Please try with other name.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //            }
        //        }
        //        else
        //        {

        //            MessageBox.Show("Report name cannot be blank. Please try with other name.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }


        //    }
        //    else if (option == "O")
        //    {
        //        //for entry into data grid cell "Observation"
        //        if (cmbqmsdblist.Text.Trim().Length >= 1)
        //        {
        //            //tpgqc.Tag = cmbqmsdblist.Text;;

        //            UpdateDataGridSelectedRow_String(dgqc, "Observation", cmbqmsdblist.Text, chkreplace.Checked);
        //        }
        //    }
        //    else if (option == "R")
        //    {
        //        //for entry into data grid cell "Remark"
        //        if (cmbqmsdblist.Text.Trim().Length >= 1)
        //        {
        //            //tpgqc.Tag = cmbqmsdblist.Text;;
        //            UpdateDataGridSelectedRow_String(dgqc, "Remark", cmbqmsdblist.Text, chkreplace.Checked);
        //        }
        //    }
        //}
        private ArrayList get_userid_tbl_ncm_process(int project, string report)
        {
            //fetch  user information and store in empmast;
            ArrayList myreturn = new ArrayList();
            mysql = "SELECT distinct user_id,id_ncm_progress  FROM epmatdb.ncm_progress where project_id = " + project + " and report = '" + report + "'";
            myconn = new MySqlConnection(connString);
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            myrdr = mycmd.ExecuteReader();
            while (myrdr.Read())
            {
                myreturn.Add(myrdr[0].ToString() + "|" + myrdr[1].ToString());
            }
            myrdr.Close();
            return myreturn;
        }

        private ArrayList Report_Rename(string report)
        {
            ArrayList myreturn = new ArrayList();
            string mymsg = string.Empty;
            MessageBoxIcon mymsgicon = MessageBoxIcon.Information;
            if (report.Length >= 1)
            {
                //check report is already stored in database if not can be modified
                TreeNode selnode = tvwncm.SelectedNode;

                string option = btn3.Text.Trim().ToUpper();
                string old_report = selnode.Text.ToUpper();
                if (option == "REN")
                {
                    if (IsReportExists(report, id_project) == -1)
                    {
                        if (MessageBox.Show(old_report + " will be renamed as " + report + ".\n\n\nCan you re-confirm this change ...?", msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            DateTime chk_dt = DateTime.Now;
                            DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
                            long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

                            mysql = "UPDATE `ncm_report` SET name = '" + report + "', user_id = " + id_emp + ", date = " + timestamp + ", sys_name = '" + skWinLib.systemname + "' WHERE name = '" + old_report + "' and id_report = '" + SK_id_report + "'";
                            mycmd.CommandText = mysql;
                            mycmd.Connection = myconn;


                            if (myconn.State == ConnectionState.Closed)
                                myconn.Open();

                            //update record
                            if (mycmd.ExecuteNonQuery() == 1)
                            {
                                mymsg = old_report + " changed to " + report;
                                UpdateUserInterface_Fromlastsave();
                            }
                            else
                            {
                                mymsgicon = MessageBoxIcon.Error;
                                mymsg = "Error " + old_report + " cannot be renamed.";
                            }
                        }
                    }
                    else
                    {
                        mymsgicon = MessageBoxIcon.Error;
                        mymsg = report + " already exist hence cannot be renamed.";
                    }

                }
            }
            else
            {
                mymsg = "Report can't be null or blank!!!";
                mymsgicon = MessageBoxIcon.Exclamation;
            }

            myreturn.Add(mymsg);
            myreturn.Add(mymsgicon);
            return myreturn;
        }
        /*
//tpgdash.Tag = cmbqmsdblist.Text;
                    //check if the process is new
                    if (lblstage.Text == "New")
                    {

                        if (selnode != null)
                        {
                            if (selnode.Text.ToUpper() != tpgqc.Text)
                            {
                                mymsg = tpgqc.Text + " mismatch(s)  " + old_report + ".\n\nCheck your selection!!!";
                                mymsgicon = MessageBoxIcon.Exclamation;
                            }
                            else
                            {

                                //if (Is_NewReportNameValid(report) == true)
                                int id_report = IsReportExists(report, id_project);
                                if (id_report == -1)
                                {
                                    //update report name
                                    DateTime chk_dt = DateTime.Now;
                                    DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
                                    long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();
                                    Update_ReportNumber(report, SK_id_report , id_emp, timestamp);

                                    selnode.Text = report;
                                    mymsg = "Renamed";

                                    //update UI
                                    Update_Dashboard_ReportNumber();

                                    //tab page name changed
                                    tpgqc.Text = report;
                                    mymsgicon = MessageBoxIcon.Information;

                                }
                                else
                                {
                                    mymsg = report + " already exists. Please try with other name.";
                                    mymsgicon = MessageBoxIcon.Stop;
                                }
                            }
                        }
                        else
                        {
                            mymsg = report + " is not the currenlty selected item.\n\nEnsure your select and RENAME";
                            mymsgicon = MessageBoxIcon.Error;
                        }
                    }
                    else
                    {
                        //not new process
                        //check earlier report already stored in database, if stored then owner should be the current employee and further process should not have begin or not approved
                        ArrayList myolddata = get_userid_tbl_ncm_process(id_project, old_report);
                        //check whether no other employee are the owner for this report
                        foreach (string chkempid in myolddata)
                        {
                            string[] splt_chkempid = chkempid.ToString().Split(new Char[] { '|' });
                            if (splt_chkempid[0] != id_emp.ToString())
                            {
                                mymsg = old_report + " already approved/started by [ " + myolddata.Count.ToString() + " ] employee(s) hence can't be renamed.\n\n\nAborting Now!!!";
                                mymsgicon = MessageBoxIcon.Exclamation;
                                myreturn.Add(mymsg);
                                myreturn.Add(mymsgicon);
                                return myreturn;
                            }
                        }

                        //check 2, whether new report already exists if not can be allowed to rename now

                        if (MessageBox.Show(old_report + " will be renamed as " + report + ".\n\n\nCan you re-confirm this change ...?", msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            ArrayList mydata = get_userid_tbl_ncm_process(id_project, report);

                            if (mydata.Count == 0 || mydata.Count == -1 || mydata == null)
                            {
                                ////Ask double confirmation
                                //mydata = get_userid_tbl_ncm_process(id_project, report);
                                if (mydata.Count == 0 || mydata.Count == -1 || mydata == null)
                                {
                                    //fetch  user information and store in empmast;
                                    mysql = "SELECT id_ncm_progress FROM epmatdb.ncm_progress where project_id = " + id_project + " and user_id = " + id_emp + " and report = '" + old_report + "'";
                                    if (myconn.State == ConnectionState.Closed)
                                        myconn.Open();
                                    mycmd.CommandText = mysql;
                                    mycmd.Connection = myconn;
                                    myrdr = mycmd.ExecuteReader();
                                    string update_cond = "(";
                                    int uct = 0;
                                    while (myrdr.Read())
                                    {
                                        update_cond = update_cond + "," + myrdr[0].ToString();
                                        uct++;
                                    }
                                    myrdr.Close();

                                    update_cond = update_cond.Replace("(,", "(");
                                    update_cond = " where user_id = " + id_emp + " and id_ncm_progress in " + update_cond + ")";


                                    //update with new report name
                                    if (myconn.State == ConnectionState.Closed)
                                        myconn.Open();
                                    mycmd.CommandText = "UPDATE `ncm_progress` set report = '" + report + "' " + update_cond + " and report ='" + old_report + "'";
                                    mycmd.Connection = myconn;
                                    //update record
                                    int flag = mycmd.ExecuteNonQuery();
                                    if (flag == -1)
                                    {
                                        mymsg = "Contact support team since renaming of " + old_report + " to " + report + " can't be done.";
                                        mymsgicon = MessageBoxIcon.Error;
                                    }
                                    else
                                    {
                                        mymsg = old_report + " changed to " + report + "\n\n" + flag.ToString() + " records(s) updated";
                                        selnode.Text = report;
                                        //tab page name changed
                                        tpgqc.Text = report;
                                    }
                                }
                                else
                                {
                                    mymsg = "Contract support team since " + report + " exists now and can't be renamed.";
                                    mymsgicon = MessageBoxIcon.Stop;
                                }

                            }
                            else
                            {
                                mymsg = report + " already exists!!!";
                                mymsgicon = MessageBoxIcon.Exclamation;
                            }
                        }

                    }
        
        */
        private void btn3_Click(object sender, EventArgs e)
        {
            string task = "Rename";
            if (btn3.Text == "NC")
                task = sk_nocomment;
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, task, "", "Pl.Wait,processing report... ", "btn" + task + "_Click", lblsbar1, lblsbar2);
            string mymsg = string.Empty;
            MessageBoxIcon mymsgicon = MessageBoxIcon.Information;
            if (task == "Rename")
            {
                //get the report name and check null entry or valid entry
                string report = cmbqmsdblist.Text.ToUpper().Trim();
                ArrayList myreturndata = Report_Rename(report);
                if (myreturndata.Count >= 1)
                {
                    mymsg = (string)myreturndata[0];
                    mymsgicon = (MessageBoxIcon)myreturndata[1];
                    if (mymsg != string.Empty)
                        MessageBox.Show(mymsg, msgCaption, MessageBoxButtons.OK, mymsgicon);
                }
            }
            else if (task == sk_nocomment)
            {
                UpdateDataGridSelectedRow_String(dgqc, "Observation", sk_nocomment, chkreplace.Checked);
            }

            skWinLib.setCompletion(skApplicationName, skApplicationVersion, task, "", "", "btn" + task + "_Click", lblsbar1, lblsbar2, s_tm);
        }


        private void btn2_Click(object sender, EventArgs e)
        {
            string option = btn2.Text.Trim().ToUpper();
            //check report is already stored in database if not can be modified
            TreeNode selnode = tvwncm.SelectedNode;
            string mymsg = string.Empty;
            string report = cmbqmsdblist.Text.ToUpper().Trim();
            MessageBoxIcon mymsgicon = MessageBoxIcon.Information;
            if (option == "-")
            {
                //tpgdash.Tag = cmbqmsdblist.Text;
                if (selnode != null)
                {
                    //check if process is new
                    if (lblstage.Text == "New")
                    {

                        if (selnode.Text.ToUpper() != report)
                        {
                            mymsg = report + " is not the currenlty selected item.\n\nEnsure your new " + report + " is active and delete";
                            mymsgicon = MessageBoxIcon.Exclamation;
                        }
                        else
                        {
                            selnode.Remove();
                            //tab page removed
                            TabControl1.TabPages.Remove(tpgqc);
                            mymsg = "Report deleted.";

                            mymsgicon = MessageBoxIcon.Information;

                        }
                    }
                    else
                    {
                        //not a new process 

                        ArrayList mydata = get_userid_tbl_ncm_process(id_project, report);
                        if (mydata.Count >= 1)
                        {
                            MessageBox.Show(report + " already exists[" + mydata.Count.ToString() + " nos] hence deletion is not allowed.\n\nExisting Now!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        if (selnode.Text.ToUpper() != report)
                        {
                            mymsg = report + " is not the currenlty selected item.\n\nEnsure your " + report + " is active and delete";
                            mymsgicon = MessageBoxIcon.Exclamation;
                        }
                        else
                        {
                            selnode.Remove();
                            //tab page removed
                            TabControl1.TabPages.Remove(tpgqc);
                            mymsg = "Report deleted.";
                            mymsgicon = MessageBoxIcon.Information;
                        }
                    }
                }
                else
                {
                    mymsg = "Invalid report selection (current report name not active).";
                    mymsgicon = MessageBoxIcon.Stop;
                }
                MessageBox.Show(mymsg, msgCaption, MessageBoxButtons.OK, mymsgicon);
            }
            //for entry into data grid cell "Remark"
            else if (option == "R")
            {
                //tpgqc.Tag = cmbqmsdblist.Text;;
                UpdateDataGridSelectedRow_String(dgqc, "Remark", cmbqmsdblist.Text, chkreplace.Checked);

            }
        }

        private void TabControl1_Enter(object sender, EventArgs e)
        {

        }

        private void tpgDash_Enter(object sender, EventArgs e)
        {
            //if (tpgqc.Text.ToUpper().Trim() != skreportname.ToUpper().Trim() )
            //{
            //    dgqc.Rows.Clear();
            //}
            //if active tab page is ncm set the user inferface accordingly
            //SetUserInteface_Tabpage("DashBoard");

            //SetUserInteface_DashBoard();

            //set enable or disable of button based on row count in datagrid qc
            //setpreview();
            //EnableDisableControl(false);
        }

        //private void TabControl1_Click(object sender, EventArgs e)
        //{
        //    //if active tab page is dashboard set the user inferface accordingly
        //     if (TabControl1.SelectedIndex == 0) 
        //        SetUserInteface_Tabpage("DASHBOARD");
        //}

        //private void btnsplitsave_Click(object sender, EventArgs e)
        //{
        //    DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Partial Save", "", "Pl.Wait,partial save ", "btnsplitsave_Click", lblsbar1, lblsbar2);
        //    //tpgqc.Tag = cmbqmsdblist.Text;;
        //    string msg = "";
        //    int updtct = 0;
        //    if (cmbqmsdblist.Text.ToUpper().Trim() == tpgqc.Text.ToUpper().Trim())
        //    {
        //        //all row's selected hence suggest user to rename instead of this
        //        MessageBox.Show("New report name required.\n\nFor split provide other/new name.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    else
        //    {
        //        //check user have selected few rows for partial save
        //        int ct = dgqc.SelectedRows.Count;
        //        if (ct == dgqc.RowCount)
        //        {
        //            //all row's selected hence suggest user to rename instead of this
        //            MessageBox.Show("Instead of selecting all row and split/partial save, try renameing(REN) the report in dashboard.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //        else
        //        {
        //            if (ct >= 1)
        //            {
        //                //check whether new report name exist if exist then ensure current user should be the only owner
        //                string report = cmbqmsdblist.Text.ToUpper().Trim();
        //                ArrayList myolddata = get_userid_tbl_ncm_process(id_project, report);
        //                //check whether no other employee are the owner for this report
        //                //bool isowner = true;
        //                foreach (string chkempid in myolddata)
        //                {
        //                    if (chkempid != id_emp.ToString())
        //                    {
        //                        //isowner = false;
        //                        MessageBox.Show(report + " already in use by [ " + myolddata.Count.ToString() + " ] employee(s) hence provide different report name for split / partial save.\n\n\nAborting Now!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
        //                        return;
        //                    }
        //                }
        //                if (MessageBox.Show(ct.ToString() + " record(s) will be split / partial saved from earlier " + tpgqc.Text + " to new " + report + ".\n\n\nPlease re-confirm for split / partial save ...?", msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        //                {
        //                    //check whether selected row are already saved if not inform user to save first
        //                    foreach (DataGridViewRow row in dgqc.SelectedRows)
        //                    {

        //                        if (row.Cells["idqc"].Value != null)
        //                        {
        //                            string tmp = row.Cells["idqc"].Value.ToString();
        //                            if (skWinLib.IsNumeric(tmp) == false)
        //                            {
        //                                MessageBox.Show("All selected row's should have been saved but " + row.Index.ToString() + " not saved earlier.\nSave all and split/partial later if required.\n\nAborting Now!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                                return;
        //                            }
        //                        }

        //                    }
        //                    //all check passed and confirm selected row's are already saved earlier and now user prefer to split/partial save
        //                    //create new stage for the report
        //                    lblreport.Text = cmbqmsdblist.Text;
        //                    int ear_stage = id_stage;
        //                    //if (isowner == true)
        //                    //    id_stage = 25;
        //                    //else
        //                    //    id_stage = CreateNewReport(id_process);

        //                    id_stage = CreateNewReport(id_process);

        //                    DateTime chk_dt = DateTime.Now;
        //                    DateTimeOffset dtoff_checkDate = new DateTimeOffset(chk_dt);
        //                    long timestamp = dtoff_checkDate.ToUnixTimeMilliseconds();

        //                    //ArrayList removerowid = new ArrayList();
        //                    foreach (DataGridViewRow row in dgqc.SelectedRows)
        //                    {
        //                        if (row.Cells["idqc"].Value != null)
        //                        {
        //                            int i_idqc = Convert.ToInt32(row.Cells["idqc"].Value.ToString());
        //                            mysql = "UPDATE `ncm` SET id_ncm_progress = " + id_stage + " WHERE idqc=" + i_idqc;
        //                            mycmd.CommandText = mysql;
        //                        }
        //                        else
        //                        {

        //                            //first time entry by user hence insert
        //                            mycmd.CommandText = "INSERT INTO epmatdb.ncm(guid,guid_type,obser,chk_rmk,id_ncm_progress,chk_user_id,chk_dt) VALUES(@guid, @guid_type, @obser, @chk_rmk, @id_ncm_progress, @chk_user_id, @chk_dt)";
        //                            //if (flag == false)
        //                            {
        //                                mycmd.Parameters.Clear();
        //                                mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
        //                                mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
        //                                mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
        //                                mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
        //                                mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);
        //                                mycmd.Parameters.Add("@chk_user_id", MySqlDbType.Int32);
        //                                mycmd.Parameters.Add("@chk_dt", MySqlDbType.Int64);
        //                            }

        //                            //mycmd.CommandText = "INSERT INTO epmatdb.ncm(project_id,guid,guid_type,obser,chk_rmk,chk_user_id,chk_sys,id_ncm_progress,entry_date) VALUES(@project_id, @guid, @guid_type, @obser, @chk_rmk, @chk_user_id, @chk_sys, @id_ncm_progress, @entry_date)";
        //                            ////if (flag == false)
        //                            //{
        //                            //    mycmd.Parameters.Clear();
        //                            //    mycmd.Parameters.Add("@project_id", MySqlDbType.Int64);
        //                            //    mycmd.Parameters.Add("@guid", MySqlDbType.VarChar);
        //                            //    mycmd.Parameters.Add("@guid_type", MySqlDbType.VarChar);
        //                            //    mycmd.Parameters.Add("@obser", MySqlDbType.VarChar);
        //                            //    mycmd.Parameters.Add("@chk_rmk", MySqlDbType.LongText);
        //                            //    mycmd.Parameters.Add("@chk_user_id", MySqlDbType.LongText);
        //                            //    mycmd.Parameters.Add("@chk_sys", MySqlDbType.VarChar);
        //                            //    mycmd.Parameters.Add("@id_ncm_progress", MySqlDbType.Int32);
        //                            //    mycmd.Parameters.Add("@entry_date", MySqlDbType.Int64);

        //                            //}

        //                            string obser = string.Empty;

        //                            if (row.Cells["observation"].Value != null)
        //                                 obser = skWinLib.mysql_replace_quotecomma(row.Cells["observation"].Value.ToString());

        //                            string guid = string.Empty;
        //                            //get variable value from data grid
        //                            if (row.Cells["guid"].Value != null)
        //                                guid = row.Cells["guid"].Value.ToString();

        //                            string guid_type = string.Empty;
        //                            //get variable value from data grid
        //                            if (row.Cells["TeklaSKtype"].Value != null)
        //                                guid_type = row.Cells["TeklaSKtype"].Value.ToString();


        //                            string remark = string.Empty;
        //                            //get variable value from data grid
        //                            if (row.Cells["remark"].Value != null)
        //                                remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());


        //                            //mycmd.Parameters["@project_id"].Value = id_project;
        //                            mycmd.Parameters["@guid"].Value = guid;
        //                            mycmd.Parameters["@guid_type"].Value = guid_type;
        //                            mycmd.Parameters["@obser"].Value = obser;
        //                            mycmd.Parameters["@chk_rmk"].Value = remark;
        //                            mycmd.Parameters["@chk_user_id"].Value = id_emp;
        //                            //mycmd.Parameters["@chk_sys"].Value = skWinLib.systemname;
        //                            mycmd.Parameters["@id_ncm_progress"].Value = id_stage;
        //                            mycmd.Parameters["@chk_dt"].Value = timestamp;
        //                        }


        //                        mycmd.Connection = myconn;
        //                        if (myconn.State == ConnectionState.Closed)
        //                            myconn.Open();

        //                        //update record
        //                        if (mycmd.ExecuteNonQuery() == 1)
        //                            updtct++;
        //                        //removerowid.Add(i_idqc);
        //                        dgqc.Rows.Remove(row);
        //                    }
        //                    ////remove selected row                        
        //                    //foreach (DataGridViewRow delrow in dgqc.SelectedRows)                        
        //                    //{
        //                    //    dgqc.Rows.Remove(delrow);
        //                    //}
        //                    //add treeview with new data
        //                    //TreeNode newchildnode = null;
        //                    ////newchildnode = tvwncm.SelectedNode;
        //                    //if (id_process < 20)
        //                    //{
        //                    //    //Checking process
        //                    //    newchildnode = tvwchecking.Nodes.Add(cmbqmsdblist.Text);

        //                    //}
        //                    //else if ((id_process >= 20) && (id_process < 50))
        //                    //{
        //                    //    //back drafting process
        //                    //    newchildnode = tvwbackdraft.Nodes.Add(cmbqmsdblist.Text);
        //                    //}
        //                    //else if ((id_process >= 50) && (id_process < 60))
        //                    //{
        //                    //    //back check process
        //                    //    newchildnode = tvwbackcheck.Nodes.Add(cmbqmsdblist.Text);
        //                    //}
        //                    //else if ((id_process >= 60))
        //                    //{
        //                    //    //Column Name, IsVisible, IsReadOnly
        //                    //    newchildnode = tvwcompleted.Nodes.Add(cmbqmsdblist.Text);
        //                    //}
        //                    ////newchildnode = tvwchecking.Nodes.Add(cmbqmsdblist.Text);
        //                    //newchildnode.Tag = id_process.ToString() +  "|" + id_stage;
        //                    //newchildnode.ForeColor =  System.Drawing.Color.Blue;
        //                    ////set old report name
        //                    //lblreport.Text = tpgqc.Text;
        //                    //id_stage = ear_stage;
        //                    ////Fetch_Model_Data();

        //                    //skWinLib.updaterowheader(dgqc);

        //                    //refresh
        //                    Update_DashBoard_TreeView();

        //                    //update row header
        //                    skWinLib.updaterowheader(dgqc);

        //                    MessageBox.Show(updtct.ToString() + " record(s) split/partial saved to " + cmbqmsdblist.Text, msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("No row(s) selected for partial save, check your selection!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //            }
        //        }
        //    }

        //    skWinLib.setCompletion(skApplicationName, skApplicationVersion, "Partial Save", "", msg, "btnsplitsave_Click", lblsbar1, lblsbar2, s_tm);

        //}

        private void btntspreview2_Click(object sender, EventArgs e)
        {
            //in model show only objects in datagridview [TOGGLE]
        }


        //private void btntsmodelselect_Click(object sender, EventArgs e)
        //{
        //    if (sk_ts_display == false)
        //    {
        //        //Tekla
        //        List<Identifier> myListIdentifier = new List<Identifier>();
        //        List<Identifier> myNocommentListIdentifier = new List<Identifier>();
        //        //tpgqc.Tag = cmbqmsdblist.Text;;

        //        try
        //        {
        //            //int rowindex = e.RowIndex;
        //            //DataGridViewRow row = mydatagrid.Rows[rowindex];
        //            ArrayList ObjectsToSelect = new ArrayList();


        //            int i = 0;
        //            foreach (DataGridViewRow row in dgqc.Rows)
        //            {
        //                Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier();
        //                string myguid = row.Cells["GUID"].Value.ToString();
        //                string myobser = row.Cells["observation"].Value.ToString();
        //                myguid = myguid.ToUpper().Replace("GUID:", "").Replace("GUID", "").Trim();
        //                myIdentifier = new Tekla.Structures.Identifier(myguid);
        //                bool flag = myIdentifier.IsValid();
        //                if (flag)
        //                {
        //                    Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
        //                    if (ModelSideObject.Identifier.IsValid())
        //                    {
        //                        if (!ObjectsToSelect.Contains(ModelSideObject))
        //                            ObjectsToSelect.Add(ModelSideObject);

        //                        if (myobser == sk_nocomment || myobser == "\nNo Comment")
        //                        {
        //                            if (!myNocommentListIdentifier.Contains(myIdentifier))
        //                                myNocommentListIdentifier.Add(myIdentifier);
        //                        }
        //                        else if (!myListIdentifier.Contains(myIdentifier))
        //                            myListIdentifier.Add(myIdentifier);
        //                    }
        //                }
        //                i++;
        //            }

        //            var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
        //            MOS.Select(ObjectsToSelect);

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Representation Error: " + ex.Message);
        //        }
        //        ModelObjectVisualization.ClearAllTemporaryStates();
        //        TSMUI.Color tsgrey = new TSM.UI.Color(0.0, 0.0, 1.0);
        //        ModelObjectVisualization.SetTemporaryStateForAll(tsgrey);

        //        TSMUI.Color tsred = new TSM.UI.Color(1.0, 0.0, 0.0);
        //        ModelObjectVisualization.SetTemporaryState(myListIdentifier, tsred);

        //        TSMUI.Color tsgreen = new TSM.UI.Color(0.0, 1.0, 0.0);
        //        ModelObjectVisualization.SetTemporaryState(myNocommentListIdentifier, tsgreen);
        //        //ModelObjectVisualization.SetTransparencyForAll(TemporaryTransparency.HIDDEN);



        //        //public static bool ClearAllTemporaryStates();
        //        //public static bool GetRepresentation(ModelObject modelObject, ref Color color);
        //        //public static bool SetTemporaryState(List<ModelObject> modelObjects, Color color);
        //        //public static bool SetTemporaryState(List<Identifier> identifiers, Color color);
        //        //public static bool SetTemporaryStateForAll(Color color);
        //        //public static bool SetTransparency(List<ModelObject> modelObjects, TemporaryTransparency transparency);
        //        //public static bool SetTransparency(List<Identifier> identifiers, TemporaryTransparency transparency);
        //        //public static bool SetTransparencyForAll(TemporaryTransparency transparency);

        //        //ModelObjectVisualization.ClearAllTemporaryStates();
        //        //ModelObjectVisualization.SetTransparencyForAll(TemporaryTransparency.HIDDEN);

        //        //var blueObjects = new List<ModelObject> { b, b2 };

        //        //var redObjects = new List<ModelObject> { b1 };


        //        //ModelObjectVisualization.SetTemporaryState(blueObjects, new TSM.UI.Color(0.0, 0.0, 1.0));
        //        //ModelObjectVisualization.SetTemporaryState(redObjects, new TSM.UI.Color(1.0, 0.0, 0.0));


        //        //TSMUI.ModelObjectVisualization.ClearAllTemporaryStates();
        //        //bool inset = TSMUI.ModelObjectVisualization.SetTemporaryStateForAll(TSMUI.Color  System.Drawing.Color.Gray);
        //        //TSMUI.ModelObjectVisualization(ModelObjectSelector)

        //        //TSMUI.dotTemporaryState.SetState(TSMUI.dotTemporaryStatesEnum.DOT_TEMPORARY_STATE_UNCHANGED, TSMUI.dotTemporaryTransparenciesEnum.DOT_TEMPORARY_TRANSPARENCY_TRANSPARENT);
        //        //TSMUI.dotTemporaryState.SetState(MyPartListDelivered, TSMI.dotTemporaryStatesEnum.DOT_TEMPORARY_STATE_ORIGINAL, TSMI.dotTemporaryTransparenciesEnum.DOT_TEMPORARY_TRANSPARENCY_VISIBLE);
        //        //TSMUI.dotTemporaryState.SetState(MyPartListCast, TSMI.dotTemporaryStatesEnum.DOT_TEMPORARY_STATE_ACTIVE, TSMI.dotTemporaryTransparenciesEnum.DOT_TEMPORARY_TRANSPARENCY_VISIBLE);
        //        //TSMUI.dotTemporaryState.SetState(MyPartListPlanned, TSMI.dotTemporaryStatesEnum.DOT_TEMPORARY_STATE_MODIFIED, TSMI.dotTemporaryTransparenciesEnum.DOT_TEMPORARY_TRANSPARENCY_VISIBLE);

        //    }

        //}

        private void pictureBox2_DoubleClick(object sender, EventArgs e)
        {
            ModelObjectVisualization.ClearAllTemporaryStates();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void tvwncm_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString().ToUpper() == "R" && e.Shift == true)
            {
                //refresh tree view
                Update_DashBoard_TreeView();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "btnDelete_Click", "", "Pl.Wait,deleting record(s)...", "btnDelete_Click", lblsbar1, lblsbar2);
            int delct = 0;
            if (isowner == true)
            {
                //tpgqc.Tag = cmbqmsdblist.Text;;
                //guid's can be added during new / inprocess only hence allowed until process is below 20
                if (id_process < (int)SKProcess.ModelChecking_PM_Approved)
                {
                    //tpgqc.Tag = cmbqmsdblist.Text;;
                    delct = DeleteRow(dgqc);

                    //based on row count set control visiblity
                    SetControlVisible_QC_DataGridRowCount();

                }
                else
                {
                    //only checker can select the model guid's and push to server
                }
            }
            else
            {
                //since not owner no action required
                MessageBox.Show("You are not the owner for report " + tpgqc.Text + "  hence aborting now!!!", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "btnDelete_Click", delct.ToString() + " record(s) deleted.", "", "btnDelete_Click", lblsbar1, lblsbar2, s_tm);

        }
        //delete record from db
        private int DeleteRecord(string idqc)
        {
            //time stamp
            long timestamp = new DateTimeOffset(DateTime.Now).ToUnixTimeMilliseconds();
            mysql = "UPDATE `ncm` SET status = 'DELETED' ,status_user_id = " + id_emp + " , status_sys = '" + skWinLib.systemname + "' , status_dt = " + timestamp + " WHERE status is null and idqc in (" + idqc + ")";
            mycmd.CommandText = mysql;
            mycmd.Connection = myconn;
            if (myconn.State == ConnectionState.Closed)
                myconn.Open();
            //update record
            return mycmd.ExecuteNonQuery();
        }
        //delete row from the select rows
        private int DeleteRow(DataGridView myDataGridView)
        {
            string s_idqc = string.Empty;
            string del_ids = ",";
            int del_ct = 0;
            foreach (DataGridViewRow row in myDataGridView.SelectedRows)
            {
                if (row.Cells["idqc"].Value != null)
                {
                    s_idqc = row.Cells["idqc"].Value.ToString();
                    if (skWinLib.IsNumeric(s_idqc) == true)
                        del_ids = del_ids + "," + s_idqc;
                }
                //delete row since no idqc found (i.e its new entry not stored in db)
                myDataGridView.Rows.Remove(row);
                del_ct++;
            }

            //based on del_id's update database and set these id's are deleted
            del_ids = del_ids.Replace(",,", "");
            if (del_ids != ",")
            {
                //delete record
                del_ct = DeleteRecord(del_ids);

            }

            ////this function will work for idqc and not for new row's (next release or future development)
            ////remove row in mirror data
            //del_ids = "," + del_ids + ",";
            //foreach (DataGridViewRow row in dgqcmirror.Rows)
            //{
            //    if (row.Cells["idqc"].Value != null)
            //    {
            //        s_idqc = row.Cells["idqc"].Value.ToString();
            //        if (del_ids.IndexOf("," + s_idqc + ",") >=0)
            //        {
            //            //delete row since no idqc found (i.e its new entry not stored in db)
            //            dgqcmirror.Rows.Remove(row);
            //        }

            //    }
            //}

            bool flag = false;
            for (int rowct = 0; rowct < dgqcmirror.RowCount; rowct++)
            {
                DataGridViewRow row = dgqcmirror.Rows[rowct];
                if (row.Cells["idqc"].Value != null)
                {
                    s_idqc = row.Cells["idqc"].Value.ToString();
                    if (del_ids.IndexOf("," + s_idqc + ",") >= 0)
                    {
                        //delete row and reset rowct - 1
                        dgqcmirror.Rows.Remove(row);
                        rowct--;
                        flag = true;
                    }

                }
            }
            //update row header and treeview 
            if (del_ct >= 1)
            {
                //update row header
                skWinLib.updaterowheader(dgqc);
                if (flag == true)
                    skWinLib.updaterowheader(dgqcmirror);

                //update dashboard
                Update_DashBoard_TreeView();
            }


            return del_ct;
        }

        private void btnhlp_Click(object sender, EventArgs e)
        {



                skWinLib.SendEmail(skApplicationVersion, skTSLib.ModelName, lblreport.Text , lblxsuser.Text, "","");
                //skWinLib.SendEmail(skApplicationVersion, skTSLib.ModelName, report, lblxsuser.Text, updtct.ToString(), processlog);

            //send ipmsg
            //"ipcmd send "192.168.2.35" "ipmsg - cmd - line - checking""
            if (sk_pdf_file_display == false)
            {
                string pdfhelpfile = "D:\\ESSKAY\\SK04-NCM.pdf";
                //string pdfhelpfile = "D:\\Vijayakumar\\10_Documention\\SK04-NCM.pdf";
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                process.StartInfo = startInfo;

                startInfo.FileName = @pdfhelpfile;

                process.Start();
                sk_pdf_file_display = true;
            }

        }

        private void btnhlp_KeyUp(object sender, KeyEventArgs e)
        {
            ////function which will fetch timesheet data from the begining
            ////this will be amended in future with start date and end date condition.
            ////ALT + CTRL
            //if (e.Alt == true && e.KeyValue == 17)
            //{
            //    FetchTSData();
            //    skWinLib.Export_DataGridView(dgvts, "d:\\esskay\\export_sk_timesheet_data.csv", "PIE", "");
            //    MessageBox.Show(dgvts.Rows.Count.ToString() + " records exported.\n\nTimeStamp : " + DateTime.Now.ToString(), msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }

        private void lblrole_Click(object sender, EventArgs e)
        {

        }

        private void dgqc_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //set enable or disable of button based on row count in datagrid qc
            //setpreview();
        }

        private void setpreview()
        {
            bool flag = false;
            //if row count is zero disable preview button
            if (dgqc.Rows.Count == 0)
            {
                btntspreview2.Enabled = false;
            }
            else
            {
                flag = true;
                if (id_process <= (int)SKProcess.Modelling_Completed)
                {
                    btntspreview2.Enabled = false;
                }
                else if (id_process <= (int)SKProcess.ModelBackDraft_Completed)
                {
                    btntspreview2.Enabled = true;
                }
                else if (id_process <= (int)SKProcess.ModelBackCheck_Completed)
                {
                    btntspreview2.Enabled = false;
                }


            }
            btnnextprocess.Enabled = flag;
            //btnDelete.Enabled = flag;
            btnSave.Enabled = flag;
            btnUnique.Enabled = flag;
            //btnexport.Enabled = flag;
            btnsplitsave.Enabled = flag;
            btnnextprocess.Enabled = flag;
            btntspreview2.Enabled = flag;
            btnclear.Enabled = flag;
            btnpreviousprocess.Enabled = flag;
            btnnextprocess.Enabled = flag;
            btnrebackdraft.Enabled = flag;


        }



        private void cmbqmsdblist_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dgqc_Enter(object sender, EventArgs e)
        {

        }
        //function which create column, row, cell value as per other tabpage
        private void MirrorTabPageData()
        {
            //dgqcmirror.DataSource  = dgqc.DataSource ;
            skWinLib.DataGridView_Setting_Before(dgqcmirror);

            dgqcmirror.Rows.Clear();
            dgqcmirror.Columns.Clear();
            ////add column
            foreach (DataGridViewColumn c in dgqc.Columns)
            {
                dgqcmirror.Columns.Add(c.Clone() as DataGridViewColumn);
                //refine the code in future
                //if (c.Name == "idqc")
                //{
                //    dgqcmirror.Columns.Add(c.Clone() as DataGridViewColumn);
                //}
                //else
                //{
                //    if (c.Visible == true)
                //        dgqcmirror.Columns.Add(c.Clone() as DataGridViewColumn);
                //}

            }
            //add row
            foreach (DataGridViewRow r in dgqc.Rows)
            {

                int index = dgqcmirror.Rows.Add(r.Clone() as DataGridViewRow);
                foreach (DataGridViewCell o in r.Cells)
                {
                    dgqcmirror.Rows[index].Cells[o.ColumnIndex].Value = o.Value;
                    //refine the code in future
                    //string rcolname = o.OwningColumn.Name;
                    //if (rcolname == "idqc")
                    //{
                    //    dgqcmirror.Rows[index].Cells[o.ColumnIndex].Value = o.Value;
                    //}
                    //else
                    //{
                    //    if (o.OwningColumn.Visible == true)
                    //        dgqcmirror.Rows[index].Cells[o.ColumnIndex].Value = o.Value;
                    //}

                }
            }
            skWinLib.DataGridView_Setting_After(dgqcmirror);
        }
        private void btncopy_Click(object sender, EventArgs e)
        {
            //tpgqc.Tag = cmbqmsdblist.Text;;
            ArrayList mydata = getDifference_Datagrid();
        }
        //private ArrayList getDifference_Datagrid()
        //{
        //    ArrayList columndata = new ArrayList();
        //    columndata.Add(columndata);
        //    columndata.Add(columndata);
        //    columndata.Add(columndata);
        //    return columndata;
        //}
        private ArrayList getDifference_Datagrid()
        {
            //compare read only columnn set as false and visible set as true
            ArrayList column_name = new ArrayList();
            ArrayList sql_column_name = new ArrayList();
            ArrayList sql_column_type = new ArrayList();

            //current datagridview with idqc (existing)
            ArrayList current_update = new ArrayList();


            //current datagridview without idqc (new)
            ArrayList current_new = new ArrayList();

            //mirror datagridview with idqc (existing)
            ArrayList mirror_exist = new ArrayList();


            //check column name 
            column_name.Add("idqc");
            //column_name.Add("guid");
            //column_name.Add("type");

            //sql column name
            sql_column_name.Add("in|idqc");
            //sql_column_name.Add("st|guid");
            //sql_column_name.Add("st|guid_type");


            string mydata = "";
            if (dgqc.RowCount >= 1)
            {
                for (int i = 0; i < dgqc.Columns.Count; i++)
                {
                    if (dgqc.Columns[i].Name.Contains("TeklaSK") == false)
                    {
                        //dgqc.Rows[0].Cells[i].OwningColumn.Visible == true && 
                        if (dgqc.Rows[0].Cells[i].OwningColumn.Visible == true && dgqc.Rows[0].Cells[i].OwningColumn.ReadOnly == false)
                        {
                            column_name.Add(dgqc.Rows[0].Cells[i].OwningColumn.Name);
                            sql_column_name.Add(dgqc.Rows[0].Cells[i].OwningColumn.Tag.ToString());

                        }
                    }


                }

                //store mirror datagridview information
                for (int i = 0; i < dgqcmirror.RowCount; i++)
                {
                    mydata = dgqcmirror.Rows[i].Cells["idqc"].Value.ToString();
                    for (int j = 1; j < column_name.Count; j++)
                    {
                        string chkcolname = column_name[j].ToString();
                        for (int colid = 0; colid < dgqcmirror.Columns.Count; colid++)
                        {
                            if (dgqcmirror.Rows[i].Cells[colid].OwningColumn.Name == chkcolname)
                            {
                                if (dgqcmirror.Rows[i].Cells[column_name[j].ToString()].Value != null)
                                    mydata = mydata + sk_seperator + skWinLib.mysql_replace_quotecomma(dgqcmirror.Rows[i].Cells[column_name[j].ToString()].Value.ToString());
                                else
                                    mydata = mydata + sk_seperator + "";
                            }

                        }
                    }

                    mirror_exist.Add(mydata);
                }

                //current datagridview information
                for (int i = 0; i < dgqc.RowCount; i++)
                {
                    //check each row
                    DataGridViewRow row = dgqc.Rows[i];
                    if (row.Cells["idqc"].Value != null)
                    {
                        mydata = row.Cells["idqc"].Value.ToString();
                        for (int j = 1; j < column_name.Count; j++)
                        {
                            string chkcolname = column_name[j].ToString();
                            for (int colid = 0; colid < dgqc.Columns.Count; colid++)
                            {
                                if (dgqcmirror.Rows[i].Cells[colid].OwningColumn.Name == chkcolname)
                                {
                                    if (row.Cells[column_name[j].ToString()].Value != null)
                                        mydata = mydata + sk_seperator + skWinLib.mysql_replace_quotecomma(row.Cells[column_name[j].ToString()].Value.ToString());
                                    else
                                        mydata = mydata + sk_seperator + "";
                                }
                            }

                        }
                        if (!mirror_exist.Contains(mydata))
                        {
                            //modified or new data by the user

                            if (row.Cells["idqc"].Value != null)
                            {
                                //guid,guid_type,obser,chk_rmk,id_ncm_progress
                                current_update.Add(mydata);
                            }

                        }
                        else
                            mirror_exist.Remove(mydata);
                    }
                    else
                    {
                        //newly selected by user added to datagrid now
                        current_new.Add(row);
                    }
                }

            }

            //same items are removed in current_user_info and only we have difference item
            ArrayList mylist = new ArrayList();
            mylist.Add(column_name);
            mylist.Add(sql_column_name);
            mylist.Add(current_new);
            mylist.Add(current_update);
            mylist.Add(mirror_exist);
            return mylist;
        }

        private string bulk_newinsertquery(DataGridViewRow row, int id_report, int id_sk_progress, int id_emp, long timestamp)
        {
            //check observation is not null or empty space
            string obser = string.Empty;
            string guid = string.Empty;
            string guid_type = string.Empty;
            string remark = string.Empty;



            //get guid
            if (row.Cells["guid"].Value != null)
                guid = row.Cells["guid"].Value.ToString();

            //get guidtype
            if (row.Cells["TeklaSKtype"].Value != null)
                guid_type = row.Cells["TeklaSKtype"].Value.ToString();

            //check idqc
            int idqc = -1;
            if (row.Cells["idqc"].Value != null)
            {
                string tmp = row.Cells["idqc"].Value.ToString();
                if (skWinLib.IsNumeric(tmp) == true)
                    idqc = Convert.ToInt32(tmp);
            }

            if (row.Cells["remark"].Value != null)
                remark = skWinLib.mysql_replace_quotecomma(row.Cells["remark"].Value.ToString());

            return " , ('" + guid + "','" + guid_type + "','" + remark + "'," + id_report + "," + id_sk_progress + "," + id_emp + "," + timestamp + ")";
        }

        //private ArrayList getDifference_Datagrid()
        //{
        //    //compare read only columnn set as false and visible set as true
        //    ArrayList columndata = new ArrayList();

        //    //current datagridview with idqc (existing)
        //    ArrayList current_exist = new ArrayList();
        //    //current datagridview without idqc (new)
        //    ArrayList current_new = new ArrayList();

        //    //mirror datagridview with idqc (existing)
        //    ArrayList mirror_exist = new ArrayList();


        //    ArrayList current_user_info = new ArrayList();
        //    ArrayList beforesave = new ArrayList();

        //    //guid,guid_type,obser,chk_rmk,id_ncm_progress
        //    ArrayList newdata = new ArrayList();
        //    columndata.Add("idqc");
        //    columndata.Add("guid");
        //    columndata.Add("type");
        //    string mydata = "";
        //    if (dgqc.RowCount >= 1)
        //    {
        //        for (int i = 0; i < dgqc.Columns.Count; i++)
        //        {
        //            if (dgqc.Rows[0].Cells[i].OwningColumn.Visible == true && dgqc.Rows[0].Cells[i].OwningColumn.ReadOnly == false)
        //                columndata.Add(dgqc.Rows[0].Cells[i].OwningColumn.Name);
        //        }

        //        //mirror datagridview information
        //        for (int i = 0; i < dgqcmirror.RowCount; i++)
        //        {
        //            mydata = "";
        //            for (int j = 0; j < columndata.Count; j++)
        //            {
        //                if (dgqcmirror.Rows[i].Cells[columndata[j].ToString()].Value != null)
        //                    mydata = mydata + sk_seperator + dgqcmirror.Rows[i].Cells[columndata[j].ToString()].Value;
        //                else
        //                    mydata = mydata + sk_seperator + "";
        //            }
        //            mirror_exist.Add(mydata);
        //        }

        //        //current datagridview information
        //        for (int i = 0; i < dgqc.RowCount; i++)
        //        {
        //            mydata = "";
        //            int idqc = -1;
        //            if (row.Cells["idqc"].Value != null)
        //            {
        //                string tmp = row.Cells["idqc"].Value.ToString();
        //                if (skWinLib.IsNumeric(tmp) == true)
        //                    idqc = Convert.ToInt32(tmp);
        //            }


        //            for (int j = 1; j < columndata.Count; j++)
        //            {
        //                if (dgqc.Rows[i].Cells[columndata[j].ToString()].Value != null)
        //                    mydata = mydata + sk_seperator + dgqc.Rows[i].Cells[columndata[j].ToString()].Value;
        //                else
        //                    mydata = mydata + sk_seperator + "";
        //            }
        //            current_user_info.Add(mydata);
        //        }
        //        current_user_info.Sort();

        //        //mirror datagridview information
        //        for (int i = 0; i < dgqcmirror.RowCount; i++)
        //        {
        //            mydata = "";
        //            for (int j = 0; j < columndata.Count; j++)
        //            {
        //                if (dgqcmirror.Rows[i].Cells[columndata[j].ToString()].Value != null)
        //                    mydata = mydata + sk_seperator + dgqcmirror.Rows[i].Cells[columndata[j].ToString()].Value;
        //                else
        //                    mydata = mydata + sk_seperator + "";
        //            }

        //            //check same data exist in current datagridview
        //            if (!current_user_info.Contains(mydata))
        //            {
        //                beforesave.Add(mydata);
        //            }
        //            else
        //            {
        //                current_user_info.Remove(mydata);
        //                //delete earlier data to avoid comparing again time saving performance improved
        //            }

        //        }
        //        beforesave.Sort();
        //    }

        //    //same items are removed in current_user_info and only we have difference item
        //    ArrayList mylist = new ArrayList();
        //    mylist.Add(columndata);
        //    mylist.Add(current_user_info);
        //    mylist.Add(beforesave);

        //    return mylist;
        //}
        private void dgqc_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void cmbqmsdblist_Leave(object sender, EventArgs e)
        {
            //if (TabControl1.SelectedIndex == 0)
            //    misentry = cmbqmsdblist.Text;
            //else
            //    qcentry = cmbqmsdblist.Text;

        }

        private void btnfilter_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "btnfilter_Click", "", "Pl.Wait,filtering guid(s)...", "btnfilter_Click", lblsbar1, lblsbar2);
            int fct = 0;
            string msg = "";
            //select objects in model and store the guid in a array list
            //if datagrid row >1 then loop through each row and check whether guid is in the array list if not then remove that row

            if (dgqc.RowCount >= 1)
            {
                TSM.ModelObjectEnumerator MyEnum = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
                ArrayList myguidlist = new ArrayList();
                while (MyEnum.MoveNext())
                {
                    TSM.ModelObject current = MyEnum.Current;
                    myguidlist.Add(current.Identifier.GUID.ToString());
                }


                ////delete guid's not listed
                //foreach(DataGridViewRow row in dgqc.Rows)
                //{
                //    string guid = string.Empty;
                //    //get guid
                //    if (row.Cells["guid"].Value != null)
                //        guid = row.Cells["guid"].Value.ToString();
                //    if (!myguidlist.Contains(guid))
                //    {
                //        dgqc.Rows.Remove(row);
                //    }

                //}
                bool flag = false;
                for (int rowct = 0; rowct < dgqc.RowCount; rowct++)
                {
                    DataGridViewRow row = dgqc.Rows[rowct];
                    string guid = string.Empty;
                    //get guid
                    if (row.Cells["guid"].Value != null)
                        guid = row.Cells["guid"].Value.ToString();
                    if (!myguidlist.Contains(guid))
                    {
                        dgqc.Rows.Remove(row);
                        rowct--;
                        flag = true;
                    }
                }
                if (flag == true)
                {
                    //reset row header
                    skWinLib.updaterowheader(dgqc);

                }
            }
            else
            {
                msg = "Nothing to Filter.";
            }


            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "btnfilter_Click", fct.ToString() + " filter(s).", msg, "btnfilter_Click", lblsbar1, lblsbar2, s_tm);

        }

        private void dgqc_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("1");
        }

        private void dgqc_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //set cmbqmsdblist text with observation/remark if observation or remark is not readonly 
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                string colname = dgqc.Columns[e.ColumnIndex].Name;
                bool bolcol = dgqc.Columns[e.ColumnIndex].Visible;
                if (bolcol == true)
                {
                    if (colname == "Observation")
                    {
                        if (dgqc.Rows[e.RowIndex].Cells["observation"].Value != null)
                            cmbqmsdblist.Text = dgqc.Rows[e.RowIndex].Cells["observation"].Value.ToString();

                    }
                    else if (colname == "Remark")
                    {
                        if (dgqc.Rows[e.RowIndex].Cells["remark"].Value != null)
                            cmbqmsdblist.Text = dgqc.Rows[e.RowIndex].Cells["remark"].Value.ToString();

                    }
                }



            }

        }

        private void lblqmsdblist_Click(object sender, EventArgs e)
        {

        }

        private void tpgdash_Click(object sender, EventArgs e)
        {

        }

        private void btnsetstage_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Move_To_BeforeProcess", "", "Pl.Wait,processing...", "btnsetstage_Click", lblsbar1, lblsbar2);
            string msg = "";
            if (dgqc.SelectedRows.Count >= 1)
            {
                int updtct = Move_To_BeforeProcess(SK_id_report, SK_id_process);
                if (updtct >= 1)
                {
                    //refresh
                    Update_DashBoard_TreeView();

                    //fetch data from mysql
                    Fetch_Model_Data();

                    //store mirror data for future comparision
                    MirrorTabPageData();

                    msg = updtct.ToString() + " record(s) moved to earlier stage";
                }
                else
                    msg = dgqc.SelectedRows.Count.ToString() + " record(s) can't be moved at this time. Contact Support Team.";
            }
            else
                msg = "No record(s) selected!!!";
            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "Move_To_BeforeProcess", "", msg, "btnsetstage_Click", lblsbar1, lblsbar2, s_tm);
        }

        private void lblstage_Click(object sender, EventArgs e)
        {

        }

        private void btnrefresh_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Refresh", "", "Pl.Wait,refreshing...", "btnrefresh_Click", lblsbar1, lblsbar2);
            string msg = "";
            esskayjn = txtproject.Text.ToString();

            UpdateUserInterface_Fromlastsave();

            //refresh treeview
            Update_DashBoard_TreeView();

            //refresh report number
            Update_Dashboard_ReportNumber();

            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "Refresh", "", msg, "btnrefresh_Click", lblsbar1, lblsbar2, s_tm);
        }

        private void btnTSClear_Click(object sender, EventArgs e)
        {
            //clear all tekla attribute
            for (int i = 0; i < chkTSAttribute.Items.Count; i++)
                chkTSAttribute.SetItemChecked(i, false);
        }

        private void btnTSAll_Click(object sender, EventArgs e)
        {
            if (sk_TOGGLE_ts_input == false)
                sk_TOGGLE_ts_input = true;  
            else 
                sk_TOGGLE_ts_input = false;


            //set all tekla attribute
            for (int i = 0; i < chkTSAttribute.Items.Count; i++)
                chkTSAttribute.SetItemChecked(i, sk_TOGGLE_ts_input);

        }

        private void btnpreviewproject_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "PreviewProject", "", "Pl.Wait,processing model preview... ", "btnpreviewproject_Click", lblsbar1, lblsbar2);
            string msg = Preview_NCM(TabControl1.SelectedTab.Name.ToUpper().Trim());
            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "PreviewProject", "", msg, "btnpreviewproject_Click", lblsbar1, lblsbar2, s_tm);

        }

        private string Preview_NCM(string CurrentTabName)
        {
            string msg = "";
            if (CurrentTabName.ToUpper() == "TPGQC")
            {

                //in model display green for nocomment, red for other observation and blue for balance
                ArrayList ObjectsToSelect = new ArrayList();

                if (sk_TOGGLE_ts_display == false)
                {
                    //Tekla
                    List<Identifier> myListIdentifier = new List<Identifier>();
                    List<Identifier> myNocommentListIdentifier = new List<Identifier>();
                    //tpgqc.Tag = cmbqmsdblist.Text;;

                    try
                    {
                        //int rowindex = e.RowIndex;
                        //DataGridViewRow row = mydatagrid.Rows[rowindex];



                        int i = 0;
                        foreach (DataGridViewRow row in dgqc.Rows)
                        {
                            Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier();
                            string myguid = row.Cells["GUID"].Value.ToString();
                            myguid = myguid.ToUpper().Replace("GUID:", "").Replace("GUID", "").Trim();
                            myIdentifier = new Tekla.Structures.Identifier(myguid);
                            bool flag = myIdentifier.IsValid();
                            if (flag)
                            {
                                Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
                                if (ModelSideObject.Identifier.IsValid())
                                {
                                    if (!ObjectsToSelect.Contains(ModelSideObject))
                                        ObjectsToSelect.Add(ModelSideObject);

                                    bool processflag = false;
                                    if (id_process < (int)SKProcess.ModelChecking_Completed)
                                    {
                                        string myobser = string.Empty;
                                        if (row.Cells["observation"].Value != null)
                                        {
                                            myobser = row.Cells["observation"].Value.ToString();
                                            if (myobser == sk_nocomment || myobser == "\nNo Comment")
                                                processflag = true;
                                        }
                                    }
                                    else if (id_process < (int)SKProcess.ModelBackDraft_Completed)
                                    {
                                        if (row.Cells["isagreed"].Value != null)
                                        {
                                            string agreed = row.Cells["isagreed"].Value.ToString();
                                            if (agreed.ToUpper().Trim() == "1" || agreed.ToUpper().Trim() == "TRUE")
                                                processflag = true;
                                        }
                                    }
                                    else if (id_process < (int)SKProcess.ModelBackCheck_Completed)
                                    {
                                        if (row.Cells["backcheck"].Value != null)
                                        {
                                            string backcheck = row.Cells["backcheck"].Value.ToString();
                                            if (backcheck.ToUpper().Trim() == "1" || backcheck.ToUpper().Trim() == "TRUE")
                                                processflag = true;
                                        }

                                    }


                                    if (processflag == true)
                                    {
                                        if (!myNocommentListIdentifier.Contains(myIdentifier))
                                            myNocommentListIdentifier.Add(myIdentifier);
                                    }
                                    else if (!myListIdentifier.Contains(myIdentifier))
                                        myListIdentifier.Add(myIdentifier);
                                }
                            }
                            else
                            {
                                //set ui color as red since these guid object are not avaialble in model (i.e deleted)
                                dgqc.Rows[row.Index].DefaultCellStyle.BackColor = System.Drawing.Color.Red;
                            }
                            i++;
                        }

                        //var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                        //MOS.Select(ObjectsToSelect);
                        if (myNocommentListIdentifier.Count >= 1)
                            skTSLib.ObjectVisualization(myListIdentifier, myNocommentListIdentifier);
                        else
                            skTSLib.ObjectVisualization(myListIdentifier);



                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Representation Error: " + ex.Message);
                    }


                    sk_TOGGLE_ts_display = true;
                }
                else
                {
                    ModelObjectVisualization.ClearAllTemporaryStates();
                    //var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    //MOS.Select(ObjectsToSelect);
                    sk_TOGGLE_ts_display = false;
                }
            }
            else if (CurrentTabName.ToUpper() == "TPGDASH")
            {
                if (sk_TOGGLE_project_display == false)
                {
                    sk_TOGGLE_project_display = true;

                    if (dgreportid.Rows.Count >= 1)
                    {

                        //clear all
                        TSMUI.ModelObjectVisualization.ClearAllTemporaryStates();


                        //all objects set as grey
                        TSMUI.Color tsgrey = new TSM.UI.Color(0.75, 0.75, 0.75);
                        TSMUI.ModelObjectVisualization.SetTemporaryStateForAll(tsgrey);

                        foreach (DataGridViewRow row in dgreportid.SelectedRows)
                        {
                            string reportname = row.Cells["rptname"].Value.ToString();
                            List<Identifier> myListIdentifier = new List<Identifier>();
                            myListIdentifier = getTeklaGUID(row.Cells["rptid"].Value.ToString());
                            if (myListIdentifier.Count >= 1)
                            {
                                var values = Guid.NewGuid().ToByteArray().Select(b => (int)b);
                                double red = values.Take(5).Sum() % 255;
                                double green = values.Skip(5).Take(5).Sum() % 255;
                                double blue = values.Skip(10).Take(5).Sum() % 255;

                                red = red / 255;
                                green = green / 255;
                                blue = blue / 255;

                                // TSMUI.Color mytscolor = new TSM.UI.Color(myred, mygreen, myblue);
                                TSMUI.Color mytscolor = new TSM.UI.Color(red, green, blue);
                                TSMUI.ModelObjectVisualization.SetTemporaryState(myListIdentifier, mytscolor);
                            }

                        }
                    }
                }
                else
                {
                    sk_TOGGLE_project_display = false;

                    //clear all
                    TSMUI.ModelObjectVisualization.ClearAllTemporaryStates();

                }

            }
            return msg;
        }
        private List<Identifier> getTeklaGUID(string id_report)
        {

            List<Identifier> myListIdentifier = new List<Identifier>();

            ArrayList ObjectsToSelect = new ArrayList();
            try
            {
                mysql = "select distinct guid from epmatdb.ncm where id_report = " + id_report + " order by id_report";
                if (myconn.State == ConnectionState.Closed)
                    myconn.Open();
                mycmd.CommandText = mysql;
                mycmd.Connection = myconn;
                myrdr = mycmd.ExecuteReader();
                Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier();
                while (myrdr.Read())
                {
                    myIdentifier = new Tekla.Structures.Identifier(myrdr[0].ToString());
                    if (myIdentifier.IsValid())
                    {
                        myListIdentifier.Add(myIdentifier);
                        //Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
                        //if (!ObjectsToSelect.Contains(ModelSideObject))
                        //    ObjectsToSelect.Add(ModelSideObject);
                    }
                }

                //var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                //MOS.Select(ObjectsToSelect);

                myrdr.Close();
                myconn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetTeklaGUID Error \n\n" + ex.Message + "\n\n" + ex.StackTrace + "\n\n" + ex.Source + "\n\n" + ex.Data.ToString());
            }
            return myListIdentifier;
        }

        private void chkzoom_CheckedChanged(object sender, EventArgs e)
        {
            //    GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            //    var chatMessage = new ChatMessage
            //    {
            //        Body = new ItemBody
            //        {
            //            Content = "Hello World"
            //        }
            //    };

            //    await graphClient.Teams["{team-id}"].Channels["{channel-id}"].Messages
            //        .Request()
            //        .AddAsync(chatMessage);
            //}
        }

        private void btnrebackdraft_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "Move_To_BackDraftProcess", "", "Pl.Wait,processing...", "btnrebackdraft_Click", lblsbar1, lblsbar2);
            string msg = "";
            if (dgqc.SelectedRows.Count >= 1)
            {
                msg = CopySelectedRows_To_BackDraftProcess();
                if (msg.IndexOf("records") >= 0)
                {
                    //refresh
                    Update_DashBoard_TreeView();

                    //fetch data from mysql
                    Fetch_Model_Data();

                    //store mirror data for future comparision
                    MirrorTabPageData();

                }
            }
            else
                msg = "No record(s) selected!!!";
            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "Move_To_BackDraftProcess", "", msg, "btnrebackdraft_Click", lblsbar1, lblsbar2, s_tm);
            
            MessageBox.Show(msg, msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        private void dgreportid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgreportid_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "NewReportAddition", "", "Pl.Wait,processing model check report... ", "dgreportid_RowHeaderMouseDoubleClick", lblsbar1, lblsbar2);
            string msg = "";
            if (dgreportid.SelectedRows.Count >= 1)
            {
                foreach (DataGridViewRow row in dgreportid.SelectedRows)
                {
                    string rptid =  row.Cells["rptid"].Value.ToString();
                    string rptname = row.Cells["rptname"].Value.ToString();
                    //whether report name already exists
                    bool newflag = false;
                    foreach (TreeNode mychildnode in tvwnew.Nodes)
                    {
                        if (mychildnode.Text == rptname)
                        {
                            newflag = true;
                            break;
                        }
                    }
                    if (newflag == false)
                    {
                        TreeNode tvwnewnode = tvwnew.Nodes.Add(rptname);
                        tvwnewnode.Tag = rptid.ToString() + "|" + (int)SKProcess.Modelling;
                    }

                }
            }
            skTSLib.setCompletion(skApplicationName, skApplicationVersion, "NewReportAddition", "", msg, "dgreportid_RowHeaderMouseDoubleClick", lblsbar1, lblsbar2, s_tm);

        }

        private void btnnextprocess_Click(object sender, EventArgs e)
        {
            DateTime s_tm = skWinLib.setStartup(skApplicationName, skApplicationVersion, "CompleteSelectedTask", "", "Pl.Wait,saving data's", "btnnextprocess_Click", lblsbar1, lblsbar2);
            string selected_idqc = ",";
            foreach (DataGridViewRow row in dgqc.SelectedRows)
            {
                string idqc = string.Empty;
                //idqc
                if (row.Cells["idqc"].Value != null)
                {
                    idqc = row.Cells["idqc"].Value.ToString();
                    selected_idqc = selected_idqc + "," + idqc;
                }
            }
            if (selected_idqc != ",")
            {
                selected_idqc = selected_idqc.Replace(",,", "");
                selected_idqc = " and idqc in (" + selected_idqc + ")";
                MoveReport_NextProcess(SK_id_report, SK_id_process, tpgqc.Text.ToString(), selected_idqc);

                //refresh
                Update_DashBoard_TreeView();

                //fetch data from mysql
                Fetch_Model_Data();

                //store mirror data for future comparision
                MirrorTabPageData();
            }
            skWinLib.setCompletion(skApplicationName, skApplicationVersion, "CompleteSelectedTask", "", "", "btnnextprocess_Click", lblsbar1, lblsbar2, s_tm);
        }

        private void btnview_Click(object sender, EventArgs e)
        {
            //check whether null entry before fetching records
            //if (cmbqmsdblist.Text.Trim().Length >= 1)
            if (tpgqc.Text.Trim().Length >= 1)
            {
                //skreportname = cmbqmsdblist.Text;
                skreportname = tpgqc.Text;
                //tpgqc.Tag = cmbqmsdblist.Text;;
                //start log
                DateTime s_tm = skTSLib.setStartup(skApplicationName, skApplicationVersion, "Fetch", "", "Pl.Wait,fetching...", "btnview_Click", lblsbar1, lblsbar2);

                //int old_SK_id_process = SK_id_process;
                //string  old_stage = lblstage.Text;

                //SK_id_process = (int)SKProcess.Modelling_QC_Completed;
                //lblstage.Text = "Completed";

                //fetch data from mysql
                string msg = Fetch_Model_Data(true);

                //based on row count set control visiblity
                //SetControlVisible_QC_DataGridRowCount();

                btnSave.Enabled = false;
                btnpreviousprocess.Enabled = false;
                btnnextprocess.Enabled = false;
                btnsplitsave.Enabled = false;
                btnAdd.Enabled = false;
                btnDelete.Enabled = false;
                btnfilter.Enabled = true;
                btnclear.Enabled = true;
                btn1.Enabled = false;
                btn2.Enabled = false;




                //restore old
                //SK_id_process = old_SK_id_process;
                //lblstage.Text = old_stage;

                skTSLib.setCompletion(skApplicationName, skApplicationVersion, "Fetch", "", msg, "btnview_Click", lblsbar1, lblsbar2, s_tm);
            }
        }
    }

                
}
