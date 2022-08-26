///Omm Muruga 01Dec2021
// ###########################################################################################
// ### Name          : Connection Processor
// ### Version       : Beta
// ###               : Visual C# / Visual Studio 2019
// ### Created       : December 2021
// ### Modified      : 
// ### Author        : Vijayakumar
// ### Released      : 
// ### Comment       :
// ### Company       : Esskay design and structures pvt. ltd.
// ### Description   : Connection Checking Tool
// ###                 
// ###                 
// ###########################################################################################
//


using System.Runtime.InteropServices;
using System.Security;
using System.Globalization;

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
using TSDT = Tekla.Structures.Datatype;

using TSG = Tekla.Structures.Geometry3d;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Solid;
using Tekla.Structures.Model.UI;
using TSMUI = Tekla.Structures.Model.UI;
using T3D = Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Drawing;
using TSD = Tekla.Structures.Drawing;


namespace ConnectionProcessor
{




    public partial class frmconnectionprocess : Form
    {

        public string VersionDirectory = System.Environment.GetEnvironmentVariable("XSDATADIR");
        //public string strpath = System.Environment.GetEnvironmentVariable("XS_MACRO_DIRECTORY") + ";" + Path.GetDirectoryName(Application.ExecutablePath);
        public static string skApplicationName = "Connection Processor";
        public static string skApplicationVersion = "2208.15";

        public string skTSVersion = "2021";
        public bool skdebug = false;
        public int debugcount = 1;


        public static string msgCaption = skApplicationName + " Ver: " + skApplicationVersion;
        public string applicationlog = "";

        public ArrayList nodeguidlist = new ArrayList();
        int groupCount = 0;
        int TotalCount = 0;
        int errorCount = 0;
        public bool flagmetric = false;
        public bool flagimperial = false;
        
        //Tekla.Macros.Wpf.Runtime.IWpfMacroHost wpf = Tekla.Macros.Runtime.IMacroRuntime.Get<Tekla.Macros.Wpf.Runtime.IWpfMacroHost>();
        public TSM.Model MyModel;
        public string modelname;
        public string modelpath;
        public string configuration;
        public string TSversion;
        public string systempath = "";
        public string XS_DIR;
        public string XSBIN;
        public string XS_FIRM;
        public string XS_PROJECT;
        public string XS_UNIT;



        public string temp_projectpath = "d:\\esskay";
        public string firmpath = "d:\\esskay\\package_master";
        //public const string logpath = "d:\\esskay\\log";
        public const string logpath = "\\\\192.168.2.3\\Tekla Support\\log";


        public string DataDir = string.Empty;
        public string currEnvi = string.Empty;

        public string strpath = System.Environment.GetEnvironmentVariable("XS_MACRO_DIRECTORY") + ";" + Path.GetDirectoryName(Application.ExecutablePath);
        public string projectpath = System.Environment.GetEnvironmentVariable("XS_PROJECT") + ";" + System.Environment.GetEnvironmentVariable("XS_PROJECT_DEFAULT") + ";" + Path.GetDirectoryName(Application.ExecutablePath);
   




        public string gainfo = "";
        public string assyinfo = "";
        public string singleinfo = "";


        public frmconnectionprocess()
        {
           
            InitializeComponent();
            



            //application log entry


            Model MyModel = new Model();
            if (MyModel.GetConnectionStatus() && MyModel.GetInfo().ModelPath != string.Empty)
            {
                this.Text = skApplicationName + " Ver. " + skApplicationVersion + " [ " + skTSVersion + " ]";

                TeklaStructuresSettings.GetAdvancedOption("XS_SYSTEM", ref systempath);
                TeklaStructuresSettings.GetAdvancedOption("XSDATADIR", ref DataDir);
                TeklaStructuresSettings.GetAdvancedOption("XS_AD_ENVIRONMENT", ref currEnvi);
                TeklaStructuresSettings.GetAdvancedOption("XS_DIR", ref XS_DIR);
                TeklaStructuresSettings.GetAdvancedOption("XSBIN", ref XSBIN);
                TeklaStructuresSettings.GetAdvancedOption("XS_FIRM", ref XS_FIRM);
                TeklaStructuresSettings.GetAdvancedOption("XS_PROJECT", ref XS_PROJECT);
                TeklaStructuresSettings.GetAdvancedOption("XS_IMPERIAL", ref XS_UNIT );


                systempath = systempath.Replace("\\\\", "\\");
                modelname = MyModel.GetInfo().ModelName;
                modelpath = MyModel.GetInfo().ModelPath;
                configuration = ModuleManager.Configuration.ToString();
                TSversion = TeklaStructuresInfo.GetCurrentProgramVersion();

                lblsbar1.Text = modelname.Replace(".db1", "");
                lblsbar1.Tag = modelpath;
                lblsbar2.Text = modelpath;
            }
            string packname = DateTime.Now.ToString().Replace(":", "");
            packname = packname.Replace(" ", "");
            TeklaStructures.Connect();





        }

    
        private void addNCMOptioninDataGrid(DataGridViewRowsAddedEventArgs e)
        {
            lblsbar2.Text = "addNCMOptioninDataGrid";
            if (e.RowIndex >= 0)
            {
                int rowindex = e.RowIndex;
                DataGridViewRow row = this.dgattibutedata.Rows[rowindex];
                DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Category"]);
                cell.DataSource = new string[] { "Human Error", "Lackness - Technical", "Lackness – Software", "Interpretation issues", "Errors in Design documents", "Incomplete other trade coordination", "Communication failure", "Others." };
                cell = (DataGridViewComboBoxCell)(row.Cells["Severity"]);
                cell.DataSource = new string[] { "Minor", "Major", "Critical" };
            }
            /*
                //Category
                foreach (DataGridViewRow row in dgQC.Rows)
            {
                DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Category"]);
                cell.DataSource = new string[] { "Human Error", "Lackness - Technical", "Lackness – Software", "Interpretation issues", "Errors in Design documents", "Incomplete other trade coordination", "Communication failure", "Others." };
            }
            //Severity
            foreach (DataGridViewRow row in dgQC.Rows)
            {
                DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)(row.Cells["Severity"]);
                cell.DataSource = new string[] { "Minor", "Major", "Critical" };
            }
            */
            lblsbar2.Text = "";

        }
        private void frmPS_Load(object sender, EventArgs e)
        {

            ////ALWIN 
            //string retval1 = "888p";
            //string retval2 = "6";
            //string opt = "-";
            //string result = string.Empty;
            //if (opt=="+")
            //{
            //    double d_retval1 = 0;
            //    double d_retval2 = 0;

            //    if (skWinLib.IsNumeric(retval1) == true)
            //        d_retval1 = Convert.ToDouble(retval1);


            //    if (skWinLib.IsNumeric(retval2) == true)
            //        d_retval2 = Convert.ToDouble(retval2);

            //    result = (d_retval1 + d_retval2).ToString();
            //}
            //else if (opt == "-")
            //{
            //    double d_retval1 = 0;
            //    double d_retval2 = 0;

            //    if (skWinLib.IsNumeric(retval1) == true)
            //        d_retval1 = Convert.ToDouble(retval1);


            //    if (skWinLib.IsNumeric(retval2) == true)
            //        d_retval2 = Convert.ToDouble(retval2);

            //    result = (d_retval1 - d_retval2).ToString();
            //}
            //if (result != "")
            //{
            //    MessageBox.Show("Output " + result);
            //}






            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Load", "");
            //check for version
            MyModel = skTSLib.ModelAccess.ConnectToModel();

            if (MyModel == null)
            {
                MessageBox.Show("Could not establish connection with Tekla [ " + skTSVersion + " ].\n\n\nCheck Tekla version and model is open,if still issue's contact support team.", msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayTeklaValidation", "Aborted");
                //System.Windows.Forms.Application.ExitThread();
                //System.Windows.Forms.Application.Exit();
                System.Environment.Exit(0);
            }


            this.Text = skApplicationName + " Ver. " + skApplicationVersion + " [ " + skTSVersion + " ]";


            if (skWinLib.systemname  == "SK-01-200017")
                lblsbar2.Visible = true;
            else
                lblsbar2.Visible = false;
            lblxsuser.Text = skWinLib.username;
            tabControl1.SelectedIndex  = 0;






        }










        
        private DrawingHandler DrawingHandler = new DrawingHandler();
        private void test1()
        {

        }
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
            }
            lblsbar2.Text = "";

        }


        


        private void btngetallattibuteinfo_Click(object sender, EventArgs e)
        {
            ModelObjectEnumerator MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();

            while (MyEnum.MoveNext())
            {
                TSM.BaseComponent Component = MyEnum.Current as TSM.BaseComponent;
                if (Component == null) return;
                {
                    getAllAttibuteInformation(Component);
                    break;
                }

            }

        }

        private void chklsthdr_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void dgQC_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //lblsbar2.Text = "dgQC_CellContentClick";
            //if (e.RowIndex >= 0)
            //{
            //    int rowindex = e.RowIndex;
            //    DataGridViewRow row = this.dgattibutedata.Rows[rowindex];
            //    txtattfieldname.Text = row.Cells[0].Value.ToString();
            //    txtattfieldvalue.Text = row.Cells[2].Value.ToString();
            //}

            //lblsbar2.Text = "";
        }


        private void label1_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "label1_Click";
            lblsbar2.Text = "";
        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            lblsbar2.Text = "btnclear_Click";
            dgattibutedata.Rows.Clear();
            lblsbar2.Text = "";
        }





        private void dgQC_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            lblsbar2.Text = "dgQC_RowsAdded";
            //addNCMOptioninDataGrid(e);
            lblsbar2.Text = "";
        }

        private void dgQC_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        

        
        private void butget_Click(object sender, EventArgs e)
        {
            

        }

        

        private ArrayList getAllAttibuteInformation(Tekla.Structures.Model.BaseComponent Component)
        {
            int ct = 0;
            ArrayList returndata = new ArrayList();
            returndata.Add("Esskay");
            //CheckedListBox.CheckedItemCollection 
            //chklstattfield.Items.Clear();
            dgattibutedata.Rows.Clear();
            ArrayList sString = new ArrayList();
            ArrayList sInt = new ArrayList();
            ArrayList sDouble = new ArrayList();
            string sattval = "||";
            int iattval = 1312810155;
            double dattval = 1312810155.200420052007;

            int errct = 0;
            var s_hashTable = new Hashtable();
            Component.GetStringUserProperties(ref s_hashTable);
            foreach (var key in s_hashTable.Keys)
            {
                if (sString.Contains(key) == false)
                {
                    sString.Add(key);
                    Component.GetAttribute(key.ToString(), ref sattval);
                    //dgattibutedata.Rows.Add(key, "str", sattval);
                    AddDataGridRowData(dgattibutedata, false, key.ToString(), "str", sattval);
                    //rowitem = chklstattfield.Items[ct];
                    ct++;
                }
                else
                {
                    errct ++;
                }
            }

            var d_hashTable = new Hashtable();
            Component.GetDoubleUserProperties(ref d_hashTable);
            foreach (var key in d_hashTable.Keys)
            {
                if (sDouble.Contains(key) == false)
                {
                    sDouble.Add(key);
                    Component.GetAttribute(key.ToString(), ref dattval);
                    //dgattibutedata.Rows.Add(key.ToString(), "double", dattval);
                    AddDataGridRowData(dgattibutedata, false, key.ToString(), "double", dattval.ToString());
                    ct++;
                }
                else
                {
                    errct ++;
                }
            }
            var i_hashTable = new Hashtable();
            Component.GetIntegerUserProperties(ref i_hashTable);
            foreach (var key in i_hashTable.Keys)
            {
                if (sInt.Contains(key) == false)
                {
                    sInt.Add(key);
                    Component.GetAttribute(key.ToString(), ref iattval);
                    //dgattibutedata.Rows.Add(key.ToString(), "Int", iattval);
                    AddDataGridRowData(dgattibutedata, false, key.ToString(), "Int", iattval.ToString());
                    ct++;
                }
                else
                {
                    errct ++;
                }    
            }




            returndata.Add(sString);
            returndata.Add(sInt);
            returndata.Add(sDouble);


            return returndata;
        }




        private void AddDataGridRowData(DataGridView MyDataGridView, bool UpdateFlag, string attribute_name, string attribute_type, string attribute_value)
        {
            MyDataGridView.Rows.Add(UpdateFlag, attribute_name, attribute_type, attribute_value);
        }


        







        // Create a node sorter that implements the IComparer interface.
        public class NodeSorter : IComparer
        {
            // Compare the length of the strings, or the strings
            // themselves, if they are the same length.
            public int Compare(object x, object y)
            {
                TreeNode tx = x as TreeNode;
                TreeNode ty = y as TreeNode;

                // Compare the length of the strings, returning the difference.

                if (tx.Text.Length >= ty.Text.Length)
                    return tx.Text.Length - ty.Text.Length;

                // If they are the same length, call Compare.
                return string.Compare(tx.Text, ty.Text);
            }
        }

        private int getNodeIndex(TreeNode chkNode, string chkValue, string findtype)
        {
            if (chkNode != null)
            {
                for (int i = 0; i < chkNode.Nodes.Count; i++)
                {
                    TreeNode compchk = chkNode.Nodes[i];
                    if (findtype.Contains("Contains"))
                    {
                        if (compchk.Text.Contains(chkValue))
                        {
                            return i;
                        }
                    }
                    if (findtype.Contains("Equals"))
                    {
                        if (compchk.Text.ToUpper() == chkValue.ToUpper())
                        {
                            return i;
                        }
                    }
                    if (findtype.Contains("Exact"))
                    {
                        if (compchk.Text == chkValue)
                        {
                            return i;
                        }
                    }

                }
            }
            
            return -1;
        }
        private bool IsTreeviewNodeContains(TreeNode chkNode, string chkValue)
        {
            for (int i = 0; i < chkNode.Nodes.Count; i++)
            {
                TreeNode compchk = chkNode.Nodes[i];

                if (compchk.Text.Contains(chkValue))
                {
                    return true;
                }
            }
            return false;
        }
        private bool IsTreeviewNodeEquals(TreeNode chkNode, string chkValue)
        {
            for (int i = 0; i < chkNode.Nodes.Count; i++)
            {
                TreeNode compchk = chkNode.Nodes[i];

                if (compchk.Text == chkValue)
                {
                    return true;
                }
            }
            return false;
        }


        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        public class ComponentDetails
        {
            public string key;
            public string ComponentType;
            public string ComponentNumber;            
            public string PrimaryGUID;
            //public List<Beam> Secondaries = new List<Beam>();
            public string SecondaryCount;
            public ArrayList Secondaries = new ArrayList();
        }
        private ArrayList getConnectionList(string SelectionOption, TreeView mytreeview)
        {
            ArrayList myConnectionData = new ArrayList();

            ArrayList detail = new ArrayList();
            ArrayList comp = new ArrayList();
            ArrayList conn = new ArrayList();

            Model MyModel = new Model();
            if (MyModel.GetConnectionStatus() && MyModel.GetInfo().ModelPath != string.Empty)
            {
                ArrayList mdlObj = new ArrayList();
                mdlObj.Add(TSM.ModelObject.ModelObjectEnum.CONNECTION);
                mdlObj.Add(TSM.ModelObject.ModelObjectEnum.DETAIL);

                //Get selected objects and put them in an enumerator/container

                ModelObjectEnumerator MyEnum;

                if (SelectionOption == "Select")
                { 
                    MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
                    while (MyEnum.MoveNext())
                    {
                        TSM.Component mycoomp = MyEnum.Current as TSM.Component;
                        if (mycoomp != null)
                            comp.Add(mycoomp);

                        TSM.Detail mydetail = MyEnum.Current as TSM.Detail;
                        if (mydetail != null)
                            detail.Add(mydetail);

                        TSM.Connection myconn = MyEnum.Current as TSM.Connection;
                        if (myconn != null)
                            conn.Add(myconn);
                    }
                }

                else
                {
                    MyEnum = MyModel.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.CONNECTION);
                    while (MyEnum.MoveNext())                    
                    {


                        TSM.Connection myconn = MyEnum.Current as TSM.Connection;
                        if (myconn != null)
                            conn.Add(myconn);
                    }
                    MyEnum = MyModel.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.DETAIL);
                    while (MyEnum.MoveNext())
                    {


                        TSM.Detail mydetail = MyEnum.Current as TSM.Detail;
                        if (mydetail != null)
                            detail.Add(mydetail);


                    }

                }
                    
                
                //Cycle through selected objects
              

            }
            int comp_ct = conn.Count;
            int dtl_ct = detail.Count;
            int primarytagid = 0;
            int secondarytagid = 0;
            int secondarycounttagid = 0;
            Tekla.Structures.Model.Operations.Operation.DisplayPrompt("Connections: " + comp.Count.ToString() + " ," + dtl_ct.ToString() + " ," + comp_ct.ToString() + ".");
            mytreeview.Nodes.Clear();
            //process 
            if (comp_ct + dtl_ct >= 1)
            {
                TreeNode root = mytreeview.Nodes.Add("Connection", "Connection( " + (comp_ct + dtl_ct).ToString() + " )");
                ArrayList myconnectiondetails = new ArrayList();
                ComponentDetails mycomponentDetails = new ComponentDetails();
                root.Tag = "|A";
                int compno_ct = 1;
   
                string profile = "";
                if (comp_ct >= 1)
                {

                    TreeNode ndcomp = root.Nodes.Add("Component( " + comp_ct.ToString() + " )");
                    ndcomp.Tag = root.Tag + "|B1";
                    ndcomp.ForeColor = System.Drawing.Color.CornflowerBlue;
                    
                    TreeNode ndcompno = ndcomp.Nodes.Add("("); 
                    TreeNode ndprimary = ndcomp.Nodes.Add("(");
                    TreeNode ndsec_ct = ndprimary.Nodes.Add("("); 
                    TreeNode ndsecondary = ndsec_ct.Nodes.Add("("); 

                    ndcomp.Nodes.Clear();
                    ndcompno.Nodes.Clear();
                    ndprimary.Nodes.Clear();
                    ndsec_ct.Nodes.Clear();
                    ndsecondary.Nodes.Clear();
                    
                    string chknumber = "|";
                    foreach (TSM.Connection myconn in conn)
                    {

                        
                        bool c3 = myconn.LoadAttributesFromFile("standard");
                        int xs_componentnumber = myconn.Number;
                        string s_conn = xs_componentnumber.ToString();

                        //s_conn = s_conn;// + " [" + myconn.Name + "] ";
                        if (!chknumber.Contains("|" + s_conn + "|"))
                        {
                            chknumber = chknumber + s_conn + "|";
                            ndcompno = ndcomp.Nodes.Add(s_conn + "(0)");
                            ndcompno.Tag = ndcomp.Tag + "|C" + compno_ct.ToString();
                            compno_ct ++;//= compno_ct + 1;
                        }

                        //s_conn = s_conn + " [" + myconn.Name + "] ";
                        //if (!chknumber.Contains("|" + s_conn + "|"))
                        //{
                        //    chknumber = chknumber + s_conn + "|";
                        //    ndcompno = ndcomp.Nodes.Add(s_conn + "(0)");
                        //    ndcompno.Tag = "C" + compno_ct.ToString();
                        //    compno_ct = compno_ct + 1;
                        //}


                        int idx = getNodeIndex(ndcomp, s_conn + "(", "Contains");
                        if (idx >= 0)
                        {
                            TreeNode chkcompno = ndcomp.Nodes[idx];
                            string[] split = chkcompno.Text.Split(new Char[] { '(' });
                            int prof_ct = Convert.ToInt32(split[1].Replace(")", ""));
                            chkcompno.Text = s_conn + "(" + (prof_ct + 1).ToString() + ")";
                            //add primary
                            TSM.ModelObject priObject = myconn.GetPrimaryObject();
                            priObject.GetReportProperty("PROFILE", ref profile);
                            string priguid = "";
                            priObject.GetReportProperty("GUID", ref priguid);
                            primarytagid ++;
                            addPrimaryData(chkcompno, profile, out ndprimary, primarytagid);

                            //add secondary count         
                            ArrayList secObject = myconn.GetSecondaryObjects();
                            int secnos = secObject.Count;
                            secondarycounttagid++;//= secondarycounttagid + 1;
                            addSecondaryCount(ndprimary, "Secondary " + secnos.ToString(), out ndsec_ct, secondarycounttagid);

                            //tvwconnection.ExpandAll();




                            //mycomponentDetails.Secondaries = <Tekla.Structures.Model.Object> secObject;
                            ArrayList secBeamObject =new ArrayList();
                            secondarytagid++;//= secondarytagid + 1;
                            addSecondaryData(secObject, ndsec_ct, out ndsecondary, secondarytagid);
                            //if (ndsecondary != null)
                            //{

                            //}



                        }
                    }

                }
                if (dtl_ct >= 1)
                {
                    TreeNode ndcomp = root.Nodes.Add("Detail( " + dtl_ct.ToString() + " )");
                    ndcomp.Tag = root.Tag + "|B2";
                }

                tvwpreviewconntype1.ExpandAll();
                tvwpreviewconntype1.Tag = "ExpandAll";
                //tvwconnection.TreeViewNodeSorter();
                //tvwconnection.TreeViewNodeSorter = new NodeSorter();
            }
            ArrayList mylist = new ArrayList();
            mylist.Add(detail);
            mylist.Add(comp);
            mylist.Add(conn);


            return mylist;
        }
        private ArrayList getConnectionObjects(string SelectionOption)
        {

            ArrayList myConnectionData = new ArrayList();
            myConnectionData.Add("Esskay");
            ArrayList detail = new ArrayList();
            ArrayList comp = new ArrayList();
            ArrayList conn = new ArrayList();

            MyModel = skTSLib.ModelAccess.ConnectToModel();
            if (MyModel.GetConnectionStatus() && MyModel.GetInfo().ModelPath != string.Empty)
            {
                //Get selected objects and put them in an enumerator/container
                ModelObjectEnumerator MyEnum;
                if (SelectionOption == "Select")
                {
                    MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
                    pbar1.Maximum = pbar1.Maximum + MyEnum.GetSize();
                    while (MyEnum.MoveNext())
                    {

                        TSM.Connection myconn = MyEnum.Current as TSM.Connection;
                        if (myconn != null)
                            conn.Add(myconn);

                        TSM.Detail mydetail = MyEnum.Current as TSM.Detail;
                        if (mydetail != null)
                            detail.Add(mydetail);

                        TSM.Component mycoomp = MyEnum.Current as TSM.Component;
                        if (mycoomp != null)
                            comp.Add(mycoomp);

                        pbar1.Value = pbar1.Value + 1;
                        pbar1.Refresh();
                    }
                }

                else
                {
                    MyEnum = MyModel.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.CONNECTION);
                    pbar1.Maximum = pbar1.Maximum + MyEnum.GetSize();
                    //pbar1.Value = 0;
                    while (MyEnum.MoveNext())
                    {

                        TSM.Connection myconn = MyEnum.Current as TSM.Connection;
                        if (myconn != null)
                            conn.Add(myconn);

                        pbar1.Value = pbar1.Value + 1;
                        pbar1.Refresh();
                    }
                    MyEnum = MyModel.GetModelObjectSelector().GetAllObjectsWithType(TSM.ModelObject.ModelObjectEnum.DETAIL);
                    pbar1.Maximum = pbar1.Maximum + MyEnum.GetSize();
                    //pbar1.Value = 0;
                    while (MyEnum.MoveNext())
                    {
                        TSM.Detail mydetail = MyEnum.Current as TSM.Detail;
                        if (mydetail != null)
                            detail.Add(mydetail);

                        pbar1.Value = pbar1.Value + 1;
                        pbar1.Refresh();
                    }

                }
            }
            myConnectionData.Add(conn);
            myConnectionData.Add(detail);
            myConnectionData.Add(comp);
            return myConnectionData;

        }

        private ArrayList getConnectionInputs(TSM.Connection myConnection, string myConnectionType)
        {
            ArrayList myreturnvalue = new ArrayList();
            myreturnvalue.Add("Esskay");

            string priProfile = string.Empty;
            string priName = string.Empty;
            string priPhase = string.Empty;
            string priGuid = string.Empty;
            int secCount = 0;
            string secProfile = string.Empty;
            string secName = string.Empty;
            string secPhase = string.Empty;
            string secGuid = string.Empty;
            string secProfileList = "|";
            string secNameList = "|";
            string secPhaseList = "|";
            string secGuidList = "|";

            string connNumber = string.Empty;
            string connCode = string.Empty;
            string connGuid = string.Empty;


            //connType = "Connection";

            connNumber = myConnection.Number.ToString();
            //myconn.GetAttribute("joint_code", ref connCode);
            connCode = myConnection.Code;
            connGuid = myConnection.Identifier.GUID.ToString();
            TSM.ModelObject priObject = myConnection.GetPrimaryObject();
            priObject.GetReportProperty("PROFILE", ref priProfile);
            priObject.GetReportProperty("NAME", ref priName);
            priObject.GetReportProperty("PHASE.NAME", ref priPhase);
            priObject.GetReportProperty("GUID", ref priGuid);
            //add secondary count         
            ArrayList secObjectList = myConnection.GetSecondaryObjects();
            secCount = secObjectList.Count;
            foreach (TSM.ModelObject secObject in secObjectList)
            {
                secObject.GetReportProperty("PROFILE", ref secProfile);
                secObject.GetReportProperty("NAME", ref secName);
                secObject.GetReportProperty("PHASE.NAME", ref secPhase);
                secObject.GetReportProperty("GUID", ref secGuid);
                secProfileList = secProfileList + "|" + secProfile;
                secNameList = secNameList + "|" + secName;
                secPhaseList = secPhaseList + "|" + secPhase;
                secGuidList = secGuidList + "|" + secGuid;
            }

            myreturnvalue.Add(priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));
            myreturnvalue.Add(secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", "") + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid);
            myreturnvalue.Add(connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));

            return myreturnvalue;
        }

        private ArrayList getDetailInputs(TSM.Detail myConnection, string myConnectionType)
        {
            ArrayList myreturnvalue = new ArrayList();
            myreturnvalue.Add("Esskay");

            string priProfile = string.Empty;
            string priName = string.Empty;
            string priPhase = string.Empty;
            string priGuid = string.Empty;
            int secCount = 0;
            string secProfile = string.Empty;
            string secName = string.Empty;
            string secPhase = string.Empty;
            string secGuid = string.Empty;
            string secProfileList = "|";
            string secNameList = "|";
            string secPhaseList = "|";
            string secGuidList = "|";

            string connNumber = string.Empty;
            string connCode = string.Empty;
            string connGuid = string.Empty;


            //connType = "Connection";

            connNumber = myConnection.Number.ToString();
            //myconn.GetAttribute("joint_code", ref connCode);
            connCode = myConnection.Code;
            connGuid = myConnection.Identifier.GUID.ToString();
            TSM.ModelObject priObject = myConnection.GetPrimaryObject();
            priObject.GetReportProperty("PROFILE", ref priProfile);
            priObject.GetReportProperty("NAME", ref priName);
            priObject.GetReportProperty("PHASE.NAME", ref priPhase);
            priObject.GetReportProperty("GUID", ref priGuid);
            //add secondary count         
            //ArrayList secObjectList = myConnection.GetSecondaryObjects();
            //secCount = secObjectList.Count;
            //foreach (TSM.ModelObject secObject in secObjectList)
            //{
            //    secObject.GetReportProperty("PROFILE", ref secProfile);
            //    secObject.GetReportProperty("NAME", ref secName);
            //    secObject.GetReportProperty("PHASE.NAME", ref secPhase);
            //    secObject.GetReportProperty("GUID", ref secGuid);
            //    secProfileList = secProfileList + "|" + secProfile;
            //    secNameList = secNameList + "|" + secName;
            //    secPhaseList = secPhaseList + "|" + secPhase;
            //    secGuidList = secGuidList + "|" + secGuid;
            //}

            myreturnvalue.Add(priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));
            myreturnvalue.Add(secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", "") + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid);
            myreturnvalue.Add(connNumber + "|" + myConnectionType + "|" + connCode + "|" + connGuid + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));

            return myreturnvalue;
        }

        private ArrayList PreviewSelectedConnection(ArrayList mySelectedObjects, TreeView tvwconnection)
        {
            ArrayList myConnectionData = new ArrayList();
            ArrayList myConnectionData_primary = new ArrayList();
            ArrayList myConnectionData_secondary = new ArrayList();


            ArrayList conn = new ArrayList();
            ArrayList detail = new ArrayList();
            ArrayList comp = new ArrayList();
            if (mySelectedObjects.Count >= 4)
            {
                conn = mySelectedObjects[1] as ArrayList;
                detail = mySelectedObjects[2] as ArrayList;
                comp = mySelectedObjects[3] as ArrayList;
            }

            int conn_ct = conn.Count;
            int dtl_ct = detail.Count;
            int comp_ct = comp.Count;

            pbar1.Maximum = pbar1.Maximum + conn_ct + dtl_ct + comp_ct;
            //pbar1.Value = 0;



            Tekla.Structures.Model.Operations.Operation.DisplayPrompt("Filtered Connections: " + conn_ct.ToString() + " ," + dtl_ct.ToString() + " ," + comp_ct.ToString() + ".");
            tvwconnection.Nodes.Clear();
            //process 
            if (conn_ct + dtl_ct + comp_ct >= 1)
            {
                ArrayList myconnectiondetails = new ArrayList();
                ArrayList mygetConnectionInputs = new ArrayList();
                ComponentDetails mycomponentDetails = new ComponentDetails();

                foreach (TSM.Connection myconnection in conn)
                {
                    mygetConnectionInputs = getConnectionInputs(myconnection, "Connection");
                    if (mygetConnectionInputs.Count == 4)
                    {
                        myConnectionData_primary.Add(mygetConnectionInputs[1]);
                        myConnectionData_secondary.Add(mygetConnectionInputs[2]);
                        myConnectionData.Add(mygetConnectionInputs[3]);
                    }

                }
                foreach (TSM.Detail mydetail in detail)
                {
                    mygetConnectionInputs = getDetailInputs(mydetail, "Detail");
                    if (mygetConnectionInputs.Count == 4)
                    {
                        myConnectionData_primary.Add(mygetConnectionInputs[1]);
                        myConnectionData_secondary.Add(mygetConnectionInputs[2]);
                        myConnectionData.Add(mygetConnectionInputs[3]);
                    }

                }
                foreach (TSM.Connection mycomponent in comp)
                {
                    mygetConnectionInputs = getConnectionInputs(mycomponent, "Connection");
                    if (mygetConnectionInputs.Count == 4)
                    {
                        myConnectionData_primary.Add(mygetConnectionInputs[1]);
                        myConnectionData_secondary.Add(mygetConnectionInputs[2]);
                        myConnectionData.Add(mygetConnectionInputs[3]);
                    }

                }

                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
            }
            myConnectionData_primary.Sort();
            myConnectionData_secondary.Sort();
            myConnectionData.Sort();
            //process tree view
            //tvwpreviewconn.Nodes.Clear();
            tvwconnection.Nodes.Clear();
            if (myConnectionData_secondary.Count >= 1)
            {
                ArrayList unqConnNumber = new ArrayList();
                ArrayList tvwnodeid = new ArrayList();
                TreeNode ndroot = tvwconnection.Nodes.Add("Connection", "Connection( " + (myConnectionData.Count).ToString() + " )");
                //TreeNode ndconnNumber = ndroot.Nodes.Add("(");
                //TreeNode ndconnCode = ndconnNumber.Nodes.Add("(");
                //TreeNode ndsecondaryCount = ndconnNumber.Nodes.Add("(");
                //TreeNode ndsecondary = ndsecondaryCount.Nodes.Add("(");
                //TreeNode ndprimary = ndsecondary.Nodes.Add("(");
                //TreeNode ndguid = ndprimary.Nodes.Add("(");

                TreeNode ndconnNumber = null;
                TreeNode ndconnCode = null;
                TreeNode ndsecondaryCount = null;
                TreeNode ndsecondary = null;
                TreeNode ndsecondaryname = null;
                TreeNode ndprimary = null;
                TreeNode ndprimaryname = null;
                TreeNode ndguid = null;

                ndroot.Tag = "|A";
                //string connNumber = string.Empty;
                //string connCode = string.Empty;
                //string connGuid = string.Empty;
                //string connType = string.Empty;

                foreach (string conndata in myConnectionData)
                {
                    string[] split = conndata.Split(new Char[] { '|' });
                    string chknode = string.Empty;
                    //connNumber + "|" + connType + "|" + connCode + "|" + connGuid + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", "")
                    if (split.Length >= 13)
                    {
                        //add connection number
                        string connNumber = split[0];
                        string connType = split[1];
                        string connCode = split[2];
                        string connGuid = split[3];
                        string priProfile = split[4];
                        string priName = split[5];
                        string priPhase = split[6];
                        string priGuid = split[7];
                        string secCount = split[8];
                        string secProfileList = split[9];
                        string secNameList = split[10];
                        string secPhaseList = split[11];
                        string secGuidList = split[12];



                        ndconnNumber = UpdateTreeviewData(ndroot, "", connNumber);
                        ndconnCode = UpdateTreeviewData(ndconnNumber, "", connCode);
                        if (Convert.ToInt32(secCount) >= 2)
                        {
                            ndsecondaryCount = UpdateTreeviewData(ndconnCode, "", secCount);
                            ndsecondary = UpdateTreeviewData(ndsecondaryCount, "", secProfileList + "[" + secNameList + "]");
                        }
                        else
                        {
                            ndsecondary = UpdateTreeviewData(ndconnCode, "", secProfileList + "[" + secNameList + "]");
                        }

                        ndprimary = UpdateTreeviewData(ndsecondary, "", priProfile + "[" + priName + "]");
                        ndguid = UpdateTreeviewData(ndprimary, "", "guid:" + connGuid);

                    }

                }
            }

            return myConnectionData;
        }

        

        //private ArrayList getConnectionData(string SelectionOption, TreeView tvwconnection)
        //{            
        //    ArrayList myConnectionData = new ArrayList();
        //    ArrayList myConnectionData_primary = new ArrayList();
        //    ArrayList myConnectionData_secondary = new ArrayList();

        //    ArrayList myconnobj = getConnectionObjects(SelectionOption);
        //    ArrayList conn = new ArrayList();
        //    ArrayList detail = new ArrayList();
        //    ArrayList comp = new ArrayList();
        //    if (myconnobj.Count >=4)
        //    {
        //        conn = myconnobj[1] as ArrayList;
        //        detail = myconnobj[2] as ArrayList;
        //        comp = myconnobj[3] as ArrayList;
        //    }

        //    int conn_ct = conn.Count;
        //    int dtl_ct = detail.Count;
        //    int comp_ct = comp.Count;


        //    Tekla.Structures.Model.Operations.Operation.DisplayPrompt("Filtered Connections: " + conn_ct.ToString() + " ," + dtl_ct.ToString() + " ," + comp_ct.ToString() + ".");
        //    tvwconnection.Nodes.Clear();
        //    //process 
        //    if (conn_ct + dtl_ct + comp_ct  >= 1)
        //    {

                    

        //        ArrayList myconnectiondetails = new ArrayList();
        //        ComponentDetails mycomponentDetails = new ComponentDetails();

        //        if (conn_ct >= 1)
        //        {
        //            string connType = string.Empty;
        //            connType = "Connection";
        //            foreach (TSM.Connection myconn in conn)
        //            {
        //                string priProfile = string.Empty;
        //                string priName = string.Empty;
        //                string priPhase = string.Empty;
        //                string priGuid = string.Empty;
        //                int secCount = 0;
        //                string secProfile = string.Empty;
        //                string secName = string.Empty;
        //                string secPhase = string.Empty;
        //                string secGuid = string.Empty;
        //                string secProfileList = "|";
        //                string secNameList = "|";
        //                string secPhaseList = "|";
        //                string secGuidList = "|";

        //                string connNumber = string.Empty;
        //                string connCode = string.Empty;
        //                string connGuid = string.Empty;

        //                connNumber = myconn.Number.ToString();
        //                //myconn.GetAttribute("joint_code", ref connCode);
        //                connCode = myconn.Code;
        //                connGuid = myconn.Identifier.GUID.ToString();
        //                TSM.ModelObject priObject = myconn.GetPrimaryObject();
        //                priObject.GetReportProperty("PROFILE", ref priProfile);
        //                priObject.GetReportProperty("NAME", ref priName);
        //                priObject.GetReportProperty("PHASE.NAME", ref priPhase);
        //                priObject.GetReportProperty("GUID", ref priGuid);
        //                //add secondary count         
        //                ArrayList secObjectList = myconn.GetSecondaryObjects();
        //                secCount = secObjectList.Count;
        //                foreach (TSM.ModelObject secObject in secObjectList)
        //                {
        //                    secObject.GetReportProperty("PROFILE", ref secProfile);
        //                    secObject.GetReportProperty("NAME", ref secName);
        //                    secObject.GetReportProperty("PHASE.NAME", ref secPhase);
        //                    secObject.GetReportProperty("GUID", ref secGuid);
        //                    secProfileList = secProfileList + "|" + secProfile;
        //                    secNameList = secNameList + "|" + secName;
        //                    secPhaseList = secPhaseList + "|" + secPhase;
        //                    secGuidList = secGuidList + "|" + secGuid;
        //                }
        //                myConnectionData_primary.Add(priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + connType + "|" + connCode + "|" + connGuid  + "|" + secCount + "|" + secProfileList.Replace("||","") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));
        //                myConnectionData_secondary.Add(secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", "") + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + connNumber + "|" + connType + "|" + connCode + "|" + connGuid );
        //                myConnectionData.Add(connNumber + "|" + connType + "|" + connCode + "|" + connGuid + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", ""));


        //            }

        //        }

        //    }
        //    myConnectionData_primary.Sort();
        //    myConnectionData_secondary.Sort();
        //    myConnectionData.Sort();
        //    //process tree view
        //    tvwpreviewconn.Nodes.Clear();
        //    tvwconnection.Nodes.Clear();
        //    if (myConnectionData_secondary.Count >= 1)
        //    {
        //        ArrayList unqConnNumber = new ArrayList();
        //        ArrayList tvwnodeid = new ArrayList();
        //        TreeNode ndroot = tvwpreviewconn.Nodes.Add("Connection", "Connection( " + (myConnectionData.Count).ToString() + " )");
        //        //TreeNode ndconnNumber = ndroot.Nodes.Add("(");
        //        //TreeNode ndconnCode = ndconnNumber.Nodes.Add("(");
        //        //TreeNode ndsecondaryCount = ndconnNumber.Nodes.Add("(");
        //        //TreeNode ndsecondary = ndsecondaryCount.Nodes.Add("(");
        //        //TreeNode ndprimary = ndsecondary.Nodes.Add("(");
        //        //TreeNode ndguid = ndprimary.Nodes.Add("(");

        //        TreeNode ndconnNumber = null;
        //        TreeNode ndconnCode = null;
        //        TreeNode ndsecondaryCount = null;
        //        TreeNode ndsecondary = null;
        //        TreeNode ndsecondaryname = null;
        //        TreeNode ndprimary = null;
        //        TreeNode ndprimaryname = null;
        //        TreeNode ndguid = null;

        //        ndroot.Tag = "|A";
        //        //string connNumber = string.Empty;
        //        //string connCode = string.Empty;
        //        //string connGuid = string.Empty;
        //        //string connType = string.Empty;

        //        foreach (string conndata in myConnectionData)
        //        {
        //            string[] split = conndata.Split(new Char[] { '|' });
        //            string chknode = string.Empty;
        //            //connNumber + "|" + connType + "|" + connCode + "|" + connGuid + "|" + priProfile + "|" + priName + "|" + priPhase + "|" + priGuid + "|" + secCount + "|" + secProfileList.Replace("||", "") + "|" + secNameList.Replace("||", "") + "|" + secPhaseList.Replace("||", "") + "|" + secGuidList.Replace("||", "")
        //            if (split.Length >= 13)
        //            {
        //                //add connection number
        //                string connNumber = split[0];
        //                string connType = split[1];
        //                string connCode = split[2];
        //                string connGuid = split[3];
        //                string priProfile = split[4];
        //                string priName = split[5];
        //                string priPhase = split[6];
        //                string priGuid = split[7];
        //                string secCount = split[8];
        //                string secProfileList = split[9];
        //                string secNameList = split[10];
        //                string secPhaseList = split[11];
        //                string secGuidList = split[12];



        //                ndconnNumber = UpdateTreeviewData(ndroot, "", connNumber);
        //                ndconnCode = UpdateTreeviewData(ndconnNumber, "", connCode);
        //                if (Convert.ToInt32(secCount) >=2)
        //                {
        //                    ndsecondaryCount  = UpdateTreeviewData(ndconnCode, "", secCount);
        //                    ndsecondary  = UpdateTreeviewData(ndsecondaryCount, "", secProfileList + "[" + secNameList + "]");
        //                }
        //                else
        //                {
        //                    ndsecondary = UpdateTreeviewData(ndconnCode, "", secProfileList + "[" + secNameList + "]");
        //                }

        //                ndprimary  = UpdateTreeviewData(ndsecondary, "", priProfile + "[" + priName  + "]");
        //                ndguid  = UpdateTreeviewData(ndprimary, "", "guid:" + connGuid);

        //            }

        //        }
        //    }
                
        //    return myConnectionData;
        //}

        private TreeNode UpdateTreeviewData(TreeNode parentnodename, string nodetag, string appendstring)
        {
            TreeNode mynode = null;
            int idx = getNodeIndex(parentnodename, appendstring, "Equals");
            if (idx >= 0)
                mynode = parentnodename.Nodes[idx];
            else
            {
                mynode = parentnodename.Nodes.Add(appendstring);
                mynode.Tag = nodetag;
            }
            return mynode;
        }
        private void addSecondaryData(ArrayList secObject, TreeNode ndprimary,out TreeNode myndsecondary, int tag)
        {
            TreeNode ndsecondary = null;
            foreach (Tekla.Structures.Model.Object mybeam in secObject)
            {

                Beam chkBeam = mybeam as Beam;
                if (chkBeam != null)
                {
                    string secProfile = string.Empty;
                    chkBeam.GetReportProperty("PROFILE", ref secProfile);
                    //ndsec_ct = ndprimary.Nodes.Add(secProfile);
                    if (IsTreeviewNodeContains(ndprimary, secProfile + "(") == false)
                    {
                        ndsecondary = ndprimary.Nodes.Add(secProfile + "(1)");
                        ndsecondary.Tag = ndprimary.Tag + "|F" + tag.ToString();
                        //mycomponentDetails.Secondaries = secBeamObject;
                        //mycomponentDetails.key = root.Tag.ToString() + ndcomp.Tag.ToString() + ndsecondary.Tag.ToString();
                        //mycomponentDetails.ComponentType = "Connection";
                        //mycomponentDetails.ComponentNumber = xs_componentnumber.ToString();
                        //mycomponentDetails.PrimaryGUID = priguid;
                        //secBeamObject.Add(chkBeam);
                    }
                    else
                    {
                        int idx = getNodeIndex(ndprimary,  secProfile + "(", "Contains");
                        if (idx >= 0)
                        {
                            ndsecondary = ndprimary.Nodes[idx];
                            string[] split = ndsecondary.Text.Split(new Char[] { '(' });
                            int prof_ct = Convert.ToInt32(split[1].Replace(")", ""));
                            ndsecondary.Text = secProfile + "(" + (prof_ct + 1).ToString() + ")";

                        }

                    }
                    ndsecondary.ForeColor = System.Drawing.Color.Blue;


                }
            }
            myndsecondary = ndsecondary;
        }

        private void addPrimaryData(TreeNode chkcompno,string profile, out TreeNode myndprimary, int tag)
        {
            TreeNode ndprimary = null;
            if (IsTreeviewNodeContains(chkcompno, "Primary: " + profile + "(") == false)
            {
                ndprimary = chkcompno.Nodes.Add("Primary: " + profile + "(1)");
                ndprimary.Tag = chkcompno.Tag + "|D" + tag.ToString();
            }

            else
            {
                int idx = getNodeIndex(chkcompno, "Primary: " + profile + "(", "Contains");
                if (idx >= 0)
                {
                    ndprimary = chkcompno.Nodes[idx];
                    string[] split = ndprimary.Text.Split(new Char[] { '(' });
                    int prof_ct = Convert.ToInt32(split[1].Replace(")", ""));
                    ndprimary.Text = "Primary: " + profile + "(" + (prof_ct + 1).ToString() + ")";
                }

            }
            myndprimary = ndprimary;
        }

        private void addSecondaryCount(TreeNode chkcompno, string SecondaryCount, out TreeNode myndprimary, int tag)
        {
            TreeNode ndprimary = null;
            if (IsTreeviewNodeContains(chkcompno, SecondaryCount + "(") == false)
            {
                ndprimary = chkcompno.Nodes.Add(SecondaryCount + "(1)");
                ndprimary.Tag = chkcompno.Tag + "|E" + tag.ToString();
            }

            else
            {
                int idx = getNodeIndex(chkcompno, SecondaryCount + "(", "Contains");
                if (idx >= 0)
                {
                    ndprimary = chkcompno.Nodes[idx];
                    string[] split = ndprimary.Text.Split(new Char[] { '(' });
                    int prof_ct = Convert.ToInt32(split[1].Replace(")", ""));
                    ndprimary.Text = SecondaryCount + "(" + (prof_ct + 1).ToString() + ")";
                }

            }
            myndprimary = ndprimary;
        }


       

        private void btntvwtoggle_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Toggle", "");
            if (tvwpreviewconntype1.Tag.ToString() == "ExpandAll")
            {
                tvwpreviewconntype1.CollapseAll();
                tvwpreviewconntype1.Tag = "CollapseAll";
            }
            else
            {
                tvwpreviewconntype1.ExpandAll();
                tvwpreviewconntype1.Tag = "ExpandAll";
            }

            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            lblsbar1.Text = "Pls. Wait, A36 to A529-50 change in progress..";
            TSM.Operations.Operation.DisplayPrompt(lblsbar1.Text);
            rtxtgradechange.Text = "Report\n";
            ArrayList connchange = new ArrayList();
            ModelObjectEnumerator MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
            string s_chk1 = "";
            string s_chk2 = "";
            int ct = 0;
            int tot = MyEnum.GetSize();
            while (MyEnum.MoveNext())
            {
                bool flag = false;
                string connlog = "";
                TSM.Connection ThisConnection = MyEnum.Current as TSM.Connection;
                if (ThisConnection != null)
                {
                    //"L" angle profile change grade from A36 TO A529-50 (i.e begins with "L"
                    s_chk1 = "";
                    s_chk2 = "";
                    //check1 -143
                    s_chk1 = skTSLib.getComponent(ThisConnection, "lprof");
                    if (s_chk1 != "")
                    {
                        s_chk2 = skTSLib.getComponent(ThisConnection, "mat");
                        if (s_chk2 != "")
                        {
                            if ((s_chk1.Substring(0, 1).ToUpper() == "L") && (s_chk2.ToUpper() == "A36"))
                            {
                                skTSLib.setComponent(ThisConnection, "mat", "A529-50");
                                flag = true;
                                connlog = connlog + "\n" + s_chk1 + " " + s_chk2 + ">>A529-50";
                            }

                        }
                    }
                    s_chk1 = "";
                    s_chk2 = "";
                    //check2 -143
                    s_chk1 = skTSLib.getComponent(ThisConnection, "prof");
                    if (s_chk1 != "")
                    {
                        
                        s_chk2 = skTSLib.getComponent(ThisConnection, "mat2");
                        if (s_chk2 != "")
                        {
                            if ((s_chk1.Substring(0, 1).ToUpper() == "L") && (s_chk2.ToUpper() == "A36"))
                            {
                                skTSLib.setComponent(ThisConnection, "mat2", "A529-50");
                                flag = true;
                                connlog = connlog + "\n" + s_chk1 + " " + s_chk2 + ">>A529-50";
                            }
                        }
                    }
                    s_chk1 = "";
                    s_chk2 = "";
                    //check3 -143
                    s_chk1 = skTSLib.getComponent(ThisConnection, "lcon");
                    if (s_chk1 != "")
                    {
                        
                        s_chk2 = skTSLib.getComponent(ThisConnection, "mat3");
                        if (s_chk2 != "")
                        {
                            if ((s_chk1.Substring(0, 1).ToUpper() == "L") && (s_chk2.ToUpper() == "A36"))
                            {
                                skTSLib.setComponent(ThisConnection, "mat3", "A529-50");
                                flag = true;
                                connlog = connlog + "\n" + s_chk1 + " " + s_chk2 + ">>A529-50";
                            }
                        }
                    }
                    s_chk1 = "";
                    s_chk2 = "";
                    //check4 -143
                    s_chk1 = skTSLib.getComponent(ThisConnection, "bcon");
                    if (s_chk1 != "")
                    {
                        
                        s_chk2 = skTSLib.getComponent(ThisConnection, "mat4");
                        if (s_chk2 != "")
                        {
                            if ((s_chk1.Substring(0, 1).ToUpper() == "L") && (s_chk2.ToUpper() == "A36"))
                            {
                                skTSLib.setComponent(ThisConnection, "mat4", "A529-50");
                                flag = true;
                                connlog = connlog + "\n" + s_chk1 + " " + s_chk2 + ">>A529-50";
                            }
                        }
                    }
                    if (flag == true)
                    {
                        connchange.Add(ThisConnection);
                        rtxtgradechange.Text = rtxtgradechange.Text + "\n\n" + ThisConnection.Number.ToString() + " id:" + ThisConnection.Identifier.ID.ToString()  +  connlog;
                        ThisConnection.Modify();
                        ct = ct + 1;

                    }
                    
                    //MyModel.CommitChanges();
                    //return;
                }


            }
            //ObjectsToSelect.Add(MyBeam);

            var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MOS.Select(connchange);

            skWinLib.worklog(skApplicationName, skApplicationVersion, "A36 to A529-50 for LProfile", ";Total: " + tot.ToString() + "; UpdatedCount: " + ct.ToString());
            lblsbar1.Text = "Ready." + ct.ToString() + " A36 to A529-50 changed.";
            TSM.Operations.Operation.DisplayPrompt(lblsbar1.Text);
            MessageBox.Show(ct.ToString() + " connection A36 grade changed to  A529-50 out of total " + tot.ToString());

        }

            private ArrayList getAttributeData_FromSelectedConnectionObject(ArrayList ConnectionObjects)
            {
                ArrayList ConnectionData = new ArrayList();
                //process selected Connection Objects
                ArrayList mySelectedObjects = new ArrayList();
                ArrayList myfilteredConnection = new ArrayList();
                ArrayList myfilteredDetail = new ArrayList();
                ArrayList myfilteredComponent = new ArrayList();
                if (ConnectionObjects.Count == 4)
                {
                    myfilteredConnection = ConnectionObjects[1] as ArrayList;
                    myfilteredDetail = ConnectionObjects[2] as ArrayList;
                    myfilteredComponent = ConnectionObjects[3] as ArrayList;
                }

                int conn_ct = myfilteredConnection.Count;
                int dtl_ct = myfilteredDetail.Count;
                int comp_ct = myfilteredComponent.Count;
                pbar1.Maximum = pbar1.Maximum + conn_ct + dtl_ct + comp_ct;
                //pbar1.Value = 0;
                mySelectedObjects = myfilteredConnection;

                foreach (var detail in myfilteredDetail)
                {
                    mySelectedObjects.Add(detail);
                }
                foreach (var comp in myfilteredComponent)
                {
                    mySelectedObjects.Add(comp);
                }



                string connnumber = string.Empty;
                string Checkconnnumber = string.Empty;
                ConnectionData.Add("Esskay");
                int unqid = 0;

                foreach (Tekla.Structures.Model.ModelObject MOE in mySelectedObjects)
                {
                    Tekla.Structures.Model.ModelObject ModelSideObject = null;
                    ArrayList ObjectsToSelect = new ArrayList();
                    bool flag = false;
                    string connguid = string.Empty;
                    string connname = string.Empty;

                    string conntype = string.Empty;
                    string exportfilename = string.Empty;
                    string[] attributedata = null;
                    ArrayList MyData = new ArrayList();
                    TSM.ModelObject PrimaryObject = null;
                    ArrayList SecondaryObjects = null;


                    string primaryassypos = string.Empty;
                    string primaryprofile = string.Empty;
                    string primaryname = string.Empty;

                    int secondarycount = 0;
                    string secondaryassypos = "+";
                    string secondaryprofile = "+";
                    string secondaryname = "+";
                    string connmodtime = string.Empty;

                    Tekla.Structures.Model.Component MyCompo = MOE as Tekla.Structures.Model.Component;
                    if (MyCompo != null)
                    {
                        flag = MyCompo.Select();
                        ModelSideObject = new Model().SelectModelObject(MyCompo.Identifier);
                        connguid = MyCompo.Identifier.GUID.ToString();
                        connname = MyCompo.Name.ToString();
                        connnumber = MyCompo.Number.ToString();
                        conntype = "ail_display_component_dialog";
                        exportfilename = "ESSKAY_COM_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
                        unqid = unqid + 1;
                        exportfilename = "ESSKAY_COM_" + connnumber + "_" + unqid.ToString();
                        //exportfilename = "SK_COM_" + connnumber + "_" + connguid.ToString();
                    }
                    else
                    {
                        Tekla.Structures.Model.Connection MyConn = MOE as Tekla.Structures.Model.Connection;
                        if (MyConn != null)
                        {
                            //MyConn.Select();                       
                            flag = MyConn.Select();
                            ModelSideObject = new Model().SelectModelObject(MyConn.Identifier);
                            PrimaryObject = MyConn.GetPrimaryObject();
                            SecondaryObjects = MyConn.GetSecondaryObjects();
                            connguid = MyConn.Identifier.GUID.ToString();
                            connname = MyConn.Name.ToString();
                            connnumber = MyConn.Number.ToString();
                            if (MyConn.ModificationTime != null)
                                connmodtime = MyConn.ModificationTime.Value.ToString("dd_MM_yy_H_mm_ss");
                            else
                                connmodtime = DateTime.Now.ToString("dd_MM_yy_H_mm_ss");
                            conntype = "ail_display_connection_dialog";
                            exportfilename = "ESSKAY_CON_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
                            unqid = unqid + 1;
                            exportfilename = "ESSKAY_CON_" + connnumber + "_" + unqid.ToString();
                            exportfilename = "SK_CON_" + connnumber + "_" + connguid.Replace("-", "") + "_" + connmodtime; ;
                        }
                        else
                        {
                            Tekla.Structures.Model.Detail MyDet = MOE as Tekla.Structures.Model.Detail;
                            if (MyDet != null)
                            {
                                flag = MyDet.Select();
                                ModelSideObject = new Model().SelectModelObject(MyDet.Identifier);
                                PrimaryObject = MyDet.GetPrimaryObject();
                                connguid = MyDet.Identifier.GUID.ToString();
                                connname = MyDet.Name.ToString();
                                connnumber = MyDet.Number.ToString();
                                conntype = "ail_display_connection_dialog";
                                exportfilename = "ESSKAY_DET_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
                                unqid = unqid + 1;
                                exportfilename = "ESSKAY_DET_" + connnumber + "_" + unqid.ToString();
                                //exportfilename = "SK_DET_" + connnumber + "_" + connguid.ToString();
                            }
                        }
                    }
                    exportfilename = exportfilename.Replace("-", "");
                    if (exportfilename != string.Empty)
                    {
                        if (Checkconnnumber == connnumber || Checkconnnumber == string.Empty)
                        {
                            Checkconnnumber = connnumber;

                            if (PrimaryObject != null)
                            {
                                PrimaryObject.GetReportProperty("ASSEMBLY_POS", ref primaryassypos);
                                PrimaryObject.GetReportProperty("PROFILE", ref primaryprofile);
                                PrimaryObject.GetReportProperty("NAME", ref primaryname);

                            }
                            if (SecondaryObjects != null)
                            {

                                secondarycount = Convert.ToInt32(SecondaryObjects.Count.ToString());

                                foreach (TSM.ModelObject secObject in SecondaryObjects)
                                {
                                    string assypos = string.Empty;
                                    string profile = string.Empty;
                                    string name = string.Empty;
                                    secObject.GetReportProperty("ASSEMBLY_POS", ref assypos);
                                    secObject.GetReportProperty("PROFILE", ref profile);
                                    secObject.GetReportProperty("NAME", ref name);
                                    if (assypos.Length >= 1)
                                    {
                                        if (secondaryassypos.IndexOf(assypos) == -1)
                                        {
                                            secondaryassypos = secondaryassypos + "+" + assypos;
                                        }
                                    }
                                    if (profile.Length >= 1)
                                    {
                                        if (secondaryprofile.IndexOf(profile) == -1)
                                        {
                                            secondaryprofile = secondaryprofile + "+" + profile;
                                        }
                                    }
                                    if (name.Length >= 1)
                                    {
                                        if (secondaryname.IndexOf(name) == -1)
                                        {
                                            secondaryname = secondaryname + "+" + name;
                                        }
                                    }
                                }
                                secondaryassypos = secondaryassypos.Replace("++", "");
                                secondaryprofile = secondaryprofile.Replace("++", "");
                                secondaryname = secondaryname.Replace("++", "");

                                //for (int sec=0; sec < secondarycount;sec++)
                                //{

                                //}
                            }
                            if (modelpath.Length == 0)
                            {
                                TSM.Model model = skTSLib.ModelAccess.ConnectToModel();
                                modelpath = model.GetInfo().ModelPath;
                            }
                            //attribute file name
                            string attfilename = modelpath + "\\attributes\\" + exportfilename + ".j" + connnumber;
                            if (!File.Exists(attfilename) || chkexist.Checked == false)
                            {
                                var akit = new MacroBuilder();
                                TeklaStructures.Connect();
                                // system connections or details
                                //TSM.Operations.Operation.RunMacro("ZoomSelected.cs");

                                ObjectsToSelect.Add(ModelSideObject);
                                var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                                MOS.Select(ObjectsToSelect);
                                //TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
                                Tekla.Structures.ModelInternal.Operation.dotStartCommand(conntype, connnumber);
                                //if (ModelSideObject.Identifier.IsValid())                        
                                //{
                                //    ModelSideObject.Select();
                                //    TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
                                //}
                                //akit.ValueChange("joint_" + connnumber.ToString(), "get_menu", "standard");
                                //akit.PushButton("load", "joint_" + connnumber.ToString());
                                //akit.ValueChange("joint_" + connnumber.ToString(), "get_menu", "< Defaults >");
                                //akit.PushButton("load", "joint_" + connnumber.ToString());
                                //akit.PushButton("get_button", "joint_" + connnumber.ToString());
                                if (connnumber == "-1")
                                    akit.PushButton("get_button", connname + "_1");
                                else
                                    akit.PushButton("get_button", "joint_" + connnumber.ToString());
                                //akit.ValueChange("joint_" + connnumber.ToString(), "joint_code", exportfilename);
                                //akit.ValueChange("joint_" + connnumber.ToString(), "joint_code", exportfilename);
                                if (connnumber == "-1")
                                    akit.ValueChange(connname + "_1", "saveas_file", exportfilename);
                                else
                                    akit.ValueChange("joint_" + connnumber.ToString(), "saveas_file", exportfilename);
                                if (connnumber == "-1")
                                    akit.PushButton("saveas", connname + "_1");
                                else
                                    akit.PushButton("saveas", "joint_" + connnumber.ToString());
                                akit.PushButton("CANCEL_button", "joint_" + connnumber.ToString());
                                akit.Run();

                                //akit.PushButton("get_button", "Rail Fitting-Elbow_1");
                                //akit.ValueChange("Rail Fitting-Elbow_1", "saveas_file", "133");
                                //akit.PushButton("saveas", "Rail Fitting-Elbow_1");
                                //get the attribute file         
                            }


                            if (File.Exists(attfilename))
                            {
                                ArrayList PriSecData = new ArrayList();
                                PriSecData.Add(connguid);
                                PriSecData.Add("");
                                PriSecData.Add(primaryassypos);
                                PriSecData.Add(primaryprofile);
                                PriSecData.Add(primaryname);
                                PriSecData.Add(secondarycount);
                                PriSecData.Add(secondaryassypos);
                                PriSecData.Add(secondaryprofile);
                                PriSecData.Add(secondaryname);

                                PriSecData.Add("");
                                MyData.Add(PriSecData);
                                dgvconnectioncompare.Tag = PriSecData.Count;
                                //process attribute file
                                attributedata = File.ReadAllLines(attfilename);
                                MyData.Add(attributedata);
                                MyData.Add(MOE);
                                ConnectionData.Add(MyData);
                                //File.Delete(attfilename);
                            }
                            else
                            {
                                MessageBox.Show("Critical Error. \n" + attfilename);
                            }

                        }


                    }
                    pbar1.Value = pbar1.Value + 1;
                    pbar1.Refresh();
            }
            return ConnectionData;

        }

        //private ArrayList getAttributeData_FromSelectedConnectionObject(ArrayList mySelectedObjects)
        //{
        //    ArrayList myfilteredConnection = new ArrayList();
        //    ArrayList myfilteredDetail = new ArrayList();
        //    ArrayList myfilteredComponent = new ArrayList();
        //    if (mySelectedObjects.Count >= 4)
        //    {
        //        myfilteredConnection = mySelectedObjects[1] as ArrayList;
        //        myfilteredDetail = mySelectedObjects[2] as ArrayList;
        //        myfilteredComponent = mySelectedObjects[3] as ArrayList;
        //    }

        //    int conn_ct = myfilteredConnection.Count;
        //    int dtl_ct = myfilteredDetail.Count;
        //    int comp_ct = myfilteredComponent.Count;


        //    //MOE.SelectInstances = false;
        //    ArrayList ConnectionData = new ArrayList();
        //    string connnumber = string.Empty;
        //    string Checkconnnumber = string.Empty;
        //    ConnectionData.Add("Esskay");
        //    int unqid = 0;

        //    foreach (Tekla.Structures.Model.ModelObject MOE in myfilteredConnection)
        //    {
        //        Tekla.Structures.Model.ModelObject ModelSideObject = null;
        //        ArrayList ObjectsToSelect = new ArrayList();
        //        bool flag = false;
        //        string connguid = string.Empty;
        //        string connname = string.Empty;

        //        string conntype = string.Empty;
        //        string exportfilename = string.Empty;
        //        string[] attributedata = null;
        //        ArrayList MyData = new ArrayList();
        //        TSM.ModelObject PrimaryObject = null;
        //        ArrayList SecondaryObjects = null;


        //        string primaryassypos = string.Empty;
        //        string primaryprofile = string.Empty;
        //        string primaryname = string.Empty;

        //        int secondarycount = 0;
        //        string secondaryassypos = "+";
        //        string secondaryprofile = "+";
        //        string secondaryname = "+";
        //        string connmodtime = string.Empty;





        //        Tekla.Structures.Model.Component MyCompo = MOE as Tekla.Structures.Model.Component;
        //        if (MyCompo != null)
        //        {
        //            flag = MyCompo.Select();
        //            ModelSideObject = new Model().SelectModelObject(MyCompo.Identifier);
        //            connguid = MyCompo.Identifier.GUID.ToString();
        //            connname = MyCompo.Name.ToString();
        //            connnumber = MyCompo.Number.ToString();
        //            conntype = "ail_display_component_dialog";
        //            exportfilename = "ESSKAY_COM_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
        //            unqid = unqid + 1;
        //            exportfilename = "ESSKAY_COM_" + connnumber + "_" + unqid.ToString();
        //            //exportfilename = "SK_COM_" + connnumber + "_" + connguid.ToString();
        //        }
        //        else
        //        {
        //            Tekla.Structures.Model.Connection MyConn = MOE as Tekla.Structures.Model.Connection;
        //            if (MyConn != null)
        //            {
        //                //MyConn.Select();                       
        //                flag = MyConn.Select();
        //                ModelSideObject = new Model().SelectModelObject(MyConn.Identifier);
        //                PrimaryObject = MyConn.GetPrimaryObject();
        //                SecondaryObjects = MyConn.GetSecondaryObjects();
        //                connguid = MyConn.Identifier.GUID.ToString();
        //                connname = MyConn.Name.ToString();
        //                connnumber = MyConn.Number.ToString();
        //                if (MyConn.ModificationTime != null)
        //                    connmodtime = MyConn.ModificationTime.Value.ToString("dd_MM_yy_H_mm_ss");
        //                else
        //                    connmodtime = DateTime.Now.ToString("dd_MM_yy_H_mm_ss");
        //                conntype = "ail_display_connection_dialog";
        //                exportfilename = "ESSKAY_CON_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
        //                unqid = unqid + 1;
        //                exportfilename = "ESSKAY_CON_" + connnumber + "_" + unqid.ToString();
        //                exportfilename = "SK_CON_" + connnumber + "_" + connguid.Replace("-", "") + "_" + connmodtime; ;
        //            }
        //            else
        //            {
        //                Tekla.Structures.Model.Detail MyDet = MOE as Tekla.Structures.Model.Detail;
        //                if (MyDet != null)
        //                {
        //                    flag = MyDet.Select();
        //                    ModelSideObject = new Model().SelectModelObject(MyDet.Identifier);
        //                    PrimaryObject = MyDet.GetPrimaryObject();
        //                    connguid = MyDet.Identifier.GUID.ToString();
        //                    connname = MyDet.Name.ToString();
        //                    connnumber = MyDet.Number.ToString();
        //                    conntype = "ail_display_connection_dialog";
        //                    exportfilename = "ESSKAY_DET_" + connnumber + "_" + DateTime.Now.ToString("ddMMMyyHHmmssfff");
        //                    unqid = unqid + 1;
        //                    exportfilename = "ESSKAY_DET_" + connnumber + "_" + unqid.ToString();
        //                    //exportfilename = "SK_DET_" + connnumber + "_" + connguid.ToString();
        //                }
        //            }
        //        }
        //        exportfilename = exportfilename.Replace("-", "");
        //        if (exportfilename != string.Empty)
        //        {
        //            if (Checkconnnumber == connnumber || Checkconnnumber == string.Empty)
        //            {
        //                Checkconnnumber = connnumber;

        //                if (PrimaryObject != null)
        //                {
        //                    PrimaryObject.GetReportProperty("ASSEMBLY_POS", ref primaryassypos);
        //                    PrimaryObject.GetReportProperty("PROFILE", ref primaryprofile);
        //                    PrimaryObject.GetReportProperty("NAME", ref primaryname);

        //                }
        //                if (SecondaryObjects != null)
        //                {

        //                    secondarycount = Convert.ToInt32(SecondaryObjects.Count.ToString());

        //                    foreach (TSM.ModelObject secObject in SecondaryObjects)
        //                    {
        //                        string assypos = string.Empty;
        //                        string profile = string.Empty;
        //                        string name = string.Empty;
        //                        secObject.GetReportProperty("ASSEMBLY_POS", ref assypos);
        //                        secObject.GetReportProperty("PROFILE", ref profile);
        //                        secObject.GetReportProperty("NAME", ref name);
        //                        if (assypos.Length >= 1)
        //                        {
        //                            if (secondaryassypos.IndexOf(assypos) == -1)
        //                            {
        //                                secondaryassypos = secondaryassypos + "+" + assypos;
        //                            }
        //                        }
        //                        if (profile.Length >= 1)
        //                        {
        //                            if (secondaryprofile.IndexOf(profile) == -1)
        //                            {
        //                                secondaryprofile = secondaryprofile + "+" + profile;
        //                            }
        //                        }
        //                        if (name.Length >= 1)
        //                        {
        //                            if (secondaryname.IndexOf(name) == -1)
        //                            {
        //                                secondaryname = secondaryname + "+" + name;
        //                            }
        //                        }
        //                    }
        //                    secondaryassypos = secondaryassypos.Replace("++", "");
        //                    secondaryprofile = secondaryprofile.Replace("++", "");
        //                    secondaryname = secondaryname.Replace("++", "");

        //                    //for (int sec=0; sec < secondarycount;sec++)
        //                    //{

        //                    //}
        //                }
        //                if (modelpath.Length == 0)
        //                {
        //                    TSM.Model model = skTSLib.ModelAccess.ConnectToModel();
        //                    modelpath = model.GetInfo().ModelPath;
        //                }
        //                //attribute file name
        //                string attfilename = modelpath + "\\attributes\\" + exportfilename + ".j" + connnumber;
        //                if (!File.Exists(attfilename) || chkexist.Checked == true)
        //                {
        //                    var akit = new MacroBuilder();
        //                    TeklaStructures.Connect();
        //                    // system connections or details
        //                    //TSM.Operations.Operation.RunMacro("ZoomSelected.cs");

        //                    ObjectsToSelect.Add(ModelSideObject);
        //                    var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
        //                    MOS.Select(ObjectsToSelect);
        //                    //TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
        //                    Tekla.Structures.ModelInternal.Operation.dotStartCommand(conntype, connnumber);
        //                    //if (ModelSideObject.Identifier.IsValid())                        
        //                    //{
        //                    //    ModelSideObject.Select();
        //                    //    TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
        //                    //}
        //                    //akit.ValueChange("joint_" + connnumber.ToString(), "get_menu", "standard");
        //                    //akit.PushButton("load", "joint_" + connnumber.ToString());
        //                    //akit.ValueChange("joint_" + connnumber.ToString(), "get_menu", "< Defaults >");
        //                    //akit.PushButton("load", "joint_" + connnumber.ToString());
        //                    //akit.PushButton("get_button", "joint_" + connnumber.ToString());
        //                    if (connnumber == "-1")
        //                        akit.PushButton("get_button", connname + "_1");
        //                    else
        //                        akit.PushButton("get_button", "joint_" + connnumber.ToString());
        //                    //akit.ValueChange("joint_" + connnumber.ToString(), "joint_code", exportfilename);
        //                    //akit.ValueChange("joint_" + connnumber.ToString(), "joint_code", exportfilename);
        //                    if (connnumber == "-1")
        //                        akit.ValueChange(connname + "_1", "saveas_file", exportfilename);
        //                    else
        //                        akit.ValueChange("joint_" + connnumber.ToString(), "saveas_file", exportfilename);
        //                    if (connnumber == "-1")
        //                        akit.PushButton("saveas", connname + "_1");
        //                    else
        //                        akit.PushButton("saveas", "joint_" + connnumber.ToString());
        //                    akit.PushButton("CANCEL_button", "joint_" + connnumber.ToString());
        //                    akit.Run();

        //                    //akit.PushButton("get_button", "Rail Fitting-Elbow_1");
        //                    //akit.ValueChange("Rail Fitting-Elbow_1", "saveas_file", "133");
        //                    //akit.PushButton("saveas", "Rail Fitting-Elbow_1");
        //                    //get the attribute file         
        //                }


        //                if (File.Exists(attfilename))
        //                {
        //                    ArrayList PriSecData = new ArrayList();
        //                    PriSecData.Add(connguid);
        //                    PriSecData.Add("");
        //                    PriSecData.Add(primaryassypos);
        //                    PriSecData.Add(primaryprofile);
        //                    PriSecData.Add(primaryname);
        //                    PriSecData.Add(secondarycount);
        //                    PriSecData.Add(secondaryassypos);
        //                    PriSecData.Add(secondaryprofile);
        //                    PriSecData.Add(secondaryname);

        //                    PriSecData.Add("");
        //                    MyData.Add(PriSecData);
        //                    dgvconnectioncompare.Tag = PriSecData.Count;
        //                    //process attribute file
        //                    attributedata = File.ReadAllLines(attfilename);
        //                    MyData.Add(attributedata);
        //                    MyData.Add(MOE);
        //                    ConnectionData.Add(MyData);
        //                    //File.Delete(attfilename);
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Critical Error. \n" + attfilename);
        //                }

        //            }


        //        }

        //    }





        //    return ConnectionData;

        //}

        private void PreviewAttributeData(ArrayList skconnection)
        {

            lblsbar1.Text = "Please Wait...";
            //get connection from tekla          
            //TSM.ModelObjectEnumerator MOE = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
            //lblsbar1.Text = "Please Wait, <" + MOE.GetSize().ToString() + "> Processing connection....";
            //ArrayList skconnection = getAttributeData_FromSelectedConnectionObject( MOE);
            ArrayList ObjectsToSelect = new ArrayList();

            dgvconnectioncompare.Rows.Clear();
            dgvconnectioncompare.Columns.Clear();
            dgvconnectioncompare.Columns.Add("SKConnection", "SKConnection");


            bool rowheader = false;
            int colid = 0;


            pbar1.Maximum = pbar1.Maximum + skconnection.Count;
            //pbar1.Value = 0;
            for (int i = 1; i < skconnection.Count; i++)
            {

                ArrayList myconndata = skconnection[i] as ArrayList;
                ArrayList prisecdata = myconndata[0] as ArrayList;
                string connguid = prisecdata[0].ToString();
                string[] attributedata = myconndata[1] as string[];
                TSM.ModelObject myObject = myconndata[2] as TSM.ModelObject;
                ObjectsToSelect.Add(myObject);
                //add first empty column

                colid = colid + 1;
                dgvconnectioncompare.Columns.Add("Value", colid.ToString());
                dgvconnectioncompare.Columns[colid].FillWeight = 1;
                dgvconnectioncompare.Columns[colid].SortMode = DataGridViewColumnSortMode.NotSortable;
                for (int a = 0; a < prisecdata.Count; a++)
                {
                    if (rowheader == false)
                    {
                        dgvconnectioncompare.Rows.Add("Field");
                        dgvconnectioncompare.Rows[a].Cells[1].Value = prisecdata[a].ToString();
                        dgvconnectioncompare.Rows[dgvconnectioncompare.Rows.Count - 1].HeaderCell.Value = (dgvconnectioncompare.Rows.Count).ToString();
                    }
                    else
                    {
                        string fieldvalue = prisecdata[a].ToString();
                        fieldvalue = fieldvalue.Replace("2147483648", "");
                        fieldvalue = fieldvalue.Replace("-.000000", "-");
                        fieldvalue = fieldvalue.Replace(".000000", "");
                        dgvconnectioncompare.Rows[a].Cells[colid].Value = fieldvalue;

                    }
                }


                for (int j = 2; j < attributedata.Length; j++)
                {
                    string attdata = attributedata[j].ToString();
                    int pos1 = attdata.ToUpper().IndexOf("_ATTRIBUTES.");
                    if (pos1 >= 0)
                    {
                        string field = attdata.Substring(pos1 + 12);
                        pos1 = field.IndexOf(" ");
                        string fieldvalue = field.Substring(pos1).Trim();
                        field = field.Substring(0, pos1);
                        fieldvalue = fieldvalue.Replace("2147483648", "");
                        fieldvalue = fieldvalue.Replace("-.000000", "-");
                        fieldvalue = fieldvalue.Replace(".000000", "");
                        if (rowheader == false)
                        {



                            dgvconnectioncompare.Rows.Add(field);
                            dgvconnectioncompare.Rows[j + prisecdata.Count - 2].Cells[1].Value = fieldvalue;
                            dgvconnectioncompare.Rows[dgvconnectioncompare.Rows.Count - 1].HeaderCell.Value = (dgvconnectioncompare.Rows.Count).ToString();
                        }
                        else
                        {
                            dgvconnectioncompare.Rows[j + prisecdata.Count - 2].Cells[colid].Value = fieldvalue;
                        }
                    }


                }
                rowheader = true;
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
            }
            //save temporary file

            string filename = modelpath + "\\SK_MASSCONNT_" + skWinLib.username + ".sknum";
            Export_DataGridView(dgvconnectioncompare, filename, "PIE", "");

            var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MOS.Select(ObjectsToSelect);

            UpdateSKFieldName();
            updatecolumnheader();
            lblsbar1.Text = "Ready";
        }
        private void btngroup_Click(object sender, EventArgs e)
        {
        }
        

        private void tvwconnection_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode mynode = tvwpreviewconntype1.SelectedNode;
            if (mynode != null)
            {
                label3.Text = mynode.Tag.ToString();
                //mynode.BackColor = System.Drawing.Color.LightCyan;
            }

        }

        private void tvwconnection_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode mynode = tvwpreviewconntype1.SelectedNode;
            if (mynode != null)
            {
                label3.Text = mynode.Tag.ToString();
                //mynode.BackColor = System.Drawing.Color.Black;
            }


        }

        private void txtmm_TextChanged(object sender, EventArgs e)
        {
            if (flagimperial == false)
            {
                flagmetric = true;
                if (txtmm.Text.Trim().Length >= 1)
                {

                    string metric = txtmm.Text;
                    if (skWinLib.IsNumeric(metric) == true)
                    {
                        double d_metric = Convert.ToDouble(metric);
                        //d_metric = d_metric / 25.4;
                        TSDT.Distance t_metric = new TSDT.Distance(d_metric);


                        txtimpinch.Text = t_metric.ToFractionalInchesString(CultureInfo.InvariantCulture);
                        txtimpfeet.Text = t_metric.ToFractionalFeetAndInchesString(CultureInfo.InvariantCulture);
                        //txtmm.Text = (d_metric * 10).ToString();
                    }
                }
            }

            flagmetric = false;
        }

        private void txtimpinch_TextChanged(object sender, EventArgs e)
        {
            if (flagmetric == false)
            {
                flagimperial = true;
                if (txtimpinch.Text.Trim().Length >= 1)
                {
                    try
                    {
                        TSDT.Distance d_imperial = TSDT.Distance.FromFractionalFeetAndInchesString(txtimpinch.Text, CultureInfo.InvariantCulture, TSDT.Distance.UnitType.Inch);
                        txtmm.Text = d_imperial.Millimeters.ToString();
                        //txtmm.Text = (Convert.ToDouble(txtmm.Text) * 10).ToString();
                        txtimpfeet.Text = d_imperial.ToFractionalFeetAndInchesString();
                    }
                    catch
                    {

                    }

                }
            }
            flagimperial = false;
        }

        //private void dgvconnectioncompare_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{

        //}


            private void updatecolumnheader()
        {
            //provide column header title
            if (dgvconnectioncompare.Columns.Count >= 1)
            {
                dgvconnectioncompare.Columns[0].HeaderCell.Value = "SKGrouping";


                for (int j = 1; j < dgvconnectioncompare.Columns.Count; j++)
                {
                    int groupqty = 0;
                    if (dgvconnectioncompare.Rows[0].Cells[j].Value != null)
                    {
                        string guid = dgvconnectioncompare.Rows[0].Cells[j].Value.ToString();
                        groupqty = guid.Length - guid.Replace("+", "").Length;
                        groupqty = groupqty + 1;
                        dgvconnectioncompare.Rows[1].Cells[j].Value = groupqty.ToString();
                    }
                    else
                        dgvconnectioncompare.Rows[1].Cells[j].Value = "";


                    dgvconnectioncompare.Columns[j].HeaderCell.Value = j.ToString();
                }
            }
        }
        private void UpdateSKFieldName()
        {
            int startrowid = Convert.ToInt32(dgvconnectioncompare.Tag.ToString());
            if (dgvconnectioncompare.Tag.ToString() == "")
                startrowid = 9;

            if (dgvconnectioncompare.Rows.Count >= startrowid)
            {
                dgvconnectioncompare.Rows[0].Cells[0].Value = "guid";
                dgvconnectioncompare.Rows[1].Cells[0].Value = "Group_Qty";
                dgvconnectioncompare.Rows[2].Cells[0].Value = "P_Mark";
                dgvconnectioncompare.Rows[3].Cells[0].Value = "P_Profile";
                dgvconnectioncompare.Rows[4].Cells[0].Value = "P_Name";
                dgvconnectioncompare.Rows[5].Cells[0].Value = "S_Qty";
                dgvconnectioncompare.Rows[6].Cells[0].Value = "S_Mark";
                dgvconnectioncompare.Rows[7].Cells[0].Value = "S_Profile";
                dgvconnectioncompare.Rows[8].Cells[0].Value = "S_Name";
                dgvconnectioncompare.Rows[9].Cells[0].Value = "SK";
            }           

        }

        private void groupAttributeData()
        {
            if (dgvconnectioncompare.Rows.Count >=1)
            {
                groupCount = 0;
                errorCount = 0;


                ArrayList colgroupingdata = new ArrayList();

                //string splitstr = "|SPLIT|";
                string splitstr = " + ";
                lblsbar1.Text = "Processing for Grouping...";
                int startrowid = 9;
                if (dgvconnectioncompare.Tag.ToString() != "" && dgvconnectioncompare.Tag.ToString() != null)
                    startrowid = Convert.ToInt32(dgvconnectioncompare.Tag.ToString());
                ////get the field values top most row's
                //foreach(DataGridViewRow myrow in dgvconnectioncompare.Rows)
                //{
                //    if (myrow.Cells[0].Value.ToString().ToUpper() == "FIELD")
                //        startrowid = startrowid + 1;
                //}



                bool loopconncode = chkIgnoreConnectionCode.Checked;
                //delete identical row values
                pbar1.Maximum = pbar1.Maximum + (dgvconnectioncompare.Rows.Count - startrowid) + dgvconnectioncompare.Columns.Count;
                //pbar1.Value = 0;
                for (int j = startrowid; j < dgvconnectioncompare.Rows.Count; j++)
                {
                    DataGridViewRow myrow = dgvconnectioncompare.Rows[j];
                    if (myrow.Cells.Count >= 3)
                    {


                        bool chkflag = false;
                        for (int i = 2; i < myrow.Cells.Count; i++)
                        {
                            if (myrow.Cells[1].FormattedValue.ToString() != myrow.Cells[i].FormattedValue.ToString())
                            {
                                chkflag = true;
                                break; // get out of the loop
                            }
                        }
                        if (chkflag == false)
                        {
                            dgvconnectioncompare.Rows.RemoveAt(j);
                            j = j - 1;
                        }
                        if (loopconncode == true)
                        {
                            if (myrow.Index != -1)
                            {
                                string chkcellval = myrow.Cells[0].FormattedValue.ToString();
                                if (chkcellval.ToUpper().IndexOf("JOINT_CODE") >= 0)
                                {
                                    dgvconnectioncompare.Rows.RemoveAt(j);
                                    j = j - 1;
                                    loopconncode = false;
                                }
                            }

                        }
                    }

                    pbar1.Value = pbar1.Value + 1;
                    pbar1.Refresh();
                }
                //dgvconnectioncompare.Refresh();
                //add all column values and store in arraylist
                for (int a = 1; a < dgvconnectioncompare.Columns.Count; a++)
                {
                    string cumcoldata = string.Empty;
                    for (int j = startrowid; j < dgvconnectioncompare.Rows.Count; j++)
                    {
                        DataGridViewRow myrow = dgvconnectioncompare.Rows[j];
                        //2147483648 replace with default
                        cumcoldata = cumcoldata + "|" + myrow.Cells[a].Value.ToString();
                    }
                    //cumcoldata = cumcoldata.Replace("2147483648", "") + splitstr + a.ToString();
                    cumcoldata = cumcoldata + splitstr + a.ToString();
                    colgroupingdata.Add(cumcoldata);
                    pbar1.Value = pbar1.Value + 1;
                    pbar1.Refresh();
                }

                //sort for checking identical
                colgroupingdata.Sort();
                pbar1.Maximum = pbar1.Maximum + colgroupingdata.Count + 5;
                //pbar1.Value = 0;


                //return;
                //grouping of column
                for (int j = 0; j < colgroupingdata.Count; j++)
                {

                    if (colgroupingdata[j].ToString().IndexOf(splitstr) >= 0)
                    {
                        string chkdata = colgroupingdata[j].ToString();
                        chkdata = chkdata.Substring(0, chkdata.IndexOf(splitstr));
                        int chkcolid = 1;
                        if (chkdata != "")
                            chkcolid = Convert.ToInt32(colgroupingdata[j].ToString().Replace(chkdata, "").Replace(splitstr, ""));
                        else
                            chkcolid = Convert.ToInt32(colgroupingdata[j].ToString().Replace(splitstr, ""));
                        for (int i = j + 1; i < colgroupingdata.Count; i++)
                        {
                            string mydata = colgroupingdata[i].ToString();
                            if (mydata != "")
                            {
                                mydata = mydata.Substring(0, mydata.IndexOf(splitstr));
                                int mycolid = 1;
                                if (mydata != "")
                                    mycolid = Convert.ToInt32(colgroupingdata[i].ToString().Replace(mydata, "").Replace(splitstr, ""));
                                else
                                    mycolid = Convert.ToInt32(colgroupingdata[i].ToString().Replace(splitstr, ""));


                                if (chkdata == mydata)
                                {

                                    for (int a = 0; a < startrowid; a++)
                                    {
                                        DataGridViewRow myrow = dgvconnectioncompare.Rows[a];
                                        if (myrow.Cells[chkcolid].Value != null && myrow.Cells[mycolid].Value != null)
                                        {
                                            //check weather already exists
                                            string chkstr1 = myrow.Cells[chkcolid].Value.ToString();
                                            string chkstr2 = myrow.Cells[mycolid].Value.ToString();
                                            if (chkstr1.IndexOf(chkstr2) == -1)
                                                myrow.Cells[chkcolid].Value = chkstr1 + splitstr + chkstr2;
                                        }

                                    }
                                    //set all row value as blank for future deletion
                                    DataGridViewRow mydeleterow = dgvconnectioncompare.Rows[0];
                                    mydeleterow.Cells[mycolid].Value = "Group with " + chkcolid.ToString();
                                    //dgvconnectioncompare.Refresh();
                                    colgroupingdata[i] = "";
                                }
                            }

                        }
                    }
                    pbar1.Value = pbar1.Value + 1;
                    pbar1.Refresh();

                }


                //delete grouped columns
                foreach (DataGridViewRow myrow in dgvconnectioncompare.Rows)
                {
                    for (int i = 1; i < dgvconnectioncompare.Columns.Count; i++)
                    {
                        string chkgroup = myrow.Cells[i].Value.ToString();
                        if (chkgroup.IndexOf("Group with") >= 0)
                        {
                            //delete the column
                            dgvconnectioncompare.Columns.RemoveAt(i);
                            groupCount = groupCount + 1;
                            i = i - 1;

                        }
                    }

                }
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();

                //

                //UpdateSKFieldName();
                skWinLib.updaterowheader(dgvconnectioncompare);
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                updatecolumnheader();
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                HighlightRowDifferenceQty();
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                //GetValue("ADVANCED_OPTION.XS_IMPERIAL") == TRUE
                if (XS_UNIT == "TRUE")
                    UpdateDataGrid_Imperial(dgvconnectioncompare);
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
            }
            
        }
        private void btngrouping_Click(object sender, EventArgs e)
        {
            pbar1.Visible = true;
            DateTime s_tm = DateTime.Now;
            //lblguid.Text = "";
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "GroupConnection", "", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);

            //Esskay Grouping Confidential
            groupAttributeData();

            TimeSpan span = DateTime.Now.Subtract(s_tm);
            //lblsbar1.Text = span.Seconds.ToString() + "Ready. " + groupCount.ToString() + " / " + (dgvconnectioncompare.Columns.Count - 1).ToString();
            lblsbar1.Text = span.Minutes.ToString() + ":" + span.Seconds.ToString() + " Ready. " + (dgvconnectioncompare.Columns.Count - 1).ToString() + " Group(s) in " + groupCount.ToString();
            pbar1.Visible = false;
        }


        private void HighlightRowDifferenceQty()
        {
            int startrowid = 9;

            if (dgvconnectioncompare.Tag.ToString() != "" && dgvconnectioncompare.Tag.ToString() != null)
                startrowid = Convert.ToInt32(dgvconnectioncompare.Tag.ToString());

           
            if (dgvconnectioncompare.Columns.Count >=4)
            {
                //delete identical row values
                for (int j = startrowid; j < dgvconnectioncompare.Rows.Count; j++)
                {
                    DataGridViewRow myrow = dgvconnectioncompare.Rows[j];
                    ArrayList unqcellvalue = new ArrayList();
                    ArrayList unqcellid = new ArrayList();
                    for (int i = 1; i < myrow.Cells.Count; i++)
                    {
                        string cellvalue = myrow.Cells[i].FormattedValue.ToString();
                        int idx = unqcellvalue.IndexOf(cellvalue);
                        if (idx == -1)
                        {
                            unqcellid.Add("1|" + i.ToString());
                            unqcellvalue.Add(cellvalue);
                        }
                        else
                        {

                            string existvalue = unqcellid[idx] as string;
                            string[] split = existvalue.Split(new Char[] { '|' });
                            unqcellid[idx] = (Convert.ToInt32(split[0]) + 1).ToString() + "|" + split[1] + "," + i.ToString();  
                        }

                    }
                    if (unqcellvalue.Count >= 1)
                    {

                        string fieldname = dgvconnectioncompare.Rows[j].Cells[0].Value.ToString();
                        if (fieldname.IndexOf("[") >=0)
                        {
                            fieldname = fieldname.Substring(0, fieldname.IndexOf("["));
                        }
                        fieldname = fieldname + "[" + unqcellvalue.Count.ToString() + "]";
                        dgvconnectioncompare.Rows[j].Cells[0].Value = fieldname;
                        //highlight if difference count = 1 only
                        if (unqcellid.Count ==2)
                        {
                            unqcellid.Sort();
                            string existvalue = unqcellid[0] as string;
                            string[] split = existvalue.Split(new Char[] { '|' });
                            if (split[1].IndexOf(",") == -1)
                            {
                                int colid = Convert.ToInt32(split[1]);
                                myrow.Cells[colid].Style.BackColor = System.Drawing.Color.LightGoldenrodYellow;
                            }
                            else
                            {
                                string[] splitcolid = split[1].Split(new Char[] { ',' });
                                for (int colid=0;colid < splitcolid.Length; colid++)
                                {
                                    myrow.Cells[Convert.ToInt32(splitcolid[colid])].Style.BackColor = System.Drawing.Color.LightGoldenrodYellow;
                                }
                            }
                            
         
                        }

                    }
                }
            }

        }


        private void UpdateDataGrid_Imperial(DataGridView dgvconnectioncompare)
        {
            ArrayList ignoreconverion = new ArrayList();
            ignoreconverion.Add("startno_");
            ignoreconverion.Add("mat");
            ignoreconverion.Add("prefix");


            int startrowid = 9;
            for (int i = 0; i < dgvconnectioncompare.Columns.Count; i++)
            {
                for (int j = startrowid; j < dgvconnectioncompare.Rows.Count; j++)
                {
                    DataGridViewRow myrow = dgvconnectioncompare.Rows[j];
                    if (myrow.Cells[i].FormattedValue != null)
                    {
                        string cellvalue = myrow.Cells[i].FormattedValue.ToString();
                        for(int ig=0;ig< ignoreconverion.Count;ig++)
                        {
                            string chkignore = ignoreconverion[ig].ToString();
                            if (cellvalue.ToUpper().Contains(chkignore.ToUpper()))
                                goto UpdateDataGrid_ImperialskipJ;
 
                        }

  


                        if (skWinLib.IsNumeric(cellvalue) ==true)
                        {
                            UpdateInteger_Imperial(cellvalue, myrow, i);
                        }
                        else
                        {
                            if (cellvalue.IndexOf("*") >= 0)
                            {
                                char c = (char)34;
                                if (cellvalue.IndexOf(c.ToString()) >= 0)
                                {
                                    cellvalue = cellvalue.Replace(c.ToString(), "").Trim();
                                    // The default unit is millimeter.
                                    TSDT.Distance.CurrentUnitType = TSDT.Distance.UnitType.Millimeter;

                                    TSDT.DistanceList distanceInmm = TSDT.DistanceList.Parse(cellvalue, CultureInfo.InvariantCulture, TSDT.Distance.UnitType.Millimeter);

                                    // We can override the unit type to display the distances in inches.
                                    string defaultFormatInInches = distanceInmm.ToString(null, null, TSDT.Distance.UnitType.Inch);
                                    myrow.Cells[i].Value = defaultFormatInInches;
                                }

                            }
                            else
                            {
                                //check for lwd = "76.5" and change to 3 if exists
                                char[] mychar = cellvalue.ToCharArray();
                                char c = (char)34;
                                if (cellvalue.IndexOf(c.ToString()) >= 0)
                                {
                                    cellvalue = cellvalue.Replace(c.ToString(), "").Trim();
                                    if (skWinLib.IsNumeric(cellvalue) == true)
                                    {
                                        UpdateInteger_Imperial(cellvalue, myrow, i);
                                    }
                                }
                            }
                        }
                    }
UpdateDataGrid_ImperialskipJ:;
                }
            }
        }

        private void UpdateInteger_Imperial(string cellvalue, DataGridViewRow myrow, int index)
        {
            //check whether integer and less than 12
            try
            {
                int chklessthan12 = Convert.ToInt32(cellvalue);
                if (chklessthan12.ToString().Trim() == cellvalue.Trim())
                {
                    //less than 12 will be mostly dropdownlist hence done 
                    if (chklessthan12 < 12)
                        return;
                }
            }
            catch
            {

            }
            TSDT.Distance mydist = TSDT.Distance.FromDecimalString(cellvalue);
            myrow.Cells[index].Value = mydist.ToFractionalFeetAndInchesString();

        }
            

        private void btnsave_Click(object sender, EventArgs e)
        {

            //process for save

            DateTime s_tm = DateTime.Now;
            pbar1.Visible = true;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "SaveConnection", "", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
            //if (tabControl1.SelectedIndex == 0)
            //    //SaveMassConnection();
            //else if (tabControl1.SelectedIndex == 1)
            //    //SaveMassConnection();
            //else
            if (tabControl1.SelectedIndex == 0)
                SaveMassConnection();

            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = "Ready " + span.Minutes.ToString() + ":" + span.Seconds.ToString();
            pbar1.Visible = false;
        }


        private void Export_DataGridView(DataGridView mydatagridview, string filename, string myFormat, string myHideColumns)
        {
            int rowct = mydatagridview.Rows.Count;
            int colct = mydatagridview.Columns.Count;
            ArrayList dgv = new ArrayList();
            string fileseperator = ",";

            if (myFormat == "CSV")
                fileseperator = ",";
            else if (myFormat == "PIE")
                fileseperator = "|";
            for (int i = 0; i < rowct; i++)
            {
                DataGridViewRow row = mydatagridview.Rows[i];
                string dgvexportdata = fileseperator;
                for (int j = 0; j < colct; j++)
                {
                    DataGridViewCell ox = mydatagridview.Rows[i].Cells[j];
                    if (j == 0)
                        dgvexportdata = ox.EditedFormattedValue.ToString();
                    else
                        dgvexportdata = dgvexportdata + fileseperator + ox.EditedFormattedValue.ToString();
                }
                if (dgvexportdata.Replace(fileseperator, "").Length >= 1)
                    dgv.Add(dgvexportdata);
            }
            //write to file
            using (TextWriter streamWriter = new StreamWriter(filename, false))
            {
                foreach (string linedata in dgv)
                    streamWriter.WriteLine(linedata);
                streamWriter.Close();

            }

        }

        private void Import_DataGridView(DataGridView mydatagridview, string filename, string myFormat, string myHideColumns)
        {
            DateTime s_tm = DateTime.Now;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            //read to file
            ArrayList dgv = new ArrayList();
            System.IO.StreamReader srdgv = new System.IO.StreamReader(filename, System.Text.Encoding.Default);
            string line = " ";
            while ((line = srdgv.ReadLine()) != null)
            {
                dgv.Add(line);
            }
            srdgv.Close();

            char[] splitchar = new Char[] { ',' };
            if (myFormat == "CSV")
                splitchar = new Char[] { ',' };
            else if (myFormat == "PIE")
                splitchar = new Char[] { '|' };

            //update data grid view
            mydatagridview.Rows.Clear();
           
            foreach (string linedata in dgv)
            {
                string[] splitrowdata = linedata.ToString().Split(splitchar);
                int colct = splitrowdata.Length - mydatagridview.Columns.Count;
                for (int i=0;i< colct;i++)
                {
                    mydatagridview.Columns.Add((i + 1).ToString(), (i + 1).ToString());
                    mydatagridview.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    mydatagridview.Columns[i].FillWeight = 1;

                }
                mydatagridview.Rows.Add(splitrowdata);
                mydatagridview.Rows[mydatagridview.Rows.Count - 1].HeaderCell.Value = (mydatagridview.Rows.Count).ToString();
            }


            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = span.Seconds.ToString() + " Ready.";
        }

        private void SaveMassConnection()
        {
            //setStartup(skApplicationName, skApplicationVersion, "Save", ";Model:" + modelname, "Please Wait!!!,....", "btnSave_Click");
            DateTime s_tm = DateTime.Now;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            string filename = modelpath + "\\SK_MASSCONNL_" + skWinLib.username + ".sknum";
            Export_DataGridView(dgvconnectioncompare, filename, "PIE", "");

            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = span.Seconds.ToString() + " Ready.";
            //setCompletion(skApplicationName, skApplicationVersion, "", "", "Ready " + processtime.ToString(@"mm\:ss"), "");

        }




        private void SelectConnection()
        {

            //string splitstr = "|SPLIT|";
            string splitstr = " + ";
            // Iterate through the SelectedCells collection and sum up the values.
            ArrayList ObjectsToSelect = new ArrayList();
            for (int counter = 0;  counter < (dgvconnectioncompare.SelectedCells.Count); counter++)
            {
                if (dgvconnectioncompare.SelectedCells[counter].FormattedValueType == Type.GetType("System.String"))
                {
                    string guidlist = dgvconnectioncompare.SelectedCells[counter].FormattedValue.ToString();
                    Tekla.Structures.Model.ModelObject ModelSideObject = null;

                    Identifier myguid = null;
                    if (guidlist.IndexOf(splitstr) >= 0)
                    {
                        //string[] guidsplit = guidlist.Replace(splitstr, "|").Split(new Char[] { '|' });
                        string[] guidsplit = guidlist.Split(new Char[] { '+' });

                        for (int i = 0; i < guidsplit.Length; i++)
                        {
                            myguid = new Identifier(guidsplit[i]);
                            if (myguid != null)
                            {
                                ModelSideObject = new Model().SelectModelObject(myguid);
                                ObjectsToSelect.Add(ModelSideObject);
                            }

                        }
                    }
                    else
                    {
                        myguid = new Identifier(guidlist);
                        if (myguid != null)
                        {
                            ModelSideObject = new Model().SelectModelObject(myguid);
                            ObjectsToSelect.Add(ModelSideObject);
                        }
                    }


                }
            }
            var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MOS.Select(ObjectsToSelect);


        }
        private void dgvconnectioncompare_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            SelectConnection();
        }

        private void dgvconnectioncompare_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            SelectConnection();
        }

 



        private void btnload_Click(object sender, EventArgs e)
        {
            DateTime s_tm = DateTime.Now;
            pbar1.Visible = true ;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "LoadConnection", "", skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
            string filename = modelpath + "\\SK_MASSCONNL_" + skWinLib.username + ".sknum";
            //if (e.Clicks == 2)
            //    filename = modelpath + "\\SK_MASSCONNT_" + skWinLib.username + ".sknum";
            if (File.Exists(filename))
                Import_DataGridView(dgvconnectioncompare, filename, "PIE", "");

            updatecolumnheader();
            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = span.Seconds.ToString() + " Ready.";
            pbar1.Visible = false;
        }

        private void chkonTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkonTop.Checked;
        }



        




        private void btnselectconnection_Click(object sender, EventArgs e)
        {

            DateTime s_tm = DateTime.Now;
            pbar1.Visible = true;
            //lblguid.Text = "";
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            skTSLib.setStartup(skApplicationName, skApplicationVersion, "StartListConnection", "");
            int selct = 0;

            if (tabControl1.SelectedIndex == 0)
            {
                //select connection from model
                ArrayList mySelectedObjects = getConnectionObjects("Select");
                //preview the selected connection in treeview
                ArrayList mySelectedConnections = PreviewSelectedConnection(mySelectedObjects, tvwpreviewconn);
                ////check different type of connection and process
                //ArrayList myAttributeData = getAttributeData_FromSelectedConnectionObject(mySelectedObjects);
                ////update user interface to show the connection data with highlight difference for two
                //PreviewAttributeData(myAttributeData);
                selct = mySelectedObjects.Count;
            }

            else if (tabControl1.SelectedIndex == 1)
            {
                ArrayList myconnobj = getConnectionObjects("Select");
                ArrayList myConnectionData1 = PreviewSelectedConnection(myconnobj, tvwpreviewconntype1);
                selct = myconnobj.Count;
            }






            TimeSpan span =skTSLib.setCompletion(skApplicationName, skApplicationVersion, "CompleteListConnection",  ";Count:" + selct.ToString(), s_tm);
            lblsbar1.Text = "Ready " + span.Minutes.ToString() + ":" + span.Seconds.ToString();
            pbar1.Visible = false;
        }


        private ArrayList SelectGUIDObjects(ArrayList myguidlist)
        {
            ArrayList ObjectsToSelect = new ArrayList();
            //ObjectsToSelect.Add("ESSKAY");
            Tekla.Structures.Identifier myIdentifier = new Tekla.Structures.Identifier();
            foreach (string guid in myguidlist)
            {
                myIdentifier = new Tekla.Structures.Identifier(guid);
                bool flag = myIdentifier.IsValid();
                if (flag)
                {
                    Tekla.Structures.Model.ModelObject ModelSideObject = new Tekla.Structures.Model.Model().SelectModelObject(myIdentifier);
                    if (ModelSideObject.Identifier.IsValid())
                    {
                        ObjectsToSelect.Add(ModelSideObject);
                    }
                }
            }

            return ObjectsToSelect;

        }


        private void tvwpreviewconn_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            string statusmsg = "";
            DateTime s_tm = DateTime.Now;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            ArrayList ObjectsToSelect = new ArrayList();

            //get treeview node
            TreeNode selnode = tvwpreviewconn.SelectedNode;

            if (selnode != null)
            {
                //clear existing guidlist
                nodeguidlist.Clear();
                if (selnode.Text.Contains("guid:"))
                {
                    nodeguidlist.Add(selnode.Text.Substring(5));
                    lblguid.Text = selnode.Text.Substring(5);
                }
                else
                {
                    getTreeViewGUID(selnode);
                    lblguid.Text = "";
                }
            }




            //Create ObjectToSelect from guidlist
            ObjectsToSelect = SelectGUIDObjects(nodeguidlist);

            //Select and Zoom
            var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MOS.Select(ObjectsToSelect);
            TSM.Operations.Operation.RunMacro("ZoomSelected.cs");

            //Check wether select + group is selected 
            if (chkselgrp.Checked == true && selnode.Parent != null)
            {
                //select connection from model
                ArrayList mySelectedObjects = getConnectionObjects("Select");
                //check different type of connection and process
                ArrayList myAttributeData = getAttributeData_FromSelectedConnectionObject(mySelectedObjects);
                //update user interface to show the connection data with highlight difference for two
                PreviewAttributeData(myAttributeData);
                //Esskay Grouping Confidential
                groupAttributeData();
                statusmsg =  (dgvconnectioncompare.Columns.Count - 1).ToString() + " Group(s) in " + ObjectsToSelect.Count.ToString() + " selected.";
            }
            else
                statusmsg = ObjectsToSelect.Count.ToString() + " selected.";

            


            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = span.Minutes.ToString() + ":" + span.Seconds.ToString() + " Ready. " + statusmsg;
        }

        private void tvwpreviewconn_AfterSelect(object sender, TreeViewEventArgs e)
        {




        }

        
        //get all checked node
        public void getTreeViewCheckedNode(TreeNodeCollection nodes)
        {
            foreach (TreeNode childNode in nodes)
            {
                if (childNode.Checked)
                {
                    getTreeViewGUID(childNode);
                }
                else
                    getTreeViewCheckedNode(childNode.Nodes);
            }
        }

        //get all guid childnode from the selected parent node
        private void getTreeViewGUID(TreeNode selnode)
        {
            if (selnode != null)
            {
                if (selnode.Text.Contains("guid:"))
                    nodeguidlist.Add(selnode.Text.Substring(5));
                else
                {
                    //ArrayList ObjectsToSelect = new ArrayList();
                    foreach (TreeNode childnode in selnode.Nodes)
                    {
                        int chlct = childnode.Nodes.Count;
                        if (chlct == 0)
                        {
                            if (childnode.Text.IndexOf("guid:") == 0)
                            {
                                if (!nodeguidlist.Contains(childnode.Text.Substring(5)))
                                    nodeguidlist.Add(childnode.Text.Substring(5));
                            }
                        }
                        else
                        {
                            getTreeViewGUID(childnode);
                        }
                    }
                }
            }


        }

        private void btnselect_Click(object sender, EventArgs e)
        {
            pbar1.Visible = true;
            pbar1.Maximum = 6;
            string statusmsg = "";
            //lblguid.Text = "";
            DateTime s_tm = DateTime.Now;
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            ArrayList ObjectsToSelect = new ArrayList();
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "ReadConnection", "",skTSLib.ModelName,skTSLib.Version,skTSLib.Configuration );
            nodeguidlist.Clear();
            getTreeViewCheckedNode(tvwpreviewconn.Nodes);
            pbar1.Maximum = pbar1.Maximum + nodeguidlist.Count + 1;
            pbar1.Refresh();
            //get treeview node
            TreeNode selnode = tvwpreviewconn.SelectedNode;

            if (nodeguidlist.Count >=1)
            {
                //get selection list from guidlist
                ObjectsToSelect = SelectGUIDObjects(nodeguidlist);
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                //select and zoom
                var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                MOS.Select(ObjectsToSelect);
                TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
                //select connection from model
                ArrayList mySelectedObjects = getConnectionObjects("Select");
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                //check different type of connection and process
                ArrayList myAttributeData = getAttributeData_FromSelectedConnectionObject(mySelectedObjects);
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                //update user interface to show the connection data with highlight difference for two
                PreviewAttributeData(myAttributeData);
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
                //Check wether select + group is selected 
                if (chkselgrp.Checked == true)
                {
                    //Esskay Grouping Confidential
                    groupAttributeData();
                    statusmsg = (dgvconnectioncompare.Columns.Count - 1).ToString() + " Group(s) in " + ObjectsToSelect.Count.ToString() + " selected.";
                }
                pbar1.Value = pbar1.Value + 1;
                pbar1.Refresh();
            }

            //Function display user based on connection guid selected by user
            if (lblguid.Text.Length >= 3)
                UpdateDataGridbasedUserConnectionSelectedGUID(lblguid.Text.ToString());
            pbar1.Value = pbar1.Value + 1;
            pbar1.Refresh();

            TimeSpan span = DateTime.Now.Subtract(s_tm);
            lblsbar1.Text = span.Minutes.ToString() + ":" + span.Seconds.ToString() + " Ready. " + statusmsg;
            pbar1.Visible = false;
        }

        private void btnclear_Click_1(object sender, EventArgs e)
        {
            //clear all data grid data
            dgvconnectioncompare.Rows.Clear();
            dgattibutedata.Rows.Clear();
            
        }

        private void btngrpguid_Click(object sender, EventArgs e)
        {
            DateTime s_tm = DateTime.Now;
            lblguid.Text = "";
            lblsbar1.Text = "Pls.Wait!!! " + s_tm.ToString("h:mm:ss");
            skTSLib.setStartup(skApplicationName, skApplicationVersion, "StartListGuid", "");
            int selct = 0;

            MyModel = skTSLib.ModelAccess.ConnectToModel();
            if (MyModel.GetConnectionStatus() && MyModel.GetInfo().ModelPath != string.Empty)
            {
                Picker mymasterconn = new Picker();
                TSM.ModelObject myconnobject =  mymasterconn.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, "Select Connection to Compare Others: ");

                if (myconnobject != null)
                {
                    TSM.Connection myconn = myconnobject as TSM.Connection;
                    if (myconn != null)
                    {
                        lblguid.Text = myconn.Identifier.GUID.ToString();
                    }
                    else
                    {
                        TSM.Detail mydetail = myconnobject as TSM.Detail;
                        if (mydetail != null)
                        {
                            lblguid.Text = mydetail.Identifier.GUID.ToString();
                        }
                        else
                        {
                            TSM.Component mycoomp = myconnobject as TSM.Component;
                            if (mycoomp != null)
                            {
                                lblguid.Text = mycoomp.Identifier.GUID.ToString();
                            }
                        }
                    }
                }
                //Function display user based on connection guid selected by user
                if (lblguid.Text.Length >=3)
                    UpdateDataGridbasedUserConnectionSelectedGUID(lblguid.Text.ToString());
            }
            TimeSpan span = skTSLib.setCompletion(skApplicationName, skApplicationVersion, "CompleteListGuid", ";Count:" + selct.ToString(), s_tm);
            lblsbar1.Text = "Ready " + span.Minutes.ToString() + ":" + span.Seconds.ToString();
        }
        private void UpdateDataGridbasedUserConnectionSelectedGUID(string masterguid)
        {

            int userguid = -1;
            //check whether user selected guid in the preview dataview control list
            for (int i= 1; i < dgvconnectioncompare.Columns.Count; i++)
            {
                string guids = dgvconnectioncompare.Rows[0].Cells[i].Value.ToString();
                if (guids.IndexOf(masterguid) >= 0)
                {
                    userguid = i;
                    break;
                }
                    
            }

            //if userguid exists then swap column data
            if(userguid != -1 && userguid != 1)
            {
                for (int i = 0; i < dgvconnectioncompare.Rows.Count; i++)
                {
                    string swapfirstcoldata = string.Empty;
                    if (dgvconnectioncompare.Rows[i].Cells[1].Value != null)
                        swapfirstcoldata = dgvconnectioncompare.Rows[i].Cells[1].Value.ToString();

                    string swapusercoldata = string.Empty;
                    if (dgvconnectioncompare.Rows[i].Cells[userguid].Value != null)
                        swapusercoldata = dgvconnectioncompare.Rows[i].Cells[userguid].Value.ToString();
                    //swap the user guid to the first column
                    if (swapfirstcoldata != swapusercoldata)
                    {
                        dgvconnectioncompare.Rows[i].Cells[1].Value = swapusercoldata;
                        dgvconnectioncompare.Rows[i].Cells[userguid].Value = swapfirstcoldata;
                    }


                }
            }

            //highlight the color
        }

        private void dgattibutedata_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            lblsbar2.Text = "dgattibutedata_RowHeaderMouseDoubleClick";
            if (e.RowIndex >= 0)
            {
                int rowindex = e.RowIndex;
                DataGridViewRow row = this.dgattibutedata.Rows[rowindex];
                 
                txtattfieldname.Text = row.Cells["AttributeName"].Value.ToString();
                txtattfieldvalue.Text = row.Cells["AttributeValue"].Value.ToString();
            }

            lblsbar2.Text = "";
        }

        
        private void btnsetallattibuteinfo_Click(object sender, EventArgs e)
        {
            bool flag = false;
            ModelObjectEnumerator MyEnum = (new TSM.UI.ModelObjectSelector()).GetSelectedObjects();
            while (MyEnum.MoveNext())
            {

                TSM.Connection ThisConnection = MyEnum.Current as TSM.Connection;
                if (ThisConnection != null)
                {
                    flag = false;
                    foreach (DataGridViewRow row in dgattibutedata.Rows)
                    {
                        DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)(row.Cells["TSUpdate"]);
                        if (cell.Value.ToString().ToUpper()  == "TRUE")
                        {
                            string attfieldname = row.Cells["AttributeName"].Value.ToString();
                            string attfieldvalue = row.Cells["AttributeValue"].Value.ToString();
                            skTSLib.setComponent(ThisConnection, attfieldname, attfieldvalue);
                            flag = true;
                        }
                    }
                    if (flag == true)
                        ThisConnection.Modify();

                }

                TSM.Detail ThisDetail = MyEnum.Current as TSM.Detail;
                if (ThisDetail != null)
                {
                    flag = false;
                    foreach (DataGridViewRow row in dgattibutedata.Rows)
                    {
                        DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)(row.Cells["TSUpdate"]);
                        if (cell.Value.ToString().ToUpper() == "TRUE")
                        {
                            string attfieldname = row.Cells["AttributeName"].Value.ToString();
                            string attfieldvalue = row.Cells["AttributeValue"].Value.ToString();
                            skTSLib.setComponent(ThisDetail, attfieldname, attfieldvalue);
                            flag = true;
                        }
                    }
                    if (flag == true)
                        ThisDetail.Modify();
                }

            }
            //redraw view
            ReDrawView();
        }
        private void ReDrawView()
        {
            var akit = new MacroBuilder();
            akit.Callback("acmd_update_views", "", "main_frame");
            akit.Callback("acmd_redraw_selected_view", "", "View_01 window_1");
            akit.Run();
        }
        private void btnsetattributes_Click(object sender, EventArgs e)
        {
            Picker mymasterconn = new Picker();
            TSM.ModelObject myconnobject = mymasterconn.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, "Select Connection to Update : ");

            if (myconnobject != null)
            {
                TSM.Connection ThisConnection = myconnobject as TSM.Connection;
                if (ThisConnection != null)
                {
                    skTSLib.setComponent(ThisConnection, txtattfieldname.Text, txtattfieldvalue.Text);
                    ThisConnection.Modify();
                }

                TSM.Detail ThisDetail = myconnobject as TSM.Detail;
                if (ThisDetail != null)
                {
                    skTSLib.setComponent(ThisDetail, txtattfieldname.Text, txtattfieldvalue.Text);
                    ThisDetail.Modify();
                }
            }
        }

        private void btngetattributes_Click(object sender, EventArgs e)
        {
            try
            {
                string getattvalue = "-1";
                Picker mymasterconn = new Picker();
                TSM.ModelObject myconnobject = mymasterconn.PickObject(Picker.PickObjectEnum.PICK_ONE_OBJECT, "Select Connection : ");

                if (myconnobject != null)
                {
                    TSM.BaseComponent Component = myconnobject as TSM.BaseComponent;
                    if (Component != null)
                        getattvalue = skTSLib.getComponent(Component, txtattfieldname.Text);
                }

                if (getattvalue != "-1")
                    txtattfieldvalue.Text = getattvalue;
            }
            catch
            {

            }


        }
    }


}

