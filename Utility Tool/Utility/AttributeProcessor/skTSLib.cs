//Omm Muruga 
//Common Function's irrespective of any application, tool, plugin
//Ensure naming of function is meaningfull to understand
//Dont have too many short variables
//If any major changes required in functions please let Vijayakumar know's about this.
//Don't modifiy existing function/class/sub unless you are aware of the consequences
//Use Commentline as much as possible for future reference and other's can understand easily
//Most function's or not tested so ensure its quality at your own risk.
//08Jan22 1502


using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Win32;
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

public class skTSLib
{


    public static TSM.Model skTSLibModel = new TSM.Model();
    public static string ModelPath = skTSLibModel.GetInfo().ModelPath;
    public static string ModelName = skTSLibModel.GetInfo().ModelName;
    public static bool IsSharedmodel = skTSLibModel.GetInfo().SharedModel;
    public static string CurrentUser = Tekla.Structures.TeklaStructuresInfo.GetCurrentUser();
    public static string Version = Tekla.Structures.TeklaStructuresInfo.GetCurrentProgramVersion();
    public static string Configuration = new Tekla.Structures.ModuleManager.ProgramConfigurationEnum().ToString();


    public class ModelAccess
    {

        private static Tekla.Structures.Model.Model m_ModelConnection = null;



        /// <summary>
        /// Checks if a model connection is established, and attempts to create one if needed.
        /// </summary>
        /// <returns>true if successfully connected to the  model, false otherwise</returns>
        private static bool ConnectedToModel
        {
            get
            {
                bool connection = false;
                ConnectToModel(out connection);
                return connection;
            }
        }

        /// <summary>
        /// Gets a Model connnection
        /// </summary>
        /// <returns>The Model or null if unable to connect</returns>
        public static Tekla.Structures.Model.Model ConnectToModel()
        {
            bool connection = false;
            return ConnectToModel(out connection);
        }

        /// <summary>
        /// Gets a model connection
        /// </summary>
        /// <param name="model">the Model connection</param>
        /// <returns>true on success, false otherwise</returns>
        public static bool ConnectToModel(out Tekla.Structures.Model.Model model)
        {
            bool connection = false;
            ConnectToModel(out connection);
            model = m_ModelConnection;
            return connection;
        }

        /// <summary>
        /// Gets a Model connection
        /// </summary>
        /// <param name="ConnectedToModel">set to true if a model connection was made, false otherwise.</param>
        /// <returns>The Model or null if unable to connect</returns>
        public static Tekla.Structures.Model.Model ConnectToModel(out bool ConnectedToModel)
        {
            ConnectedToModel = false;
            if (m_ModelConnection == null)
            {
                try
                {
                    m_ModelConnection = new Tekla.Structures.Model.Model();
                }
                catch
                {
                    ConnectedToModel = false;
                    m_ModelConnection = null;
                }
            }
            try
            {
                if (m_ModelConnection.GetConnectionStatus())
                {
                    if (m_ModelConnection.GetInfo().ModelName == "")
                    {
                        ConnectedToModel = false;
                        m_ModelConnection = null;
                    }
                    else
                    {
                        ConnectedToModel = true;

                        ModelPath = skTSLibModel.GetInfo().ModelPath;
                        ModelName = skTSLibModel.GetInfo().ModelName;
                        IsSharedmodel = skTSLibModel.GetInfo().SharedModel;
                        CurrentUser = Tekla.Structures.TeklaStructuresInfo.GetCurrentUser();
                        Version = Tekla.Structures.TeklaStructuresInfo.GetCurrentProgramVersion();
                        Configuration = new Tekla.Structures.ModuleManager.ProgramConfigurationEnum().ToString();

                    }
                }
                else
                {
                    ConnectedToModel = false;
                    m_ModelConnection = null;
                }
            }
            catch
            {
                ConnectedToModel = false;
                m_ModelConnection = null;
            }

            //
            //if (m_ModelConnection == null)
            //{
            //    System.Windows.Forms.Application.ExitThread();
            //}


            return m_ModelConnection;
        }
    }
    public class SKDrawing
    {

    }
    public class Model
    {
        //TSM.Model skTSLibModel = new TSM.Model();

        private static bool IsBeam(Beam MyBeam)
        {



            string IDstring = Convert.ToString(MyBeam.Identifier.ID);

            Double strXS = MyBeam.StartPoint.X;
            Double strYS = MyBeam.StartPoint.Y;
            Double strZS = MyBeam.StartPoint.Z;
            Double strXE = MyBeam.EndPoint.X;
            Double strYE = MyBeam.EndPoint.Y;
            Double strZE = MyBeam.EndPoint.Z;


            if (Math.Abs(strZS - strZE) < 1)
            {
                return true;
            }
            else
            {
                return false;

            }
        }

        private TSM.Beam GetBeam(int CurrentID)
        {
            //lblsbar2.Text = "";
            Beam MyBeam = new Beam();
            try
            {
                //Set up model connection w/TS
                //Model skTSLibModel = new TSM.Model();
                if (skTSLibModel.GetConnectionStatus() &&
                    skTSLibModel.GetInfo().ModelPath != string.Empty)
                {
                    TransformationPlane CurrentTP = skTSLibModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());

                    Identifier PartID = new Identifier(CurrentID);
                    TSM.ModelObject MOB = skTSLibModel.SelectModelObject(PartID);

                    MyBeam = MOB as Beam;

                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);

                } // end of if that checks to see if there is a model and one is open

                else
                {
                    //If no connection to model is possible show error message
                    MessageBox.Show("Tekla Structures not open or model not open.", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }  //end of try loop

            catch { }
            //lblsbar2.Text = "";

            return MyBeam;

        }
        private TSM.Beam GetBeam(int CurrentID, bool Status)
        {
            //lblsbar2.Text = "";
            Beam MyBeam = new Beam();
            try
            {
                //Set up model connection w/TS
                //Model skTSLibModel = new Model();
                if (skTSLibModel.GetConnectionStatus() &&
                    skTSLibModel.GetInfo().ModelPath != string.Empty)
                {
                    TransformationPlane CurrentTP = skTSLibModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    if (Status == true)
                    {
                        skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                    }
                    Identifier PartID = new Identifier(CurrentID);
                    TSM.ModelObject MOB = skTSLibModel.SelectModelObject(PartID);

                    MyBeam = MOB as Beam;

                    if (Status == true)
                    {
                        skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);
                    }
                } // end of if that checks to see if there is a model and one is open

                else
                {
                    //If no connection to model is possible show error message
                    MessageBox.Show("Tekla Structures not open or model not open.", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }  //end of try loop

            catch { }
            //lblsbar2.Text = "";

            return MyBeam;

        }
        public static void export_screw_database()
        {
            string catalogName = "export_sk_screwdb.cs";
            string XS_MACRO_DIRECTORY = string.Empty;
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref XS_MACRO_DIRECTORY);



            if (XS_MACRO_DIRECTORY.IndexOf(';') > 0)
                XS_MACRO_DIRECTORY = XS_MACRO_DIRECTORY.Remove(XS_MACRO_DIRECTORY.IndexOf(';'));

            string screwdbscript = "namespace Tekla.Technology.Akit.UserScript{" +
            "public class Script {" +
            "public static void Run(Tekla.Technology.Akit.IScript akit){" +
            "akit.Callback(\"acmd_export_screw_database\", \"\", \"main_frame\"); }}};";

            File.WriteAllText(Path.Combine(XS_MACRO_DIRECTORY, catalogName), screwdbscript);
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + catalogName);

        }

        public static bool IsParallel(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {

            Tekla.Structures.Geometry3d.Line Line1 = new Tekla.Structures.Geometry3d.Line(Beam1.StartPoint, Beam1.EndPoint);
            Tekla.Structures.Geometry3d.Line Line2 = new Tekla.Structures.Geometry3d.Line(Beam2.StartPoint, Beam2.EndPoint);

            return Tekla.Structures.Geometry3d.Parallel.LineToLine(Line1, Line2);
        }


        public static double getflangethickness(string profile)
        {
            double result = 0;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_I || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_C)
            {
                double val = (libProfile.aProfileItemParameters[3] as ProfileItemParameter).Value;
                result = val;
            }
            else if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_L)
            {
                double val = (libProfile.aProfileItemParameters[2] as ProfileItemParameter).Value;
                result = val;
            }
            return result;
        }

        public static double getwebthickness(string profile)
        {
            double result = 0;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_I || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_C)
            {
                double val = (libProfile.aProfileItemParameters[2] as ProfileItemParameter).Value;
                result = val;
            }


            return result;
        }

        public static double get_thickness(string profile, string boltcat)
        {
            double result = 0;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            //BoltItem.BoltItemTypeEnum teklaProfileType = libProfile.;




            return result;
        }
        public static double getwidth(string profile)
        {
            double result = 0;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_I || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_C || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_L)
            {
                double val = (libProfile.aProfileItemParameters[1] as ProfileItemParameter).Value;
                result = val;
            }
            //else if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_P)
            //{
            //    double val = (libProfile.aProfileItemParameters[4] as ProfileItemParameter).Value;
            //    result = val;
            //}

            return result;
        }
        public static double getrootradious(string profile)
        {
            double result = 0;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_I || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_C || teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_L)
            {
                double val = (libProfile.aProfileItemParameters[4] as ProfileItemParameter).Value;
                result = val;
            }
            //else if (teklaProfileType == ProfileItem.ProfileItemTypeEnum.PROFILE_P)
            //{
            //    double val = (libProfile.aProfileItemParameters[4] as ProfileItemParameter).Value;
            //    result = val;
            //}

            return result;
        }

        public static double getKvalue(string profile)
        {
            double result = -1;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            //get k value directly
            foreach (ProfileItemParameter userparam in libProfile.aProfileItemUserParameters)
            {
                string key = userparam.StringValue;
                if (key.ToUpper().Trim() == "K")
                {
                    result = userparam.Value;
                    break;
                }
            }
            return result;
        }

        public static double getK1value(string profile)
        {
            double result = -1;
            LibraryProfileItem libProfile = new LibraryProfileItem();
            libProfile.Select(profile);
            ProfileItem.ProfileItemTypeEnum teklaProfileType = libProfile.ProfileItemType;
            //get k value directly
            foreach (ProfileItemParameter userparam in libProfile.aProfileItemUserParameters)
            {
                string key = userparam.StringValue;
                if (key.ToUpper().Trim() == "K1")
                {
                    result = userparam.Value;
                    break;
                }
            }
            return result;
        }


        public static void selectDataGridselectedrows(DataGridView mydatagrid, int colid = 0, string idtype = "GUID", bool zoom = true)
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
                        myguid = myguid.ToUpper().Replace("GUID:", "").Replace("GUID","").Trim();
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
                            //ModelSideObject.Select();
                            if (!ObjectsToSelect.Contains(ModelSideObject))
                                ObjectsToSelect.Add(ModelSideObject);
                        }



                    }
                    i++;

                }

                var MOS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                MOS.Select(ObjectsToSelect);
                if (zoom == true) TSM.Operations.Operation.RunMacro("ZoomSelected.cs");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Select/Zoom: " + ex.Message);
            }


        }

        
        public static bool IsPointNearRootRadius(TSG.Point ClashCheckPart_Point, TSG.Point Member_TopMiddleEndPoint,double width, double rootradius, double flangethickness, double webthickness)
        {
            double chkx = ClashCheckPart_Point.X;
            double chky = Math.Abs(ClashCheckPart_Point.Y);
            double chkz = Math.Abs(ClashCheckPart_Point.Z);

            double memx = Member_TopMiddleEndPoint.X;
            double memy = Member_TopMiddleEndPoint.Y;
            double memz = Member_TopMiddleEndPoint.Z;

            double flgthkrootradius = flangethickness + rootradius;
            double webthk2rootradius = (webthickness / 2) + rootradius;

            //check weather clash part is below start point x  and greater than end point x
            if ((chkx < 0) || (chkx > memx))
                return false;


            //check weather clash part isgreater than end point y
            if (chky > memy)
                return false;

            //check weather clash part is greater than end point z
            if (chkz > (width / 2.0))
                return false;

            double chkdist = TSG.Distance.PointToPoint(new TSG.Point(memz + webthk2rootradius, memy - flgthkrootradius, 0), new TSG.Point(chkz, chky, 0));
            
            if (chkdist > flgthkrootradius && flgthkrootradius > webthk2rootradius)
                return false;
            else if (chkdist > webthk2rootradius && flgthkrootradius < webthk2rootradius)
                return false;

            return true;
        }

        public static ArrayList getRootRadiusNearPoint(Tekla.Structures.Solid.FaceEnumerator faces, TSG.Point Member_TopMiddleEndPoint, double width, double rootradius, double flangethickness, double webthickness)
        {
            ArrayList result = new ArrayList();
            result.Add("Esskay");
            while (faces.MoveNext())
            {
                if (faces.Current is Tekla.Structures.Solid.Face face)
                {
                    Tekla.Structures.Solid.LoopEnumerator Loops = face.GetLoopEnumerator();
                    while (Loops.MoveNext())
                    {
                        if (Loops.Current is Tekla.Structures.Solid.Loop Loop)
                        {
                            Tekla.Structures.Solid.VertexEnumerator CornerPoints = Loop.GetVertexEnumerator();
                            while (CornerPoints.MoveNext())
                            {
                                if (CornerPoints.Current is TSG.Point myPoint)
                                {
                                    bool isnear = IsPointNearRootRadius(myPoint, Member_TopMiddleEndPoint, width, rootradius, flangethickness, webthickness);
                                    if (isnear == true)
                                    {
                                        if (!result.Contains(myPoint))
                                        {
                                            result.Add(myPoint);
                                            //TSM.ControlPoint test = new TSM.ControlPoint (myPoint);
                                            //bool flag = test.Insert();

                                        }
                                    }
 

                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
        public static ArrayList getUniqueFacePoints(Tekla.Structures.Solid.FaceEnumerator faces)
        {
            ArrayList result = new ArrayList();
            result.Add("Esskay");
            while (faces.MoveNext())
            {
                if (faces.Current is Tekla.Structures.Solid.Face face)
                {
                    Tekla.Structures.Solid.LoopEnumerator Loops = face.GetLoopEnumerator();
                    while (Loops.MoveNext())
                    {
                        if (Loops.Current is Tekla.Structures.Solid.Loop Loop)
                        {
                            Tekla.Structures.Solid.VertexEnumerator CornerPoints = Loop.GetVertexEnumerator();
                            while (CornerPoints.MoveNext())
                            {
                                if (CornerPoints.Current is TSG.Point myPoint)
                                {
                                    if (!result.Contains(myPoint))
                                    {
                                        result.Add(myPoint);
                                        //TSM.ControlPoint test = new TSM.ControlPoint (myPoint);
                                        //bool flag = test.Insert();

                                    }
                                       
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        public static bool IsMainPart(Tekla.Structures.Model.ModelObject part)
        {
            bool res = false;
            int ismain = 0;
            part.GetReportProperty("MAIN_PART", ref ismain);
            if (ismain == 1)
                res = true;
            return res;
        }
        //public static ArrayList getChilderenObject(TSM.ModelObject ModelObj)
        //{
        //    ArrayList result = new ArrayList();
        //    result.Add("Esskay");
        //    if (ModelObj is Assembly assmbly)
        //    {
        //        List<ModelObject> concreteparts = new List<ModelObject>();

        //        var main = assmbly.GetMainPart();
        //        result.Add( assmbly.GetSecondaries());
        //        var subs = assmbly.GetSubAssemblies()
        //    }

        //    return result;

        //}

        public static bool IsParallelBeams(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {

            LineSegment Line1 = new LineSegment(Beam1.StartPoint, Beam1.EndPoint);
            LineSegment Line2 = new LineSegment(Beam2.StartPoint, Beam2.EndPoint);

            return Tekla.Structures.Geometry3d.Parallel.LineSegmentToLineSegment(Line1, Line2);
        }


        public static string GetFramingCondition(Tekla.Structures.Geometry3d.Point MyPoint, Tekla.Structures.Model.Beam Primary, Tekla.Structures.Model.Beam Secondary, List<Tekla.Structures.Model.Beam> Secondaries, System.Windows.Forms.CheckBox CheckWrapAroundFlag,double nodeTolerance)
        {
            string sFramingCondition = "";

            if (CheckWrapAroundFlag.Checked == true)
            {
                if (IsBrace(Secondary) == true)
                {
                    double dAngle = GetAngle(Primary, Secondary);

                    if (dAngle > 20 && dAngle < 70)
                    {
                        Tekla.Structures.Model.Beam Primary3 = GetWebPrimary(Primary, Secondary, Secondaries, nodeTolerance);
                        Tekla.Structures.Model.Beam Primary2 = GetFlangePrimary(Primary, Secondary, Secondaries, nodeTolerance );

                        if (Primary2 != null)
                        {
                            if (Is2Flange(Primary, Secondary) == true)
                            {
                                sFramingCondition = "Wrap Around @ Flange";
                            }
                        }
                        if (Primary3 != null)
                        {
                            if (Is2Flange(Primary, Secondary) == false)
                            {
                                sFramingCondition = "Wrap Around @ Web";
                            }
                        }
                    }
                    else if (dAngle > 85)
                    {

                        Tekla.Structures.Model.Beam Primary3 = GetWebPrimary(Primary, Secondary, Secondaries, nodeTolerance);
                        Tekla.Structures.Model.Beam Primary2 = GetFlangePrimary(Primary, Secondary, Secondaries,nodeTolerance);

                        if (Primary2 != null && Primary3 != null)
                        {
                            if (IsColumn(Primary) == true)
                            {
                                sFramingCondition = "Wrap Around @ Column";
                            }
                            else if (IsBeam(Primary) == true)
                            {
                                sFramingCondition = "Wrap Around @ Beam";
                            }
                        }
                    }
                }
            }

            //int Tolerance = Convert.ToInt32(txtnodetolerance.Text);
            //double nodeTolerance = Convert.ToDouble(txtnodetolerance.Text);
            if (sFramingCondition == "")
            {
                if (IsSplice(Primary, Secondary) == true)
                {
                    if (Distance.PointToPoint(Primary.StartPoint, MyPoint) < nodeTolerance || Distance.PointToPoint(Primary.EndPoint, MyPoint) < nodeTolerance)
                    {
                        if (Distance.PointToPoint(Secondary.StartPoint, MyPoint) < nodeTolerance || Distance.PointToPoint(Secondary.EndPoint, MyPoint) < nodeTolerance)
                        {
                            if (IsColumn(Primary) == true && IsColumn(Secondary) == true)
                            {
                                sFramingCondition = "Column to Column Splice";
                            }
                            else
                            {
                                sFramingCondition = "Beam to Beam Splice";
                            }
                        }
                    }
                }
            }
            if (sFramingCondition == "")
            {
                if (IsParallelBeams(Primary, Secondary) == true)
                {
                    sFramingCondition = "Parallel Members";
                }
            }
            if (sFramingCondition == "")
            {
                if (IsColumn(Primary) == true)
                {
                    if (Is2Flange(Primary, Secondary) == true)
                    {
                        sFramingCondition = "Beam to Column Flange";
                    }
                    else
                    {
                        sFramingCondition = "Beam to Column Web";
                    }
                }
                else if (IsColumn(Secondary) == true)
                {
                    if (Is2Flange(Primary, Secondary) == true)
                    {
                        sFramingCondition = "Column to Beam Flange";
                    }
                    else
                    {
                        sFramingCondition = "Column to Beam Web";
                    }
                }
                else if (IsBeam(Primary) == true)
                {
                    if (Is2Flange(Primary, Secondary) == true)
                    {
                        sFramingCondition = "Beam to Beam Flange";
                    }
                    else
                    {
                        sFramingCondition = "Beam to Beam Web";
                    }
                }
                else if (skTSLib.Model.GetDistance(MyPoint, Primary) < nodeTolerance && skTSLib.Model.GetDistance(MyPoint, Secondary) < nodeTolerance)
                {
                    sFramingCondition = "Connected Ends";
                }
            }
            return sFramingCondition;
        }


        public static string GetSingleCondition(Tekla.Structures.Geometry3d.Point MyPoint, Tekla.Structures.Model.Beam Secondary, double Tolerance)
        {
            string sFramingCondition = "";
            //int Tolerance = Convert.ToInt32(txtnodetolerance.Text);
            //double Tolerance = Convert.ToDouble(txtnodetolerance.Text);


            if (IsColumn(Secondary) == true)
            {
                if (Secondary.StartPoint.Z - MyPoint.Z > Tolerance || Secondary.EndPoint.Z - MyPoint.Z > Tolerance)
                {
                    sFramingCondition = "Base Plate";
                }
                else
                {
                    sFramingCondition = "Cap Plate";
                }
            }
            else
            {
                sFramingCondition = "Free End";
            }
            return sFramingCondition;
        }

        public static string GetHittingFlange(Tekla.Structures.Model.Beam Primary, Tekla.Structures.Model.Beam Secondary,double nodeTolerance)
        {
            string sFramingCondition = "";

            if (Is2TopSide(Primary, Secondary, nodeTolerance) == true)
            {
                sFramingCondition = "Top";
            }
            else
            {
                sFramingCondition = "Bottom";
            }
            return sFramingCondition;
        }

        public static double GetAngle(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {
            Beam1.Select();
            Beam2.Select();

            Vector Beam1Vector = new Vector(Beam1.GetCoordinateSystem().AxisX);
            Vector Beam2Vector = new Vector(Beam2.GetCoordinateSystem().AxisX);

            double angle = 0.0;
            angle = Beam1Vector.GetAngleBetween(Beam2Vector);
            angle = angle * 180 / Math.PI;
            if (angle > 90)
            {
                angle = 180 - angle;
            }

            return angle;
        }
        public static Tekla.Structures.Model.Beam GetFlangePrimary(Tekla.Structures.Model.Beam Primary, Tekla.Structures.Model.Beam Secondary, List<Tekla.Structures.Model.Beam> Secondaries, double nodeTolerance)
        {
            Tekla.Structures.Model.Beam Primary2 = new Tekla.Structures.Model.Beam();
            bool Available = false;
            bool bTop = Is2TopSide(Primary, Secondary, nodeTolerance);

            foreach (Tekla.Structures.Model.Beam MyBeam in Secondaries)
            {
                if (MyBeam != Primary && MyBeam != Secondary)
                {
                    if (Is2TopSide(Primary, MyBeam, nodeTolerance) == bTop && Is2Flange(Primary, MyBeam) == true)
                    {
                        if (IsColumn(MyBeam) == true || IsBeam(MyBeam) == true)
                        {
                            Primary2 = MyBeam;
                            Available = true;
                        }
                    }
                }
            }
            if (Available == true)
            {
                return Primary2;
            }
            else
            {
                return null;
            }
        }

        public static Tekla.Structures.Model.Beam GetWebPrimary(Tekla.Structures.Model.Beam Primary, Tekla.Structures.Model.Beam Secondary, List<Tekla.Structures.Model.Beam> Secondaries, double nodeTolerance)
        {
            Tekla.Structures.Model.Beam Primary2 = new Tekla.Structures.Model.Beam();
            bool Available = false;
            bool bFront = Is2FrontSide(Primary, Secondary, nodeTolerance);

            foreach (Tekla.Structures.Model.Beam MyBeam in Secondaries)
            {
                if (MyBeam != Primary && MyBeam != Secondary)
                {
                    if (Is2FrontSide(Primary, MyBeam, nodeTolerance) == bFront && Is2Flange(Primary, MyBeam) == false)
                    {
                        if (IsColumn(MyBeam) == true || IsBeam(MyBeam) == true)
                        {
                            Primary2 = MyBeam;
                            Available = true;
                        }
                    }
                }
            }
            if (Available == true)
            {
                return Primary2;
            }
            else
            {
                return null;
            }
        }


        public static bool IsRightSide(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2, Tekla.Structures.Geometry3d.Point MyPoint, double nodeTolerance)
        {
            //int Tolerance = Convert.ToInt32(txtnodetolerance.Text);
            //double Tolerance = Convert.ToDouble(txtnodetolerance.Text);
            bool condition = false;
            try
            {  // model connection is not necessary
               //Set up model Component w/TS
            
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {
                    TransformationPlane CurrentTP = skTSLibModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(Beam1.GetCoordinateSystem()));
                    Beam2.Select();

                    if (Beam2.StartPoint.X - MyPoint.X > nodeTolerance || Beam2.EndPoint.X - MyPoint.X > nodeTolerance)
                    {
                        condition = true;
                    }

                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);
                    Beam2.Select();
                } // end of if that checks to see if there is a model and one is open
                else
                {
                    //If no Component to model is possible show error message
                    MessageBox.Show("Tekla Structures not open or model not open.",
                        "Component Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }  //end of try loop

            catch (Exception ex)
            {
                MessageBox.Show("Original error: IsRightSide " + ex.Message);
            }


            return condition;


        }
        public static bool Is2FrontSide(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2, double nodeTolerance)
        {
            //int Tolerance = Convert.ToInt32(txtnodetolerance.Text);
            //double Tolerance = Convert.ToDouble(txtnodetolerance.Text);
            bool condition = false;
            try
            {  // model connection is not necessary
               //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {
                    TransformationPlane CurrentTP = skTSLibModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(Beam1.GetCoordinateSystem()));
                    Beam2.Select();

                    if (Beam2.StartPoint.Z > nodeTolerance || Beam2.EndPoint.Z > nodeTolerance)
                    {
                        condition = true;
                    }

                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);
                    Beam2.Select();
                } // end of if that checks to see if there is a model and one is open
                else
                {
                    //If no Component to model is possible show error message
                    MessageBox.Show("Tekla Structures not open or model not open.",
                        "Component Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }  //end of try loop

            catch (Exception ex)
            {
                MessageBox.Show("Original error: IS2FRONTSIDE " + ex.Message);
            }


            return condition;


        }



        public static bool Is2TopSide(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2, double nodeTolerance)
        {
            //int Tolerance = Convert.ToInt32(txtnodetolerance.Text);
            //double Tolerance = Convert.ToDouble(txtnodetolerance.Text);
            bool condition = false;
            try
            {  // model connection is not necessary
               //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {
                    TransformationPlane CurrentTP = skTSLibModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(Beam1.GetCoordinateSystem()));
                    Beam2.Select();

                    if (Beam2.StartPoint.Y > nodeTolerance || Beam2.EndPoint.Y > nodeTolerance)
                    {
                        condition = true;
                    }
                    // i did it because the current is changes the nodepoint
                    skTSLibModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(CurrentTP);
                    Beam2.Select();

                } // end of if that checks to see if there is a model and one is open
                else
                {
                    //If no Component to model is possible show error message
                    MessageBox.Show("Tekla Structures not open or model not open.",
                        "Component Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }  //end of try loop

            catch (Exception ex)
            {
                MessageBox.Show("Original error: IS2TOPSIDE" + ex.Message);
            }


            return condition;


        }

        public static bool IsSecondaryLower(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {
            double ZValue = 0;

            if (Beam1.StartPoint.Z > Beam1.EndPoint.Z)
            {
                ZValue = Beam1.StartPoint.Z;
            }
            else
            {
                ZValue = Beam1.EndPoint.Z;
            }

            if (Beam2.StartPoint.Z < ZValue || Beam2.EndPoint.Z < ZValue)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool IsSplice(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {

            CoordinateSystem cs1 = Beam2.GetCoordinateSystem();
            CoordinateSystem cs2 = Beam1.GetCoordinateSystem();  // for primary

            double dot = Vector.Dot(cs1.AxisX.GetNormal(), cs2.AxisX.GetNormal());



            bool Connected = false;

            if (Distance.PointToPoint(Beam1.StartPoint, Beam2.StartPoint) < 10)
            {
                Connected = true;
            }
            else if (Distance.PointToPoint(Beam1.StartPoint, Beam2.EndPoint) < 10)
            {
                Connected = true;
            }
            else if (Distance.PointToPoint(Beam1.EndPoint, Beam2.EndPoint) < 10)
            {
                Connected = true;
            }
            else if (Distance.PointToPoint(Beam1.EndPoint, Beam2.StartPoint) < 10)
            {
                Connected = true;
            }


            if (Math.Abs(dot) > 0.95 && Connected == true) 
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsWebPerpendicular(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {

            CoordinateSystem cs1 = Beam1.GetCoordinateSystem();  // for primary
            CoordinateSystem cs2 = Beam2.GetCoordinateSystem();

            double dot = Vector.Dot(cs2.AxisY.GetNormal(), cs1.AxisX.GetNormal());

            if (Math.Abs(dot) < 0.5)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool Is2Flange(Tekla.Structures.Model.Beam Beam1, Tekla.Structures.Model.Beam Beam2)
        {
            CoordinateSystem cs1 = Beam1.GetCoordinateSystem();  // for primary
            CoordinateSystem cs2 = Beam2.GetCoordinateSystem();

            double dot = Vector.Dot(cs2.AxisX.GetNormal(), cs1.AxisY.GetNormal());

            if (Math.Abs(dot) > 0.2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsBrace(Tekla.Structures.Model.Beam MyBeam)
        {
            if (IsColumn(MyBeam) == false && IsBeam(MyBeam) == false)
            {
                return true;
            }
            else
            {
                return false;
            }
        }



        public static bool IsColumn(Tekla.Structures.Model.Beam MyBeam)
        {
            MyBeam.Select();

            Double douXS = MyBeam.StartPoint.X;
            Double douYS = MyBeam.StartPoint.Y;
            Double douZS = MyBeam.StartPoint.Z;
            Double douXE = MyBeam.EndPoint.X;
            Double douYE = MyBeam.EndPoint.Y;
            Double douZE = MyBeam.EndPoint.Z;

            if (Math.Abs(douYS - douYE) < 10 && Math.Abs(douXS - douXE) < 10)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsDetailAvailable(Tekla.Structures.Geometry3d.Point MyPoint, string sComponent)
        {
            bool Available = false;
            try
            {
                //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {

                    Tekla.Structures.Geometry3d.Point MinP = new Tekla.Structures.Geometry3d.Point(MyPoint);
                    Tekla.Structures.Geometry3d.Point MaxP = new Tekla.Structures.Geometry3d.Point(MyPoint);

                    MinP.X = MinP.X - 50;
                    MinP.Y = MinP.Y - 50;
                    MinP.Z = MinP.Z - 50;

                    MaxP.X = MaxP.X + 50;
                    MaxP.Y = MaxP.Y + 50;
                    MaxP.Z = MaxP.Z + 50;

                    ModelObjectEnumerator CrossEnum = skTSLibModel.GetModelObjectSelector().GetObjectsByBoundingBox(MinP, MaxP);

                    while (CrossEnum.MoveNext())
                    {
                        Detail Crossing = CrossEnum.Current as Detail;

                        if (Crossing != null && Crossing.Name == sComponent)
                        {
                            Available = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ISDETAILAVAILABLE. Original error: " + ex.Message);
            }
            return Available;
        }
        public static List<Tekla.Structures.Model.Beam> GetCrossParts(Tekla.Structures.Geometry3d.Point MyPoint, double Tolerance, string sIgnoreClass, string sIgnoreName)
        {
            List<Tekla.Structures.Model.Beam> CrossParts = new List<Tekla.Structures.Model.Beam>();

            try
            {

                //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {


                    Tekla.Structures.Geometry3d.Point MinP = new Tekla.Structures.Geometry3d.Point(MyPoint);
                    Tekla.Structures.Geometry3d.Point MaxP = new Tekla.Structures.Geometry3d.Point(MyPoint);

                    MinP.X = MinP.X - Tolerance;
                    MinP.Y = MinP.Y - Tolerance;
                    MinP.Z = MinP.Z - Tolerance;

                    MaxP.X = MaxP.X + Tolerance;
                    MaxP.Y = MaxP.Y + Tolerance;
                    MaxP.Z = MaxP.Z + Tolerance;

                    ModelObjectEnumerator CrossEnum = skTSLibModel.GetModelObjectSelector().GetObjectsByBoundingBox(MinP, MaxP);

                    while (CrossEnum.MoveNext())
                    {
                        Tekla.Structures.Model.Beam CrossBeam = CrossEnum.Current as Tekla.Structures.Model.Beam;

                        if (CrossBeam != null)
                        {
                            string PProfileType = "";
                            CrossBeam.GetReportProperty("PROFILE_TYPE", ref PProfileType);
                            if (PProfileType != "B" && sIgnoreClass.Contains(CrossBeam.Class) == false && sIgnoreName.Trim().ToUpper().Replace(" ", "").Contains(CrossBeam.Name.Trim().ToUpper().Replace(" ", "")) == false)
                            //if (PProfileType != "B")
                            {
                                CrossParts.Add(CrossBeam);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetCrossParts. Original error: " + ex.Message);
            }
            return CrossParts;
        }
        public static List<Tekla.Structures.Model.Beam> GetCrossParts(Tekla.Structures.Geometry3d.Point MyPoint1, Tekla.Structures.Geometry3d.Point MyPoint2, double Tolerance, string sIgnoreClass, string sIgnoreName)
        {
            List<Tekla.Structures.Model.Beam> CrossParts = new List<Tekla.Structures.Model.Beam>();

            try
            {

                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {

                    ModelObjectEnumerator CrossEnum = skTSLibModel.GetModelObjectSelector().GetObjectsByBoundingBox(MyPoint1, MyPoint2);

                    while (CrossEnum.MoveNext())
                    {
                        Tekla.Structures.Model.Beam CrossBeam = CrossEnum.Current as Tekla.Structures.Model.Beam;

                        if (CrossBeam != null)
                        {
                            string PProfileType = "";
                            CrossBeam.GetReportProperty("PROFILE_TYPE", ref PProfileType);
                            if (PProfileType != "B" && sIgnoreClass.Contains(CrossBeam.Class) == false && sIgnoreName.Trim().ToUpper().Replace(" ", "").Contains(CrossBeam.Name.Trim().ToUpper().Replace(" ", "")) == false)
                             {
                                CrossParts.Add(CrossBeam);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetCrossParts. Original error: " + ex.Message);
            }
            return CrossParts;
        }
        public static ArrayList GetCrossPartsID(Tekla.Structures.Geometry3d.Point MyPoint, Double Tolerance)
        {
            ArrayList CrossParts = new ArrayList();
            //string ids = "";

            try
            {

                //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {


                    Tekla.Structures.Geometry3d.Point MinP = new Tekla.Structures.Geometry3d.Point(MyPoint);
                    Tekla.Structures.Geometry3d.Point MaxP = new Tekla.Structures.Geometry3d.Point(MyPoint);

                    MinP.X = MinP.X - Tolerance;
                    MinP.Y = MinP.Y - Tolerance;
                    MinP.Z = MinP.Z - Tolerance;

                    MaxP.X = MaxP.X + Tolerance;
                    MaxP.Y = MaxP.Y + Tolerance;
                    MaxP.Z = MaxP.Z + Tolerance;

                    ModelObjectEnumerator CrossEnum = skTSLibModel.GetModelObjectSelector().GetObjectsByBoundingBox(MinP, MaxP);
               
                    while (CrossEnum.MoveNext())
                    {
                       
                        //Tekla.Structures.Model.ModelObject myObject = CrossEnum.Current as Tekla.Structures.Model.ModelObject;
                        //ids = ids + "\n" + myObject.Identifier.ID.ToString();

                        Tekla.Structures.Model.Beam CrossBeam = CrossEnum.Current as Tekla.Structures.Model.Beam;

                        if (CrossBeam != null)
                        {
                            string PProfileType = "";
                            CrossBeam.GetReportProperty("PROFILE_TYPE", ref PProfileType);

                            if (PProfileType != "B")
                            {
                                CrossParts.Add(CrossBeam.Identifier.ID);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetCrossPartsID. Original error: " + ex.Message);
            }
            return CrossParts;
        }

        //added newly for including contour plates
        public static ArrayList GetCrossParts(Tekla.Structures.Geometry3d.Point MyPoint, Double Tolerance)
        {
            ArrayList CrossParts = new ArrayList();
            string ids = "";

            try
            {

                //Set up model Component w/TS
                skTSLibModel = new TSM.Model();
                ModelPath = skTSLibModel.GetInfo().ModelPath;
                if (skTSLibModel.GetConnectionStatus() && ModelPath != string.Empty)
                {


                    Tekla.Structures.Geometry3d.Point MinP = new Tekla.Structures.Geometry3d.Point(MyPoint);
                    Tekla.Structures.Geometry3d.Point MaxP = new Tekla.Structures.Geometry3d.Point(MyPoint);

                    MinP.X = MinP.X - Tolerance;
                    MinP.Y = MinP.Y - Tolerance;
                    MinP.Z = MinP.Z - Tolerance;

                    MaxP.X = MaxP.X + Tolerance;
                    MaxP.Y = MaxP.Y + Tolerance;
                    MaxP.Z = MaxP.Z + Tolerance;

                    ModelObjectEnumerator CrossEnum = skTSLibModel.GetModelObjectSelector().GetObjectsByBoundingBox(MinP, MaxP);

                    while (CrossEnum.MoveNext())
                    {

                        Tekla.Structures.Model.ModelObject myObject = CrossEnum.Current as Tekla.Structures.Model.ModelObject;
                        ids = ids + "\n" + myObject.Identifier.ID.ToString();

                        Tekla.Structures.Model.Beam CrossBeam = CrossEnum.Current as Tekla.Structures.Model.Beam;
                        string PProfileType = "";
                        if (CrossBeam != null)
                        {
                            CrossBeam.GetReportProperty("PROFILE_TYPE", ref PProfileType);

                            //if (PProfileType != "B")
                            //{
                                CrossParts.Add(CrossBeam);
                                
                            //}
                        }
                        Tekla.Structures.Model.ContourPlate CrossPlate = CrossEnum.Current as Tekla.Structures.Model.ContourPlate;
                        if (CrossPlate != null)
                        {

                            CrossPlate.GetReportProperty("PROFILE_TYPE", ref PProfileType);

                            //if (PProfileType == "B")
                            //{
                                CrossParts.Add(CrossPlate);
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetCrossPartsID. Original error: " + ex.Message);
            }
            return CrossParts;
        }
        public static bool UpdateConnectionAttribute(Tekla.Structures.Model.BaseComponent Component, ArrayList AttributeList)
        {
            bool flag = false;
            try
            {
                foreach (string Attribute in AttributeList)
                {
                    if (Attribute.IndexOf("|") >= 0)
                    {
                        string[] split = Attribute.Split(new Char[] { '|' });
                        if (split.Count() >= 2)
                        {
                            string attribute_type = split[1].Trim().Substring(0, 2).ToUpper();
                            if (attribute_type == "ST")
                            {
                                Component.SetAttribute(split[0], split[2]);
                                flag = true;
                            }
                            else if (attribute_type == "IN")
                            {
                                Component.SetAttribute(split[0], Convert.ToInt32(split[2]));
                                flag = true;
                            }
                            else if (attribute_type == "DO")
                            {
                                Component.SetAttribute(split[0], Convert.ToDouble(split[2]));
                                flag = true;
                            }

                        }

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Actual Error: " + ex.ToString(), "Connection Attribute Excel Error");
            }
            return flag;
        }

        public static bool checkAttributeFile(string skmodelpath, string attributefile, string componentnumber)
        {
            bool flag = false;
            string sFolder = skmodelpath + "\\attributes";
            string sFile = attributefile + ".j" + componentnumber.ToString();

            string[] file_Paths = Directory.GetFiles(sFolder, sFile, SearchOption.TopDirectoryOnly);

            if (file_Paths.Count() >= 1)
                flag = true;

            return flag;
        }

        public static LineSegment GetIntersection(Beam Beam1, Beam Beam2)
        {

            Tekla.Structures.Geometry3d.Line Line1 = new Tekla.Structures.Geometry3d.Line(Beam1.StartPoint, Beam1.EndPoint);
            Tekla.Structures.Geometry3d.Line Line2 = new Tekla.Structures.Geometry3d.Line(Beam2.StartPoint, Beam2.EndPoint);
            LineSegment Interseg = null;

            if (Tekla.Structures.Geometry3d.Parallel.LineToLine(Line1, Line2) == false)
            {
                Interseg = Intersection.LineToLine(Line1, Line2);
            }
            return Interseg;
        }

        public static double GetDistance(Tekla.Structures.Geometry3d.Point refPoint, Beam Beam2)
        {
            LineSegment Lineseg = new LineSegment();

            Lineseg.Point1 = Beam2.StartPoint;
            Lineseg.Point2 = Beam2.EndPoint;

            double intDistance = Distance.PointToLineSegment(refPoint, Lineseg);

            return intDistance;
        }

    }

}
