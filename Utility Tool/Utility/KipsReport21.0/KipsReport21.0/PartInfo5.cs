using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace KipsReport21._0
{
    public class PartInfo5
    {
        public string Guid { get; set; }
        public string Profile { get; set; }
        public string StartConCode { get; set; }
        public string StartConProfile { get; set; }
        public string StartConName { get; set; }
        public string StartKips { get; set; }
        public string EndConCode { get; set; }
        public string EndConProfile { get; set; }
        public string EndConName { get; set; }
        public string EndKips { get; set; }
        public string StartConnectedIn { get; set; }
        public string EndConnectedIn { get; set; }
        public string SeqNumber { get; set; }
        public string StartTos { get; set; }
        public string EndTos { get; set; }
        public string Tos { get; set; }
        public string StartTosDiff { get; set; }
        public string EndTosDiff { get; set; }
        public string StartBoltInfo { get; set; }
        public string EndBoltInfo { get; set; }
        public string StartWeldInfo { get; set; }
        public string EndWeldInfo { get; set; }
        public string PartMark { get; set; }
        public string StartMoment { get; set; }
        public string EndMoment{ get; set; }
        public string StartAxial { get; set; }
        public string EndAxial { get; set; }
        public PartInfo5(Part part)
        {
            this.Guid ="GUID:ID"+ part.Identifier.GUID.ToString().ToUpper();
            this.Profile = part.Profile.ProfileString;
            this.StartKips = getReactionStart(part);
            this.EndKips = getReactionEnd(part);
            this.SeqNumber = getPhase(part);
            this.PartMark = part.GetPartMark();
            this.StartMoment = getMomentStart(part);
            this.EndMoment = getMomentEnd(part);
            this.StartAxial = getAxialStart(part);
            this.EndAxial = getAxialEnd(part);


            string startConnectionProfile = "";
            string startConnectionCode = "";
            string endConnectionProfile = "";
            string endConnectionCode = "";
            string startConnectionName = "";
            string endConnectionName = "";
            string startConnected= "-";
            string endConnectedIn = "-";
            string startTos = "-";
            string endTos = "-";
            string Tos = "-";
            string startTosDiff = "-";
            string endTosDiff = "-";
            string startBoltInfo = "-";
            string endBoltInfo = "-";
            string startWeldInfo = "-";
            string endWeldInfo = "-";
            Model model = new Model();
            ModelObjectEnumerator modelObjectEnumerator = part.GetComponents();
            while(modelObjectEnumerator.MoveNext())
            {
                var obj = modelObjectEnumerator.Current;
                if(obj.GetType().Equals(typeof(Connection)))
                {
                    Connection connection = obj as Connection;
                    ArrayList arrayList = connection.GetSecondaryObjects();
                    Part secPart = arrayList[0] as Part;
                    if(secPart.Identifier.GUID.Equals(part.Identifier.GUID))
                    {
                        model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                        model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(part.GetCoordinateSystem()));
                        string conLocation = getConLocation(connection, part);
                        
                        //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(connection.GetCoordinateSystem()));
                        //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                        if (conLocation==Constant.Location.START.ToString())
                        {
                            startConnectionCode = connection.Code;
                            Part mainPart = connection.GetPrimaryObject() as Part;
                            startConnectionProfile = mainPart.Profile.ProfileString;
                            startConnectionName = mainPart.Name;
                             startBoltInfo = getBoltInfo(connection);
                            startWeldInfo = getWeldInfo(connection);
                            //Point point = (mainPart as Beam).GetCoordinateSystem().Origin;
                            //Matrix matrix = MatrixFactory.ToCoordinateSystem(part.GetCoordinateSystem());
                            //Point CONVpOINT = matrix.Transform(point);
                            ////CoordinateSystem coordinateSystem = mainPart.GetCoordinateSystem();
                            //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                            //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(mainPart.GetCoordinateSystem()));
                            if (startConnectionName.Contains("COLUMN"))
                            {
                                Vector yVector = mainPart.GetCoordinateSystem().AxisY;
                                yVector.Normalize();
                                if(yVector.X !=0)
                                {
                                    startConnected = "FLANGE";
                                }
                                else if (yVector.Z != 0)
                                {
                                    startConnected = "WEB";
                                }
                                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                            }
                            else
                            {
                                try
                                {
                                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                                    Point partZ = (part as Beam).StartPoint;
                                    Point startPartZ = (mainPart as Beam).StartPoint;
                                    bool check = isSlope(mainPart);

                                    if (check == false)
                                    {
                                        double tosDistanceInMillimeter = Tekla.Structures.Geometry3d.Distance.PointToPoint(partZ, startPartZ);
                                        Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(tosDistanceInMillimeter);
                                        startTosDiff = distance1.ToFractionalFeetAndInchesString();
                                        Tekla.Structures.Datatype.Distance distance2 = new Tekla.Structures.Datatype.Distance(startPartZ.Z);
                                        startTos = getTopLevel(mainPart);
                                        startTosDiff = getStartTosDiff(part, mainPart);

                                    }
                                    else
                                    {
                                        startTosDiff = "Check";
                                    }
                                }
                                catch
                                {

                                }
                               
                               

                            }
                            

                        }
                        else if(conLocation == Constant.Location.END.ToString())
                        {
                            
                            endConnectionCode = connection.Code;
                            Part mainPart = connection.GetPrimaryObject() as Part;
                            endConnectionProfile = mainPart.Profile.ProfileString;
                            endConnectionName = mainPart.Name;
                            endBoltInfo = getBoltInfo(connection);
                            endWeldInfo = getWeldInfo(connection);
                            if (endConnectionName.Contains("COLUMN"))
                            {
                                Vector yVector = mainPart.GetCoordinateSystem().AxisY;
                                yVector.Normalize();
                                if (yVector.X != 0)
                                {
                                    endConnectedIn = "FLANGE";
                                }
                                else if (yVector.Z != 0)
                                {
                                    endConnectedIn = "WEB";
                                }
                                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                            }
                            else
                            {
                                try
                                {


                                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                                    Point partZ = (part as Beam).EndPoint;
                                    Point startPartZ = (mainPart as Beam).StartPoint;
                                    bool check = isSlope(mainPart);
                                    if (check == false)
                                    {
                                        double tosDistanceInMillimeter = Tekla.Structures.Geometry3d.Distance.PointToPoint(partZ, startPartZ);
                                        Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(tosDistanceInMillimeter);
                                        endTosDiff = distance1.ToFractionalFeetAndInchesString();
                                        Tekla.Structures.Datatype.Distance distance2 = new Tekla.Structures.Datatype.Distance(startPartZ.Z);
                                        endTos = getTopLevel(mainPart);
                                        endTosDiff = getEndTosDiff(part, mainPart);


                                    }
                                    else
                                    {
                                        endTosDiff = "Check";
                                    }
                                }
                                catch
                                {

                                }
                                

                            }

                        }
                       


                    }

                }
            }
            Point partZValue = (part as Beam).EndPoint;
            Tekla.Structures.Datatype.Distance distance3 = new Tekla.Structures.Datatype.Distance(partZValue.Z);
            Tos = getTopLevel(part);
            this.StartConCode = startConnectionCode;
            this.EndConCode = endConnectionCode;
            this.StartConProfile = startConnectionProfile;
            this.EndConProfile = endConnectionProfile;
            this.StartConName = startConnectionName;
            this.EndConName = endConnectionName;
            this.StartConnectedIn = startConnected;
            this.EndConnectedIn = endConnectedIn;
            this.StartTosDiff = startTosDiff;
            this.EndTosDiff = endTosDiff;
            this.StartTos = startTos;
            this.EndTos = endTos;
            this.Tos = Tos;
            this.StartBoltInfo = startBoltInfo;
            this.EndBoltInfo = endBoltInfo;
            this.StartWeldInfo = startWeldInfo;
            this.EndWeldInfo = endWeldInfo;
        }
        private string getConLocation(Connection connection,Part part)
        {
            string conLocation = "";
            Tekla.Structures.Geometry3d.Point point = connection.GetCoordinateSystem().Origin;
            double length = 0;
            part.GetReportProperty("LENGTH", ref length);
            if (point.X < length / 2)
            {
                conLocation = Constant.Location.START.ToString();
            }
            else
            {
                conLocation = Constant.Location.END.ToString();
            }
            return conLocation;
        }
        private string getReactionStart(Part beam)
        {
            double value = 0;
            beam.GetReportProperty("USERDEFINED.shear1", ref value);
            value = value / 4448.22;
            long a = Convert.ToInt64(value);
            return a.ToString();
        }
        private string getReactionEnd(Part beam)
        {
            double value = 0;
            beam.GetReportProperty("USERDEFINED.shear2", ref value);
            value = value / 4448.22;
            long a = Convert.ToInt64(value);
            return a.ToString();
        }
        private string getPhase(Part beam)
        {
            int value =0;
            beam.GetReportProperty("PHASE", ref value);
          
            return value.ToString();
        }
        private int getStartZ(Part part)
        {
            double startZ = 0;
            part.GetReportProperty("START_Z", ref startZ);

            return Convert.ToInt32(startZ);
        }
        private int getEndZ(Part part)
        {
            double endZ = 0;
            part.GetReportProperty("END_Z", ref endZ);

            return Convert.ToInt32(endZ);
        }
        private bool isSlope(Part part)
        {
            bool flag = true;
            int start = getStartZ(part);
            int end = getEndZ(part);
            if(start == end)
            {
                flag = false;
            }
            return flag;
            
        }
        private string getTopLevel(Part beam)
        {
            string value = "";
            beam.GetReportProperty("TOP_LEVEL", ref value);

            return value;
        }
        private string getStartTosDiff(Part part,Part mainPart)
        {
            string value = "";
            double z1 = 0;
            part.GetReportProperty("START_Z_IN_WORK_PLANE", ref z1);
            double z2 = 0;
            mainPart.GetReportProperty("START_Z_IN_WORK_PLANE", ref z2);
            Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(z1);
            z1 =  distance1.Millimeters;
            Tekla.Structures.Datatype.Distance distance2 = new Tekla.Structures.Datatype.Distance(z2);
            z2 = distance2.Millimeters;
            double diff = z2 - z1;

            Tekla.Structures.Datatype.Distance diffCon = new Tekla.Structures.Datatype.Distance(diff);
            value = diffCon.ToFractionalFeetAndInchesString();

            return value;

        }
        private string getEndTosDiff(Part part, Part mainPart)
        {
            string value = "";
            double z1 = 0;
            part.GetReportProperty("END_Z_IN_WORK_PLANE", ref z1);
            double z2 = 0;
            mainPart.GetReportProperty("END_Z_IN_WORK_PLANE", ref z2);
            Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(z1);
            z1 = distance1.Millimeters;
            Tekla.Structures.Datatype.Distance distance2 = new Tekla.Structures.Datatype.Distance(z2);
            z2 = distance2.Millimeters;
            double diff = z2 - z1;
            Tekla.Structures.Datatype.Distance diffCon = new Tekla.Structures.Datatype.Distance(diff);
            value = diffCon.ToFractionalFeetAndInchesString();

            return value;

        }
        private string getBoltInfo(Connection connection)
        {
            string boltInfo = "";
            ModelObjectEnumerator modelObjectEnumerator = connection.GetChildren();
            while(modelObjectEnumerator.MoveNext())
            {
                BoltGroup boltGroup = modelObjectEnumerator.Current as BoltGroup;
                if(boltGroup != null)
                {
                    double value = 0;
                    boltGroup.GetReportProperty("DIAMETER", ref value);
                    Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(value);
                    boltInfo = distance1.ToFractionalFeetAndInchesString();
                    boltInfo = boltInfo + " " + boltGroup.BoltStandard;
                }

            }
            return boltInfo;
        }
        private string getWeldInfo(Connection connection)
        {
            string weldInfo = "";
            ModelObjectEnumerator modelObjectEnumerator = connection.GetChildren();
            while (modelObjectEnumerator.MoveNext())
            {
                BaseWeld baseWeld = modelObjectEnumerator.Current as BaseWeld;
                if (baseWeld != null)
                {
                    double above = baseWeld.SizeAbove;
                    double below = baseWeld.SizeBelow;
                    Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(above);
                    Tekla.Structures.Datatype.Distance distance2 = new Tekla.Structures.Datatype.Distance(below);
                    weldInfo = distance1.ToFractionalFeetAndInchesString() + "&" + distance2.ToFractionalFeetAndInchesString();
                    //double value = 0;
                    //baseWeld.GetReportProperty("DIAMETER", ref value);
                    //Tekla.Structures.Datatype.Distance distance1 = new Tekla.Structures.Datatype.Distance(value);
                    //boltInfo = distance1.ToFractionalFeetAndInchesString();
                    //boltInfo = boltInfo + " " + boltGroup.BoltStandard;
                }

            }
            return weldInfo;
        }
        private string getMomentStart(Part beam)
        {
            double value = 0;
            beam.GetReportProperty("USERDEFINED.moment1", ref value);
            value = value / 1355.82;
            long a = Convert.ToInt64(value);
            return a.ToString();
        }
        private string getMomentEnd(Part beam)
        {
            double value = 0;
            beam.GetReportProperty("USERDEFINED.moment2", ref value);
            value = value / 1355.82;
            long a = Convert.ToInt64(value);
            return a.ToString();
        }
        private string getAxialStart(Part beam)
        {
            string value = "";
            beam.GetReportProperty("USERDEFINED.EDAS_USER_FIELD_1", ref value);
            return value;
        }
        private string getAxialEnd(Part beam)
        {
            string value = "";
            beam.GetReportProperty("USERDEFINED.EDAS_USER_FIELD_2", ref value);
            return value;
        }
    }
}
