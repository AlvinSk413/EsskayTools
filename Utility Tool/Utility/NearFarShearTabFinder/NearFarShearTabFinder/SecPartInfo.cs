using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace NearFarShearTabFinder
{
    class SecPartInfo
    {
        public Part Part { get; set; }
        public Vector VectorX { get; set; }
        public Vector VectorY { get; set; }
        public Vector VectorZ { get; set; }
        public Point startPoint { get; set; }
        public List<SecPartInfo> SecPartInfos { get; set; }
        public SecPartInfo(Part part,Model model)
        {
            double len = 0;
            part.GetReportProperty("LENGTH", ref len);
            double halfLen = Math.Round(len / 2, 2); 

            List<Part> parts1 = new List<Part>();
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(part.GetCoordinateSystem()));
            CoordinateSystem coordinateSystem = part.GetCoordinateSystem();
            Vector vectorX = coordinateSystem.AxisX;
            Vector vectorY = coordinateSystem.AxisY;
            Vector vectorZ = vectorY.Cross(vectorX);
            Vector partVectorZNormal= vectorZ.GetNormal();

            List<Part> parts = new List<Part>();
            //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
            //model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(part.GetCoordinateSystem()));
            ModelObjectEnumerator boltEnum = part.GetBolts();
            while(boltEnum.MoveNext())
            {
                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                BoltGroup boltGroup = boltEnum.Current as BoltGroup;
                string material = boltGroup.BoltStandard;
                if((material.ToUpper().Contains("SC"))|| (material.ToUpper().Contains("A490X_TC")))
                {
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(boltGroup.GetCoordinateSystem()));
                    CoordinateSystem boltcoordinateSystem = boltGroup.GetCoordinateSystem();
                    Vector boltVectorX = boltcoordinateSystem.AxisX;
                    Vector boltVectorY = boltcoordinateSystem.AxisY;
                    Vector boltVectorZ = boltVectorY.Cross(boltVectorX);
                    Vector boltVectorZNormal = boltVectorZ.GetNormal();
                    if ((Math.Abs(boltVectorZNormal.Z) == Math.Abs(partVectorZNormal.Z))
                        && (Math.Abs(boltVectorZNormal.X) == Math.Abs(partVectorZNormal.X))
                        && (Math.Abs(boltVectorZNormal.Y) == Math.Abs(partVectorZNormal.Y)))
                    {
                        Part partToBeBolted = boltGroup.PartToBeBolted;
                        Part partToBoltTo = boltGroup.PartToBoltTo;
                        if (!partToBeBolted.Identifier.GUID.ToString().Equals(part.Identifier.GUID.ToString()))
                        {
                            parts1.Add(partToBeBolted);
                        }
                        else if (!partToBoltTo.Identifier.GUID.ToString().Equals(part.Identifier.GUID.ToString()))
                        {
                            parts1.Add(partToBoltTo);
                        }
                    }

                }
                

            }
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
            Matrix matrix = MatrixFactory.ToCoordinateSystem(part.GetCoordinateSystem());
            List<SolidPartMaxMin> solidPartMaxMins = new List<SolidPartMaxMin>();
            foreach(Part part1 in parts1)
            {
                Point max= part1.GetSolid().MaximumPoint;
                Point min = part1.GetSolid().MinimumPoint;
                SolidPartMaxMin solidPartMaxMin = new SolidPartMaxMin();
                solidPartMaxMin.Part = part1;
                solidPartMaxMin.solidMin =matrix.Transform( min);
                solidPartMaxMin.solidMax = matrix.Transform(max);
                if((solidPartMaxMin.solidMin.X<halfLen)&& (solidPartMaxMin.solidMax.X < halfLen))
                {
                    solidPartMaxMin.location = "START";
                }
                else if ((solidPartMaxMin.solidMin.X > halfLen)&& (solidPartMaxMin.solidMax.X > halfLen))
                {
                    solidPartMaxMin.location = "END";
                }
                if ((solidPartMaxMin.solidMin.Z < 0) && (solidPartMaxMin.solidMax.Z < 0))
                {
                    solidPartMaxMin.nearFar = "FAR";
                }
                else if ((solidPartMaxMin.solidMin.Z > 0) && (solidPartMaxMin.solidMax.Z > 0))
                {
                    solidPartMaxMin.nearFar = "NEAR";
                }
                solidPartMaxMins.Add(solidPartMaxMin);
            }
            string notes6 = "";
            if(solidPartMaxMins.Count>0)
            {
                foreach(SolidPartMaxMin solidPartMaxMin in solidPartMaxMins)
                {
                    if(solidPartMaxMin.location=="START")
                    {
                        if(solidPartMaxMin.nearFar == "NEAR")
                        {
                            part.SetUserProperty("NOTES6", "NO PAINT (N/S) OF WEB");
                        }
                        else if(solidPartMaxMin.nearFar =="FAR")
                        {
                            part.SetUserProperty("NOTES6", "NO PAINT (F/S) OF WEB");
                        }
                    }
                    else if(solidPartMaxMin.location =="END")
                    {
                        if (solidPartMaxMin.nearFar == "NEAR")
                        {
                            part.SetUserProperty("NOTES7", "NO PAINT (N/S) OF WEB");
                        }
                        else if (solidPartMaxMin.nearFar == "FAR")
                        {
                            part.SetUserProperty("NOTES7", "NO PAINT (F/S) OF WEB");
                        }
                    }

                    //if(notes6 =="")
                    //{
                    //    notes6 = solidPartMaxMin.location + "_" + solidPartMaxMin.nearFar;
                    //}
                    //else
                    //{
                    //    notes6 = notes6+"_"+solidPartMaxMin.location + "_" + solidPartMaxMin.nearFar;
                    //}
                }
                //part.SetUserProperty("NOTES6", notes6);
            }
            else
            {
                part.SetUserProperty("NOTES6", "");

            }
        }
    }
}
