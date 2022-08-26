using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;

namespace NearFarShearTabFinder
{
    public class SolidPartMaxMin
    {
        public Part Part { get; set; }
        public Point solidMax { get; set; }
        public Point solidMin { get; set; }
        public string location { get; set; }
        public string nearFar { get; set; }
    }
}
