using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Datatype;
using Tekla.Structures.Model;

namespace ABMLengthVSLiveLength
{
    public class LiveModelInfo
    {
        public string liveLength { get; set; }
        public string liveProfile { get; set; }
        public string liveMaterialGrade { get; set; }
        public string liveLengthInMillimeter { get; set; }
        public string guid { get; set; }
        public string prelimmark { get; set; }
        public string phase { get; set; }
        public LiveModelInfo(Part part)
        {
            this.liveProfile = getProfileString(part);
            this.liveMaterialGrade = part.Material.MaterialString;
            this.liveLength = getLengthString(part);
            this.liveLengthInMillimeter = getLength(part).ToString();
            this.guid = "GUID:ID" + part.Identifier.GUID.ToString().ToUpper();
            this.prelimmark = getPrelimMarkString(part);
            this.phase = getPhase(part);
            //string s = getProfileString(part);
        }
        private double getLength(Part part)
        {
            double length = 0;
            part.GetReportProperty("LENGTH_NET", ref length);
            double netLength = Math.Round(length, 2);
            return netLength;
        }
        private string getLengthString(Part part)
        {
            double length = 0;
            part.GetReportProperty("LENGTH_NET", ref length);
            double netLength = Math.Round(length, 2);
            //Distance distance = new Distance(netLength);            
            //string s = distance.ToFractionalFeetAndInchesString();
            string s = Util1.GetFeetAndInches(netLength, 16);
            return s;
        }
        private string getProfileString(Part part)
        {
            string profile = "";
            part.GetReportProperty("PROFILE", ref profile);

            return profile;
        }
        private string getPrelimMarkString(Part part)
        {
            string profile = "";
            part.GetReportProperty("USERDEFINED.PRELIM_MARK", ref profile);

            return profile;
        }
        private string getPhase(Part part)
        {
            int profile = 0;
            part.GetReportProperty("PHASE", ref profile);

            return profile.ToString();
        }
    }
}
