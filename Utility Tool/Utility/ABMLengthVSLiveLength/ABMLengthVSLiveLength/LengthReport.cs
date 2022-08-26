using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Datatype;

namespace ABMLengthVSLiveLength
{
    public class LengthReport
    {
        public string guid { get; set; }
        public string liveLength { get; set; }
        public string abmLength { get; set; }
        public string difference { get; set; }
        public string profile { get; set; }
        public string prelimMark { get; set; }
        public string lengthDiffBy { get; set; }
        double exceedValue = 0; /*3.18;/*1/8*/
        double shortValue = 0;/*50.8;/*2"*/
        public string userInfo11 { get; set; }
        public LengthReport(LiveModelInfo liveModelInfo, UserFieldInfo userFieldInfo)
        {
            this.guid = liveModelInfo.guid;
            this.liveLength = liveModelInfo.liveLength;
            this.abmLength = userFieldInfo.abmLength;
            double diff = Math.Round((Convert.ToDouble(liveModelInfo.liveLengthInMillimeter) - Convert.ToDouble(userFieldInfo.abmLengthInMillimeter)), 2);
            //Distance differenceLength = new Distance(diff);
            //string differenceInLength = differenceLength.ToFractionalFeetAndInchesString();
            string differenceInLength = Util1.GetFeetAndInches(diff, 16);
            this.difference = differenceInLength;
            this.profile = liveModelInfo.liveProfile;
            this.prelimMark = liveModelInfo.prelimmark;
            if (diff > exceedValue)
            {
                lengthDiffBy = "EXCESS";

            }
            else if ((diff < -shortValue) && (diff < 0))
            {
                if(diff<-152.4)
                {
                    lengthDiffBy = "SHORT";
                }
                else
                {

                }
                
            }
            else
            {
                lengthDiffBy = "NOCHANGE";
            }
            this.userInfo11 = userFieldInfo.userInfo11;


        }
    }
}
