using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures.Model;

namespace ABMLengthVSLiveLength
{
    public class UserFieldInfo
    {
        public string abmLength { get; set; }
        public string abmLengthInMillimeter { get; set; }
        public string guid { get; set; }
        public string userInfo11 { get; set; }
        public UserFieldInfo(Part part)
        {
            string userfield = "";
            part.GetReportProperty("NOTES12", ref userfield);
            this.abmLength = userfield;
            this.abmLengthInMillimeter = convertStringToDouble(this.abmLength).ToString();
            this.guid = "GUID:ID" + part.Identifier.GUID.ToString().ToUpper();
            string userfield11 = "";
            part.GetReportProperty("NOTES11", ref userfield11);
            this.userInfo11 = userfield11;
        }
        private double convertStringToDouble(string ABMlength)
        {
            double feetValue = 0;
            double inchValue = 0;
            double fractionValue = 0;
            bool feet = ABMlength.Contains("'");
            //bool inch = ABMlength.Contains("\"");
            bool fraction = ABMlength.Contains("/");
            if (fraction)
            {
                bool inch = ABMlength.Contains(" ");

                if (feet)
                {
                    string[] feetseparator = new string[1] { "'-" };
                    string[] feetInch = ABMlength.Split(feetseparator, StringSplitOptions.RemoveEmptyEntries);
                    string feetvalueInString = ABMlength.Split(feetseparator, StringSplitOptions.RemoveEmptyEntries)[0];
                    feetValue = Convert.ToDouble(feetvalueInString);
                    if (inch)
                    {
                        //string[] inchseparator = new string[1] { "\"" };
                        string[] inchseparator = new string[1] { " " };
                        string[] inchFractionvalueInString = feetInch[1].Split(inchseparator, StringSplitOptions.RemoveEmptyEntries);
                        string inchValueInString = inchFractionvalueInString[0];
                        inchValue = Convert.ToDouble(inchValueInString);
                        if (fraction)
                        {
                            string[] fractionSeparator = new string[1] { "/" };
                            string[] fractionvalueInString = inchFractionvalueInString[1].Split(fractionSeparator, StringSplitOptions.RemoveEmptyEntries);
                            fractionValue = Math.Round((Convert.ToDouble(fractionvalueInString[0]) / Convert.ToDouble(fractionvalueInString[1].Trim('\"'))), 2);
                        }

                    }


                }
                else
                {
                    string[] inchseparator = new string[1] { "\"" };
                    string[] inchFractionvalueInString = ABMlength.Split(inchseparator, StringSplitOptions.RemoveEmptyEntries);
                    string inchValueInString = inchFractionvalueInString[0];
                    inchValue = Convert.ToDouble(inchValueInString);
                    if (fraction)
                    {
                        string[] fractionSeparator = new string[1] { "/" };
                        string[] fractionvalueInString = inchFractionvalueInString[1].Split(fractionSeparator, StringSplitOptions.RemoveEmptyEntries);
                        fractionValue = Math.Round((Convert.ToDouble(fractionvalueInString[0]) / Convert.ToDouble(fractionvalueInString[1])), 2);
                    }
                }
                double length = 0;
                double feetValueInMillimeter = feetValue * 12 * 25.4;
                double inchValueInMillimeter = inchValue * 25.4;
                double fractionValueInMilliemeter = fractionValue * 25.4;
                length = Math.Round((feetValueInMillimeter + inchValueInMillimeter + fractionValueInMilliemeter), 2);
                return length;
            }
            else
            {
                bool inch = ABMlength.Contains("\"");

                if (feet)
                {
                    string[] feetseparator = new string[1] { "'-" };
                    string[] feetInch = ABMlength.Split(feetseparator, StringSplitOptions.RemoveEmptyEntries);
                    string feetvalueInString = ABMlength.Split(feetseparator, StringSplitOptions.RemoveEmptyEntries)[0];
                    feetValue = Convert.ToDouble(feetvalueInString.Trim('\''));
                    if (inch)
                    {
                        string[] inchseparator = new string[1] { "\"" };
                        //string[] inchseparator = new string[1] { " " };
                        string[] inchFractionvalueInString = feetInch[1].Split(inchseparator, StringSplitOptions.RemoveEmptyEntries);
                        string inchValueInString = inchFractionvalueInString[0];
                        inchValue = Convert.ToDouble(inchValueInString);
                        if (fraction)
                        {
                            string[] fractionSeparator = new string[1] { "/" };
                            string[] fractionvalueInString = inchFractionvalueInString[1].Split(fractionSeparator, StringSplitOptions.RemoveEmptyEntries);
                            fractionValue = Math.Round((Convert.ToDouble(fractionvalueInString[0]) / Convert.ToDouble(fractionvalueInString[1].Trim('\"'))), 2);
                        }

                    }


                }
                else
                {
                    string[] inchseparator = new string[1] { "\"" };
                    string[] inchFractionvalueInString = ABMlength.Split(inchseparator, StringSplitOptions.RemoveEmptyEntries);
                    string inchValueInString = inchFractionvalueInString[0];
                    inchValue = Convert.ToDouble(inchValueInString);
                    if (fraction)
                    {
                        string[] fractionSeparator = new string[1] { "/" };
                        string[] fractionvalueInString = inchFractionvalueInString[1].Split(fractionSeparator, StringSplitOptions.RemoveEmptyEntries);
                        fractionValue = Math.Round((Convert.ToDouble(fractionvalueInString[0]) / Convert.ToDouble(fractionvalueInString[1])), 2);
                    }
                }
                double length = 0;
                double feetValueInMillimeter = feetValue * 12 * 25.4;
                double inchValueInMillimeter = inchValue * 25.4;
                double fractionValueInMilliemeter = fractionValue * 25.4;
                length = Math.Round((feetValueInMillimeter + inchValueInMillimeter + fractionValueInMilliemeter), 2);
                return length;
            }

        }
    }
}
