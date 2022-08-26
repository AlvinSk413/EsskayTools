using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace ABMLengthVSLiveLength
{
    public partial class ABMLengthVSLiveLength : Form
    {
        public static string skApplicationName = "ABMLengthVsLiveLength";
        public static string skApplicationVersion = "2201.02";
        public ABMLengthVSLiveLength()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Compare", "");
            label1.Visible = true;
            Model model = new Model();
            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnumerator = modelObjectSelector.GetSelectedObjects();
            List<LengthReport> lengthReports = new List<LengthReport>();
            List<LengthReport> lengthReportExcess = new List<LengthReport>();
            List<LengthReport> lengthReportShort = new List<LengthReport>();
            List<LengthReport> lengthReportNochange = new List<LengthReport>();
            List<LengthReport> lengthReportExcludedContent= new List<LengthReport>();
            List<LiveModelInfo> newlyAdded = new List<LiveModelInfo>();
            while (modelObjectEnumerator.MoveNext())
            {
                Part part = modelObjectEnumerator.Current as Part;
                if (part != null)
                {
                    string ABMlength = "";
                    part.GetReportProperty("NOTES12", ref ABMlength);

                    if (ABMlength != "")
                    {
                        UserFieldInfo userFieldInfo = new UserFieldInfo(part);
                        LiveModelInfo liveModelInfo = new LiveModelInfo(part);
                        LengthReport lengthReport = new LengthReport(liveModelInfo, userFieldInfo);
                        lengthReports.Add(lengthReport);
                        if (lengthReport.lengthDiffBy == "EXCESS")
                        {
                            lengthReportExcess.Add(lengthReport);
                        }
                        else if (lengthReport.lengthDiffBy == "SHORT")
                        {
                            lengthReportShort.Add(lengthReport);
                        }
                        else if(lengthReport.lengthDiffBy == "NOCHANGE")
                        {
                            lengthReportNochange.Add(lengthReport);
                        }
                        else
                        {
                            lengthReportExcludedContent.Add(lengthReport);
                        }

                    }
                    else
                    {
                        LiveModelInfo liveModelInfo = new LiveModelInfo(part);
                        newlyAdded.Add(liveModelInfo);
                    }
                }

            }
            if(lengthReportExcess.Count>0)
            {
                createReport(lengthReportExcess, "ABMLength_Vs_LiveLengthExcess_" + DateandTime());
            }
            if(lengthReportShort.Count>0)
            {
                createReport(lengthReportShort, "ABMLength_Vs_LiveLengthShort_" + DateandTime());
            }
            if(lengthReportNochange.Count>0)
            {
                createReport(lengthReportNochange, "ABMLength_Vs_LiveLengthNoChange_" + DateandTime());
            }
            if (lengthReportExcludedContent.Count > 0)
            {
                createReport(lengthReportExcludedContent, "ABMLength_Vs_LiveLengthExcluded_" + DateandTime());
            }
            if(newlyAdded.Count>0)
            {
                createReport(newlyAdded, "NewlyAdded_" + DateandTime());
            }



            label1.Visible = false;
            MessageBox.Show("Completed");
        }
        public void createReport(List<LengthReport> lengthReports, string filename)
        {
            Model model = new Model();
            ModelInfo modelInfo = model.GetInfo();
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation");
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation\\ABMvsLive");
            string reportPath = modelInfo.ModelPath + "\\Automation\\ABMvsLive\\" + filename + ".csv";
            StreamWriter streamWriter = new StreamWriter(reportPath, false);
            streamWriter.Write("GUID");
            streamWriter.Write(",");
            streamWriter.Write("LiveProfile");
            streamWriter.Write(",");
            streamWriter.Write("ABMLength");
            streamWriter.Write(",");
            streamWriter.Write("LiveLength");
            streamWriter.Write(",");
            streamWriter.Write("Diff");
            streamWriter.Write(",");
            streamWriter.Write("PrelimMark");
            streamWriter.Write(",");
            streamWriter.Write("DiffSummary");
            streamWriter.Write(",");
            streamWriter.Write("Notes11ByUser");
            streamWriter.WriteLine();
            foreach (LengthReport obj in lengthReports)
            {
                streamWriter.Write(obj.guid);
                streamWriter.Write(",");
                streamWriter.Write(obj.profile);
                streamWriter.Write(",");
                streamWriter.Write(obj.abmLength);
                streamWriter.Write(",");
                streamWriter.Write(obj.liveLength);
                streamWriter.Write(",");
                streamWriter.Write(obj.difference);
                streamWriter.Write(",");
                streamWriter.Write(" "+obj.prelimMark);
                streamWriter.Write(",");
                streamWriter.Write(obj.lengthDiffBy);
                streamWriter.Write(",");
                streamWriter.Write(obj.userInfo11);
                streamWriter.WriteLine();
            }
            streamWriter.Close();
        }
        public void createReport(List<LiveModelInfo> lengthReports, string filename)
        {
            Model model = new Model();
            ModelInfo modelInfo = model.GetInfo();
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation");
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation\\ABMvsLive");
            string reportPath = modelInfo.ModelPath + "\\Automation\\ABMvsLive\\" + filename + ".csv";
            StreamWriter streamWriter = new StreamWriter(reportPath, false);
            streamWriter.Write("GUID");
            streamWriter.Write(",");
            streamWriter.Write("LiveProfile");
            streamWriter.Write(",");
            streamWriter.Write("LiveLength");
            streamWriter.Write(",");
            streamWriter.Write("Prelim Mark");
            streamWriter.WriteLine();
            foreach (LiveModelInfo obj in lengthReports)
            {
                streamWriter.Write(obj.guid);
                streamWriter.Write(",");
                streamWriter.Write(obj.liveProfile);
                streamWriter.Write(",");
                streamWriter.Write(obj.liveLength);
                streamWriter.Write(",");
                streamWriter.Write(" " + obj.prelimmark);
                streamWriter.WriteLine();
            }
            streamWriter.Close();
        }
        public static string DateandTime()
        {
            DateTime dateTime = System.DateTime.Now;
            return dateTime.ToString("g").Replace(":", "_").Replace('/', '_');
        }
    }
}
