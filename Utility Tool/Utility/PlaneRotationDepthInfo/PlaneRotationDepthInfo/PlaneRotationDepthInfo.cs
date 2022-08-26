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

namespace PlaneRotationDepthInfo
{
    public partial class PlaneRotationDepthInfo : Form
    {
        public static string skApplicationName = "Plane Rotaion Depth Info";
        public static string skApplicationVersion = "2201.02";

        public PlaneRotationDepthInfo()
        {
            InitializeComponent();
        }
            
        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Run", "");
            int ct = 0;

            var progress = new Tekla.Structures.Model.Operations.Operation.ProgressBar();
            bool displayResult = progress.Display(100, "Beam Position Info Extract ", "Processing....", "cancel..", " ");
            List<Report5> reports = new List<Report5>();
            Model model = new Model();
            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnumerator = modelObjectSelector.GetSelectedObjects();
            int totalCount = modelObjectEnumerator.GetSize();
            int ii = 1;
            int percentValue = 100 * ii / totalCount;
            while (modelObjectEnumerator.MoveNext())
            {
                ct++;
                percentValue = 100 * ii / totalCount;
                progress.SetProgress(percentValue.ToString(), 100 * ii / totalCount);
                Part part = modelObjectEnumerator.Current as Part;
                if(part != null)
                {
                    string s = part.GetType().ToString();
                    if(s.Equals("Tekla.Structures.Model.Beam"))
                    {
                        Report5 report = new Report5();
                        report.guid ="GUID:"+ part.Identifier.GUID.ToString().ToUpper();
                        report.Name = part.Name;
                        report.Plane = part.Position.Plane.ToString();
                        report.Rotation = part.Position.Rotation.ToString();
                        report.Depth = part.Position.Depth.ToString();
                        reports.Add(report);
                    }
                }
                ii++;
                percentValue = 100 * ii / totalCount;
                progress.SetProgress(percentValue.ToString(), 100 * ii / totalCount);
            }
            createReport(reports, "PlaneRotationDepthInfo");
            progress.Close();
            MessageBox.Show("Completed");

            skWinLib.worklog(skApplicationName, skApplicationVersion, "Run", ";Total:" + ct.ToString());

        }
        public void createReport(List<Report5> missing, string location)
        {
            Model model = new Model();
            ModelInfo modelInfo = model.GetInfo();
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation");
            string reportPath = modelInfo.ModelPath + "\\Automation\\" + location + ".csv";
            StreamWriter streamWriter = new StreamWriter(reportPath, false);
            streamWriter.Write("Name");
            streamWriter.Write(",");
            streamWriter.Write("Plane");
            streamWriter.Write(",");
            streamWriter.Write("Rotation");
            streamWriter.Write(",");
            streamWriter.Write("Depth");
            streamWriter.Write(",");
            streamWriter.Write("Guid");
            streamWriter.WriteLine();
            foreach (Report5 obj in missing)
            {
                streamWriter.Write(obj.Name);
                streamWriter.Write(",");
                streamWriter.Write(obj.Plane);
                streamWriter.Write(",");
                streamWriter.Write(obj.Rotation);
                streamWriter.Write(",");
                streamWriter.Write(obj.Depth);
                streamWriter.Write(",");
                streamWriter.Write(obj.guid);
                streamWriter.WriteLine();
            }
            streamWriter.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //this.TopMost = checkBox1.Checked;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
