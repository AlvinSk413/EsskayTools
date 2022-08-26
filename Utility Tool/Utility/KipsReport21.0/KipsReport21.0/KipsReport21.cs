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

namespace KipsReport21._0
{
    public partial class KipsReport21 : Form
    {
        public static string skApplicationName = "KipsReport21.0";
        public static string skApplicationVersion = "2201.02";
        public KipsReport21()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "KipsReport", "");
            Model model = new Model();
            Tekla.Structures.Model.UI.ModelObjectSelector modelHandler = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnumerator = modelHandler.GetSelectedObjects();
            label1.Text = "Processing..........";
            label1.Visible = true;

            List<PartInfo5> partInfos = new List<PartInfo5>();
            while(modelObjectEnumerator.MoveNext())
            {
                Part part = modelObjectEnumerator.Current as Part;
                PartInfo5 partInfo = new PartInfo5(part);
                partInfos.Add(partInfo);
            }
            createReport(partInfos);
            label1.Text = "Completed";
            MessageBox.Show("Completed");
            label1.Visible = false;
        }
        public void createReport(List<PartInfo5> partInfos)
        {
            Model model = new Model();
            ModelInfo modelInfo = model.GetInfo();
            System.IO.Directory.CreateDirectory(modelInfo.ModelPath + "\\Automation");
            //string dateTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            //string path = "KipsReport" +"_"+ dateTime.Replace('-','_').Replace(':','_')+".csv";
            string reportPath = modelInfo.ModelPath + "\\Automation\\"+ "KipsReport.csv";
            StreamWriter streamWriter = new StreamWriter(reportPath, false);
            streamWriter.Write("LeftPartName");
            streamWriter.Write(",");
            streamWriter.Write("LeftConnectedIn");
            streamWriter.Write(",");
            streamWriter.Write("LeftPartProfile");
            streamWriter.Write(",");
            streamWriter.Write("LeftConCode");
            streamWriter.Write(",");
            streamWriter.Write("LeftShear");
            streamWriter.Write(",");
            streamWriter.Write("LeftMoment");
            streamWriter.Write(",");
            streamWriter.Write("LeftAxial");
            streamWriter.Write(",");
            //streamWriter.Write("StartTos");
            //streamWriter.Write(",");
            //streamWriter.Write("StartTosDiff");
            //streamWriter.Write(",");
            //streamWriter.Write("StartBoltInfo");
            //streamWriter.Write(",");
            //streamWriter.Write("StartWeldInfo");
            //streamWriter.Write(",");
            streamWriter.Write("Profile");
            streamWriter.Write(",");
            streamWriter.Write("Tos");
            streamWriter.Write(",");
            //streamWriter.Write("EndBoltInfo");
            //streamWriter.Write(",");
            //streamWriter.Write("EndWeldInfo");
            //streamWriter.Write(",");
            streamWriter.Write("RightShear");
            streamWriter.Write(",");
            streamWriter.Write("RightMoment");
            streamWriter.Write(",");
            streamWriter.Write("RightAxial");
            streamWriter.Write(",");
            //streamWriter.Write("EndTos");
            //streamWriter.Write(",");
            //streamWriter.Write("EndTosDiff");
            //streamWriter.Write(",");
            streamWriter.Write("RightConCode");
            streamWriter.Write(",");
            streamWriter.Write("RightPartProfile");
            streamWriter.Write(",");
            streamWriter.Write("RightConnectedIn");
            streamWriter.Write(",");
            streamWriter.Write("RightPartName");
            streamWriter.Write(",");
            streamWriter.Write("ProfileGuid");
            streamWriter.Write(",");
            streamWriter.Write("PartMark");
            streamWriter.Write(",");
            streamWriter.Write("Seq");


            streamWriter.WriteLine();
            foreach (PartInfo5 partInfo in partInfos)
            {
                streamWriter.Write(partInfo.StartConName);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartConnectedIn);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartConProfile);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartConCode);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartKips);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartMoment);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.StartAxial);
                streamWriter.Write(",");
                //streamWriter.Write(partInfo.StartTos);
                //streamWriter.Write(",");
                //streamWriter.Write(partInfo.StartTosDiff);
                //streamWriter.Write(",");
                //streamWriter.Write(partInfo.StartBoltInfo);
                //streamWriter.Write(",");
                //streamWriter.Write(partInfo.StartWeldInfo);
                //streamWriter.Write(",");
                streamWriter.Write(partInfo.Profile);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.Tos);
                streamWriter.Write(",");
                //streamWriter.Write(partInfo.EndBoltInfo);
                //streamWriter.Write(",");
                //streamWriter.Write(partInfo.EndWeldInfo);
                //streamWriter.Write(",");
                streamWriter.Write(partInfo.EndKips);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.EndMoment);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.EndAxial);
                streamWriter.Write(",");
                //streamWriter.Write(partInfo.EndTos);
                //streamWriter.Write(",");
                //streamWriter.Write(partInfo.EndTosDiff);
                //streamWriter.Write(",");
                streamWriter.Write(partInfo.EndConCode);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.EndConProfile);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.EndConnectedIn);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.EndConName);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.Guid);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.PartMark);
                streamWriter.Write(",");
                streamWriter.Write(partInfo.SeqNumber);
                streamWriter.WriteLine();
            }
            streamWriter.Close();
        }
    }
}
