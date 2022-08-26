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

namespace PackageChecker
{
    public partial class PackageChecker : Form
    {
        public static string skApplicationName = "PackageChecker";
        public static string skApplicationVersion = "2201.02";

        public PackageChecker()
        {
            InitializeComponent();
        }
        public List<PartInfo4> partInfos { get; set; }
        public List<string> fileName { get; set; }
        public List<string> dxfFileName { get; set; }
        public List<string> reportFileName { get; set; } 
        public string kssFilePath { get; set; }
        public string folderPath { get; set; } = "";
        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Import KSS", "");
            int ct = 0;

            if (comboBox1.SelectedItem != null)
            {
                if (comboBox1.SelectedItem.ToString().Equals("IFF"))
                {
                    IFFImport();

                }
                else if (comboBox1.SelectedItem.ToString().Equals("IFA"))
                {
                    IFAImport();
                }

            }
            else
            {
                MessageBox.Show("---Select PACKAGE TYPE---");
            }

            skWinLib.worklog(skApplicationName, skApplicationVersion, "Impoprt KSS", ";Total:" + ct.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Select Folder", "");
            int ct = 0;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (richTextBox1.Text == "")
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    this.folderPath = folderBrowserDialog.SelectedPath;
                    label1.Text = this.folderPath;
                    label1.Visible = true;
                }
                else
                {
                    label1.Text = "Select Folder";
                    label1.Visible = true;
                }
            }
            else
            {
                this.folderPath = richTextBox1.Text;
                label1.Text = this.folderPath;
                label1.Visible = true;
            }
            skWinLib.worklog(skApplicationName, skApplicationVersion, "Select Folder", ";Total:" + ct.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "check", "");
            int ct = 0;

            if (comboBox1.SelectedItem != null)
            {
                if (comboBox1.SelectedItem.ToString().Equals("IFF"))
                {
                    IFFCheck();

                }
                else if (comboBox1.SelectedItem.ToString().Equals("IFA"))
                {
                    IFACheck();
                }

            }
            else
            {
                MessageBox.Show("---Select PACKAGE TYPE---");
            }
            skWinLib.worklog(skApplicationName, skApplicationVersion, "check", ";Total:" + ct.ToString());
        }
        public void createReport(List<string> missing,string name)
        {
            StreamWriter streamWriter = new StreamWriter(this.folderPath + "\\" + name+".csv", false);
            foreach (string customFileInfos2 in missing)
            {
                streamWriter.Write(customFileInfos2);
                streamWriter.WriteLine();
            }
            streamWriter.Close();

        }
        public void IFFCheck()
        {
            if (this.folderPath != "")
            {
                List<customFileInfos> customFileInfos = new List<customFileInfos>();
                string[] files = Directory.GetFiles(this.folderPath, "***.***", SearchOption.AllDirectories);
                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    string fileName = fileInfo.Name;
                    string fileextension = fileInfo.Extension;
                    customFileInfos customFileInfo = new customFileInfos();
                    customFileInfo.fileName = fileName;
                    customFileInfo.fileExtension = fileextension;
                    customFileInfos.Add(customFileInfo);
                }
                IEnumerable<IGrouping<string, customFileInfos>> grouppingFilesByExtension = customFileInfos.GroupBy(x => x.fileExtension);
                List<string> pdf = new List<string>();
                List<string> dxf = new List<string>();
                List<string> ncFiles = new List<string>();
                List<string> others = new List<string>();
                foreach (IGrouping<string, customFileInfos> customFileInfos1 in grouppingFilesByExtension)
                {
                    if (customFileInfos1.Key.Equals(".pdf"))
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            string[] fileName = customFileInfos2.fileName.Split('.');
                            string name = fileName[0].TrimStart().TrimEnd();
                            pdf.Add(name);
                        }
                    }
                    else if (customFileInfos1.Key.Equals(".dxf"))
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            dxf.Add(customFileInfos2.fileName.Replace(".dxf", ""));
                        }

                    }
                    else if (customFileInfos1.Key.Equals(".nc1"))
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            ncFiles.Add(customFileInfos2.fileName.Replace(".nc1", ""));
                        }

                    }
                    else
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            others.Add(customFileInfos2.fileName.ToUpper());
                        }

                    }
                }
                bool ptPt = others.Any(x => x.Contains("POINT_TO_POINT"));
                bool fieldBoltList = others.Any(x => x.Contains("FIELD BOLT LIST"));
                bool drawingList = others.Any(x => x.Contains("DRAWING LIST"));
                bool mis = others.Any(x => x.Contains("MIS"));
                List<string> missingReport = new List<string>();
                if (!ptPt)
                {
                    missingReport.Add("POINT_TO_POINT");
                }
                if (!fieldBoltList)
                {
                    missingReport.Add("FIELD BOLT LIST");
                }
                if (!drawingList)
                {
                    missingReport.Add("DRAWING LIST");
                }
                if (!mis)
                {
                    missingReport.Add("MIS");
                }
                List<string> duplicate = pdf.GroupBy(x => x).Where(g => g.Count() > 1).Select(x => x.Key).ToList();


                List<string> missingPdf = fileName.Except(pdf).ToList();
                List<string> missingNC = fileName.Except(ncFiles).ToList();
                List<string> missingDXF = dxfFileName.Except(dxf).ToList();
                createReport(missingPdf, "PDF_missing");
                createReport(missingNC, "NC_missing");
                createReport(missingDXF, "DXF_missing");
                createReport(missingReport, "Report_missing");
                createReport(duplicate, "duplicatePdf");
                MessageBox.Show("Completed");
                label4.Visible = false;
            }
            else
            {
                label1.Text = "Select Folder";
                label1.Visible = true;
            }
        }
        public void IFACheck()
        {
            if (this.folderPath != "")
            {
                List<customFileInfos> customFileInfos = new List<customFileInfos>();
                string[] files = Directory.GetFiles(this.folderPath, "***.***", SearchOption.AllDirectories);
                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    string fileName = fileInfo.Name;
                    string fileextension = fileInfo.Extension;
                    customFileInfos customFileInfo = new customFileInfos();
                    customFileInfo.fileName = fileName;
                    customFileInfo.fileExtension = fileextension;
                    customFileInfos.Add(customFileInfo);
                }
                IEnumerable<IGrouping<string, customFileInfos>> grouppingFilesByExtension = customFileInfos.GroupBy(x => x.fileExtension);
                List<string> pdf = new List<string>();
                List<string> dxf = new List<string>();
                List<string> ncFiles = new List<string>();
                List<string> others = new List<string>();
                foreach (IGrouping<string, customFileInfos> customFileInfos1 in grouppingFilesByExtension)
                {
                    if (customFileInfos1.Key.Equals(".pdf"))
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            string[] fileName = customFileInfos2.fileName.Split('.');
                            string name = fileName[0].TrimStart().TrimEnd();
                            pdf.Add(name);
                        }
                    }
                    else
                    {
                        foreach (customFileInfos customFileInfos2 in customFileInfos1)
                        {
                            others.Add(customFileInfos2.fileName.ToUpper());
                        }

                    }
                }
                bool mis = others.Any(x => x.Contains("MIS"));
                bool drawingList = others.Any(x => x.Contains("DRAWING LIST"));
                List<string> missingReport = new List<string>();
                if (!drawingList)
                {
                    missingReport.Add("DRAWING LIST");
                }
                if(!mis)
                {
                    missingReport.Add("MIS");
                }
                List<string> duplicate = pdf.GroupBy(x => x).Where(g => g.Count() > 1).Select(x => x.Key).ToList();


                List<string> missingPdf = fileName.Except(pdf).ToList();
                createReport(missingPdf, "PDF_missing");
                createReport(missingReport, "Report_missing");
                createReport(duplicate, "duplicatePdf");
                MessageBox.Show("Completed");
                label4.Visible = false;
            }
            else
            {
                label1.Text = "Select Folder";
                label1.Visible = true;
            }
        }
        public void IFFImport()
        {
            List<string> fileName = new List<string>();
            List<string> dxfFileName = new List<string>();
            List<PartInfo4> partInfos = new List<PartInfo4>();
            List<string> requiredString = new List<string>();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.kssFilePath = openFileDialog1.FileName;
                StreamReader streamReader = new StreamReader(kssFilePath);
                string entireFile = streamReader.ReadToEnd();
                string[] content = entireFile.Split('\r');
                label2.Text = this.kssFilePath;
                label2.Visible = true;
                for (int i = 3; i < content.Length; i++)
                {
                    string s = content[i];
                    bool flag = s.StartsWith("\nD");
                    if (flag)
                    {
                        string[] splitContent = s.Split(',');
                        if (splitContent[4] != "")
                        {
                            requiredString.Add(splitContent[4]);
                            PartInfo4 partInfo = new PartInfo4();
                            partInfo.partMark = splitContent[4];
                            partInfo.Name = splitContent[6];
                            bool checkingPartMark = partInfos.Any(x => x.partMark.Equals(partInfo.partMark));
                            if (!checkingPartMark)
                            {
                                partInfos.Add(partInfo);
                                fileName.Add(partInfo.partMark);
                                if ((partInfo.Name.Equals("FB")) || (partInfo.Name.Equals("PL")) || (partInfo.Name.Equals("SHT")))
                                {
                                    dxfFileName.Add(partInfo.partMark);
                                }

                            }

                        }

                    }
                    else
                    {

                    }
                }
                streamReader.Close();

                this.partInfos = partInfos;
                this.fileName = fileName;
                this.dxfFileName = dxfFileName;
                label4.Text = "Completed";
                label4.Visible = true;
            }
            else
            {
                label2.Text = "Select KSS File";
                label2.Visible = true;
            }

        }
        public void IFAImport()
        {
            List<string> fileName = new List<string>();
            List<string> dxfFileName = new List<string>();
            List<PartInfo4> partInfos = new List<PartInfo4>();
            List<string> requiredString = new List<string>();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.kssFilePath = openFileDialog1.FileName;
                StreamReader streamReader = new StreamReader(kssFilePath);
                string entireFile = streamReader.ReadToEnd();
                string[] content = entireFile.Split('\r');
                label2.Text = this.kssFilePath;
                label2.Visible = true;
                for (int i = 3; i < content.Length; i++)
                {
                    string s = content[i];
                    bool flag = s.StartsWith("\nD");
                    if (flag)
                    {
                        string[] splitContent = s.Split(',');
                        if (splitContent[4] != "")
                        {
                            requiredString.Add(splitContent[4]);
                            PartInfo4 partInfo = new PartInfo4();
                            partInfo.partMark = splitContent[4];
                            partInfo.Name = splitContent[6];
                            bool checkingPartMark = partInfos.Any(x => x.partMark.Equals(partInfo.partMark));
                            if (!checkingPartMark)
                            {
                                partInfos.Add(partInfo);
                                List<bool> listOfresult = new List<bool>();
                                foreach(char c in partInfo.partMark)
                                {
                                    listOfresult.Add(char.IsLower(c));
                                }
                                bool result = listOfresult.Any(x => x.Equals(true));
                                if(!result)
                                {
                                    fileName.Add(partInfo.partMark);
                                }

                            }

                        }

                    }
                    else
                    {

                    }
                }
                streamReader.Close();

                this.partInfos = partInfos;
                this.fileName = fileName;
                label4.Text = "Completed";
                label4.Visible = true;
            }
            else
            {
                label2.Text = "Select KSS File";
                label2.Visible = true;
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /*private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBox1.Checked;
        }*/
    }
}
