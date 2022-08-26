using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace NearFarShearTabFinder
{
    public partial class NearFarShearTabFinder : Form
    {
        public static string skApplicationName = "Shear Tab Near Far Finder";
        public static string skApplicationVersion = "2201.02";

        public NearFarShearTabFinder()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Run", "");
            int ct = 0;

            List<Part> parts = new List<Part>();
            var progress = new Tekla.Structures.Model.Operations.Operation.ProgressBar();
            bool displayResult = progress.Display(100, "ShearTabLocationFinder", "Processing....", "cancel..", " ");
            Model model = new Model();
            Tekla.Structures.Model.UI.ModelObjectSelector modelObjectSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelObjectEnumerator modelObjectEnumerator = modelObjectSelector.GetSelectedObjects();
            int totalCount = modelObjectEnumerator.GetSize();
            int ii = 1;
            int percentValue = 100 * ii / totalCount;
            while (modelObjectEnumerator.MoveNext())
            {
                ct++;
                Part part = modelObjectEnumerator.Current as Part;
                percentValue = 100 * ii / totalCount;
                progress.SetProgress(percentValue.ToString(), 100 * ii / totalCount);
                if ((part != null)&&(part.Name.ToUpper().Contains("BEAM")))
                {
                    string profType = "";
                    part.GetReportProperty("PROFILE_TYPE", ref profType);
                    if(profType == "I")
                    {
                        SecPartInfo secPartInfo = new SecPartInfo(part,model);
                    }
                }
                ii++;
                percentValue = 100 * ii / totalCount;
                progress.SetProgress(percentValue.ToString(), 100 * ii / totalCount);
            }
            progress.Close();
            MessageBox.Show("COMPLETED");
            skWinLib.worklog(skApplicationName, skApplicationVersion, "Run", ";Total:" + ct.ToString());

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

       /* private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBox1.Checked;
        }*/

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
