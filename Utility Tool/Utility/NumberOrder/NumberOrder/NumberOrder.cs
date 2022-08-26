using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NumberOrder
{
    public partial class NumberOrder : Form
    {
        public static string skApplicationName = "Number Order";
        public static string skApplicationVersion = "2201.02";

        public NumberOrder()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Number Order", "");
            int ct = 0;

            string textbox1 = textBox1.Text.ToString();
            string[] vs1 = textbox1.Split(',');
            
            AlphaNumericSegregate alphaNumericSegregate = new AlphaNumericSegregate(vs1);
            List<AlphaNumericSegregate> alphaNumericSegregates = alphaNumericSegregate.alphaNumericSegregates;
            var obj = from value in alphaNumericSegregate.alphaNumericSegregates
                      group value by value.alpha
                      
                      into newGroup
                      select new
                      {
                          key = newGroup.Key,
                          listOfnumbers = newGroup.OrderBy(x => x.number)

                      };
            string concatenate = "";
            int objCount = 0;
            int completedCount = 0;
            foreach (var value in obj)
            {
                ct++;
                IOrderedEnumerable<AlphaNumericSegregate> ss= value.listOfnumbers;
                int a = 0;
                int countOfSeq = 0;
                int temp = 0;
                foreach (AlphaNumericSegregate alphaNumericSegregate1 in ss)
                {
                    ct++;
                    if(a==0)
                    {
                        if(temp==0)
                        {
                            temp = Convert.ToInt32(alphaNumericSegregate1.number);
                            concatenate = concatenate + alphaNumericSegregate1.value;
                        }
                        else
                        {
                            temp = Convert.ToInt32(alphaNumericSegregate1.number);
                            concatenate = concatenate + "-" + alphaNumericSegregate1.value;
                            countOfSeq++;
                        }
                        
                        
                        a++;
                    }
                    else
                    {
                        int previousTemp = temp;
                        int calcTemp = previousTemp + 1;
                       int currentValue =Convert.ToInt32(alphaNumericSegregate1.number);
                        if(calcTemp == currentValue)
                        {
                            temp = currentValue;
                            if (countOfSeq == 0)
                            {
                                concatenate = concatenate + "-" + alphaNumericSegregate1.value;
                            }
                            else
                            {
                                string[] commaSeparate = concatenate.Split(',');
                                string preprevious = "";
                                if (commaSeparate.Count()>0)
                                {
                                    
                                    preprevious = commaSeparate[commaSeparate.Count()-1].Split('-')[1];
                                }
                                else
                                {
                                    preprevious = concatenate.Split('-')[1];
                                }
                                
                                string previous =alphaNumericSegregate1.value;
                                concatenate = concatenate.Replace(preprevious, previous);
                            }
                            countOfSeq ++;

                        }
                        else
                        {
                            temp = Convert.ToInt32(alphaNumericSegregate1.number);
                            concatenate = concatenate + "," + alphaNumericSegregate1.value;
                            a = 0;
                            countOfSeq = 0;

                        }
                        //else
                        //{
                        //    if(countOfSeq >0)
                        //    {
                        //        concatenate = concatenate + "-" + alphaNumericSegregate1.alpha + previousTemp.ToString();
                        //    }
                        //}
                    }
                }
                completedCount++;
                if (obj.Count()-completedCount>0)
                {
                    concatenate = concatenate + ",";
                }

            }
            label1.Text = concatenate;
            skWinLib.worklog(skApplicationName, skApplicationVersion, "Number Order", ";Total:" + ct.ToString()); 
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

       /* private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBox1.Checked;
        }*/
    }
}
