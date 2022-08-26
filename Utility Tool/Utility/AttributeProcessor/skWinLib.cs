//Omm Muruga 
//Common Function's irrespective of any application, tool, plugin
//Ensure naming of function is meaningfull to understand
//Dont have too many short variables
//If any major changes required in functions please let Vijayakumar know's about this.
//Don't modifiy existing function/class/sub unless you are aware of the consequences
//Use Commentline as much as possible for future reference and other's can understand easily
//Most function's or not tested so ensure its quality at your own risk.
//08Jan22 1502
//23Feb22 1859 SharePoint Two Nuget Package added


using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.VisualBasic;
using Microsoft.SharePoint.Client;

public class skWinLib
{
    //version 22Dec21 1058
    public static string username = System.Environment.UserName;
    public static string systemname = System.Environment.MachineName;
    public static bool WFH = false;


    //registry 14Dec21
    public static string userRoot = "HKEY_CURRENT_USER\\SOFTWARE\\ESSKAY";
    //string logpath = "\\\\192.168.2.3\\Tekla Support\\log"; Path changed by IT and access provided for Trichy and Chennai Centre 21Dec21 1921

    public static string logpath = "\\\\192.168.2.3\\Automation Log;d:\\esskay";
    //public string logpath = "\\\\192.168.2.4\\Automation Log;d:\\esskay;";

    public static string serverpath = "";
    //public class Application
    //{



    public static void accesslog(string skApplicationName, string skApplicationVersion, string task, string skremarks, string modelname = "", string teklaversion = "", string teklaconfiguration = "")
    {

        if ((systemname.ToUpper() != "ALVIN-LAPTOP" && systemname.ToUpper() != "VIJAYAKUMAR-LAP")) //|| (skremarks.ToUpper() == "TESTING")
        {
            skWinLib.Esskay_Tool_Validation(skApplicationName, skApplicationVersion);
            string keyName = userRoot + "\\" + skApplicationName;
            try
            {
                //21121930
                Registry.SetValue(keyName, "Version", skApplicationVersion);
                Registry.SetValue(keyName, "username", username);
                Registry.SetValue(keyName, "systemname", systemname);


                string login = "alvin_413@esskaystructure.onmicrosoft.com";
                string password = "Mechons@1994";
                login = "alvin_413@esskaystructure.onmicrosoft.com";
                password = "Mechons@1994";

                var SecurePassword = new SecureString();

                foreach (char c in password)
                {
                    SecurePassword.AppendChar(c);
                }

                ClientContext clientContext = new ClientContext("https://esskaystructure.sharepoint.com/sites/TeklaTools");
                Web oWebsite = clientContext.Web;
                clientContext.Load(oWebsite, w => w.Title);
                var onlineCredentials = new SharePointOnlineCredentials(login, SecurePassword);
                clientContext.Credentials = onlineCredentials;

                clientContext.ExecuteQuery();
                Console.WriteLine("Link Established for Tool");
                //loo through tools
                clientContext.Load(oWebsite.Lists, lists => lists.Include(list => list.Title, list => list.Id));

                // Execute query.
                clientContext.ExecuteQuery();

                // Enumerate the web.Lists.
                // string appendstr = "";
                foreach (List list in oWebsite.Lists)
                {
                    //appendstr = appendstr + ", " + list.Title;
                    if (list.Title == "Log")
                    {

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = list.AddItem(itemCreateInfo);
                        newItem["Title"] = systemname;
                        newItem["User"] = username;
                        newItem["Application"] = skApplicationName;
                        newItem["VersionName"] = skApplicationVersion;
                        newItem["Task"] = task;
                        newItem["DateTime"] = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                        newItem["Remarks"] = skremarks;
                        newItem["Total"] = "";
                        newItem["Error"] = "";

                        if (modelname != "")
                            newItem["TS_Model"] = modelname;

                        if (teklaversion != "")
                            newItem["TS_Version"] = teklaversion;

                        if (teklaconfiguration != "")
                            newItem["TS_Config"] = teklaconfiguration;

                        newItem.Update();

                        clientContext.ExecuteQuery();
                        Console.WriteLine("My New Item! aDDED");
                        //  Console.ReadLine();

                    }
                    //  Console.WriteLine("List = " + list.Title);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show("Actual Error: " + ex.ToString() + "\nSend this error message to support centre.", "SKSupportAssistance");
            }
        }

        
    }

    public static void Esskay_Tool_Validation(string skApplicationName, string skApplicationVersion)
    {
        //Security Pass1
        //Check whether user is part of esskay domain
        //Security Pass2
        //Check whether esskay dll is accessible in d:\\esskay or c:\esskay
        //Security Pass3
        //Check registry key for esskay
        bool skflag = false;
        string Offical_Version = "";
        //if ((systemname.ToUpper() != "ALVIN-LAPTOP" && systemname.ToUpper() != "VIJAYAKUMAR-LAP"))
        {
            string keyName = userRoot + "\\" + skApplicationName;
            try
            {
                //21121930
                Registry.SetValue(keyName, "Version", skApplicationVersion);
                Registry.SetValue(keyName, "username", username);
                Registry.SetValue(keyName, "systemname", systemname);


                string login = "alvin_413@esskaystructure.onmicrosoft.com";
                string password = "Mechons@1994";
                login = "alvin_413@esskaystructure.onmicrosoft.com";
                password = "Mechons@1994";

                var SecurePassword = new SecureString();

                foreach (char c in password)
                {
                    SecurePassword.AppendChar(c);
                }


                //  ClientContext clientContext = new ClientContext("https://esskaystructure.sharepoint.com/sites/TeklaTools");

                ClientContext clientContext = new ClientContext("https://esskaystructure.sharepoint.com/sites/TeklaTools");
                Web oWebsite = clientContext.Web;
                clientContext.Load(oWebsite, w => w.Title);
                var onlineCredentials = new SharePointOnlineCredentials(login, SecurePassword);
                clientContext.Credentials = onlineCredentials;

                clientContext.ExecuteQuery();
                Console.WriteLine("Link Established for Tool");
                //loo through tools
                clientContext.Load(oWebsite.Lists, lists => lists.Include(list => list.Title, list => list.Id));

                // Execute query.
                clientContext.ExecuteQuery();

                // Enumerate the web.Lists.
                // string appendstr = "";
                foreach (List list in oWebsite.Lists)
                {
                    //appendstr = appendstr + ", " + list.Title;
                    if (list.Title.ToUpper() == "LIST")
                    {

                        ListItemCollection listItemCollection = list.GetItems(CamlQuery.CreateAllItemsQuery());
                        clientContext.Load(listItemCollection,
                                    eachItem => eachItem.Include(
                                                                item => item,
                                                                item => item["Title"],
                                                                item => item["LatestVersion"],
                                                                item => item["IsLowerVersionAllowed"]
                                                                )
                                    );
                        clientContext.ExecuteQuery();
                        //loop through and check the application name
                        foreach (ListItem listItem in listItemCollection)
                        {
                            if ((string)listItem["Title"].ToString().ToUpper() == skApplicationName.ToUpper())
                            {
                                Offical_Version = (string)listItem["LatestVersion"].ToString();
                                double d_Offical_Version = Convert.ToDouble(Offical_Version);
                                bool IsLowerVersionAllowed = (bool)listItem["IsLowerVersionAllowed"];
                                //check version is same or low
                                if (d_Offical_Version == Convert.ToDouble(skApplicationVersion))
                                    skflag = true;
                                else if (d_Offical_Version >= Convert.ToDouble(skApplicationVersion))
                                    skflag = false;

                                if (skflag != true)
                                {
                                    //Check can this application is allowed for lower version in future
                                    if (IsLowerVersionAllowed == false)
                                    {
                                        skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Aborted");
                                        MessageBox.Show(skApplicationName + "\n\n\nCurrent Version is " + skApplicationVersion + ".\nLatest version is " + Offical_Version + ".\n\n\nAborting Now!!!.Contact Support Team.", "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                        skflag = true;
                                        System.Windows.Forms.Application.ExitThread();

                                    }
                                }
                                break;
                            }
                        }
                        break;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Actual Error: " + ex.ToString() + "\nSend this error message to support centre.", "SKSupportAssistance");
                System.Windows.Forms.Application.ExitThread();
            }
            if (skflag == false)
            {
                if (Convert.ToInt32(DateTime.Now.ToString("yyyy")) >= 2023)
                //{                    
                //    if (Offical_Version.Trim().Length == 0)
                //    {
                //        skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Information");
                //        MessageBox.Show(skApplicationName + "\n\n\nCurrent Version " + skApplicationVersion + " is outdated.\n\n\nContact Support Team.", "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //    }
                //    else
                //    {
                //        skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Warning");
                //        MessageBox.Show(skApplicationName + "\n\n\nCurrent Version is " + skApplicationVersion + ".\nLatest version is " + Offical_Version + ".\n\n\nContact Support Team.", "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    }
                //}
                //else
                {
                    if (Offical_Version.Trim().Length == 0)
                    {
                        skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Aborted.");
                        MessageBox.Show(skApplicationName + "\n\n\nCurrent Version " + skApplicationVersion + " is outdated.\n\n\nContact Support Team.", "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        skWinLib.worklog(skApplicationName, skApplicationVersion, "EsskayValidation", "Aborted");
                        MessageBox.Show(skApplicationName + "\n\n\nCurrent Version is " + skApplicationVersion + ".\nLatest version is " + Offical_Version + ".\n\n\nContact Support Team.", "Esskay Support Team", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    System.Windows.Forms.Application.ExitThread();
                }                

            }

        }
    }

    public static void worklog(string skApplicationName, string skApplicationVersion, string task, string skremarks, string modelname = "", string teklaversion = "", string teklaconfiguration = "")
    {
        if ((systemname.ToUpper() != "ALVIN-LAPTOP" && systemname.ToUpper() != "VIJAYAKUMAR-LAP") || (skremarks.ToUpper() == "TESTING"))
        {
            try
            {
                //21121930
                string keyName = userRoot + "\\" + skApplicationName;
                Registry.SetValue(keyName, "Version", skApplicationVersion);
                Registry.SetValue(keyName, "username", username);
                Registry.SetValue(keyName, "systemname", systemname);

                string login = "alvin_413@esskaystructure.onmicrosoft.com";
                string password = "Mechons@1994";
                login = "alvin_413@esskaystructure.onmicrosoft.com";
                password = "Mechons@1994";

                var SecurePassword = new SecureString();

                foreach (char c in password)
                {
                    SecurePassword.AppendChar(c);
                }


                //  ClientContext clientContext = new ClientContext("https://esskaystructure.sharepoint.com/sites/TeklaTools");

                ClientContext clientContext = new ClientContext("https://esskaystructure.sharepoint.com/sites/TeklaTools");
                Web oWebsite = clientContext.Web;
                clientContext.Load(oWebsite, w => w.Title);
                var onlineCredentials = new SharePointOnlineCredentials(login, SecurePassword);
                clientContext.Credentials = onlineCredentials;

                clientContext.ExecuteQuery();
                Console.WriteLine("Link Established for Tool");
                //loo through tools
                clientContext.Load(oWebsite.Lists, lists => lists.Include(list => list.Title, list => list.Id));

                // Execute query.
                clientContext.ExecuteQuery();

                // Enumerate the web.Lists.
                //string appendstr = "";
                foreach (List list in oWebsite.Lists)
                {
                    //appendstr = appendstr + ", " + list.Title;
                    if (list.Title == "Log")
                    {

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = list.AddItem(itemCreateInfo);
                        newItem["Title"] = systemname;
                        newItem["User"] = username;
                        newItem["Application"] = skApplicationName;
                        newItem["VersionName"] = skApplicationVersion;
                        newItem["Task"] = task;
                        newItem["DateTime"] = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                        newItem["Remarks"] = skremarks;
                        newItem["Total"] = "";
                        newItem["Error"] = "";
                        //, string modelname = "", string teklaversion = "", string teklaconfiguration = "")
                        if (modelname != "")
                            newItem["TS_Model"] = modelname;

                        if (teklaversion != "")
                            newItem["TS_Version"] = teklaversion;

                        if (teklaconfiguration != "")
                            newItem["TS_Config"] = teklaconfiguration;

                        newItem.Update();
                        clientContext.ExecuteQuery();
                        Console.WriteLine("My New Item! aDDED");
                        //  Console.ReadLine();


                    }
                    //  Console.WriteLine("List = " + list.Title);
                }


                string[] split = logpath.Split(new Char[] { ';' });

                foreach (string s in split)
                {
                    if (s.Trim() != "")
                    {
                        if (Directory.Exists(s))
                        {
                            //log file based on yearwise
                            string YYYY = DateTime.Today.Year.ToString();
                            using (TextWriter streamWriter = new StreamWriter(s + "\\" + YYYY + "_work.log", true))
                            {
                                streamWriter.WriteLine("System: " + systemname + "; User: " + username + "; Application: " + skApplicationName + "; Version: " + skApplicationVersion + "; Task: " + task + "; DateTime: " + DateTime.Now.ToString() + "; Remarks " + skremarks);
                                streamWriter.Close();
                            }

                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Actual Error: " + ex.ToString() + "\nSend this error message to support centre.", "SKSupportAssistance");
            }
        }

    }

    public static void AutoSizeListView(ListView myListView, string HeaderName, System.Windows.Forms.CheckBox CheckItem)
    {
        if (CheckItem.Checked == true)
        {
            int colct = myListView.Columns.Count;
            string chkHeaderName = HeaderName;
            chkHeaderName = chkHeaderName.Replace(" ", "").ToUpper();
            for (int i = 0; i < colct; i++)
            {
                string colname = myListView.Columns[i].Text;
                colname = colname.Replace(" ", "").ToUpper();
                if (chkHeaderName.Contains("," + colname + ","))
                    myListView.AutoResizeColumn(i, ColumnHeaderAutoResizeStyle.ColumnContent);
            }

        }

    }

    public static void Export_DataGridView(DataGridView mydatagridview, string filename, string myFormat, string myHideColumns)
    {
        int rowct = mydatagridview.Rows.Count;
        int colct = mydatagridview.Columns.Count;
        ArrayList dgv = new ArrayList();
        string fileseperator = ",";

        if (myFormat == "CSV")
            fileseperator = ",";
        else if (myFormat == "PIE")
            fileseperator = "|";
        for (int i = 0; i < rowct; i++)
        {
            DataGridViewRow row = mydatagridview.Rows[i];
            string dgvexportdata = fileseperator;
            for (int j = 0; j < colct; j++)
            {
                DataGridViewCell ox = mydatagridview.Rows[i].Cells[j];
                if (j == 0)
                    dgvexportdata = ox.EditedFormattedValue.ToString();
                else
                    dgvexportdata = dgvexportdata + fileseperator + ox.EditedFormattedValue.ToString();
            }
            if (dgvexportdata.Replace(fileseperator, "").Length >= 1)
                dgv.Add(dgvexportdata);
        }
        //write to file
        using (TextWriter streamWriter = new StreamWriter(filename, false))
        {
            foreach (string linedata in dgv)
                streamWriter.WriteLine(linedata);
            streamWriter.Close();

        }

    }

    public static void Import_DataGridView(DataGridView mydatagridview, string filename, string myFormat, string myHideColumns)
    {
        if (System.IO.File.Exists(filename) == true)
        {
            //read file if exists
            ArrayList dgv = new ArrayList();
            System.IO.StreamReader srdgv = new System.IO.StreamReader(filename, System.Text.Encoding.Default);
            string line = " ";
            while ((line = srdgv.ReadLine()) != null)
            {
                dgv.Add(line);
            }
            srdgv.Close();

            char[] splitchar = new Char[] { ',' };
            if (myFormat == "CSV")
                splitchar = new Char[] { ',' };
            else if (myFormat == "PIE")
                splitchar = new Char[] { '|' };

            //update data grid view
            mydatagridview.Rows.Clear();
            foreach (string linedata in dgv)
            {
                string[] splitrowdata = linedata.ToString().Split(splitchar);
                int colct = splitrowdata.Length - mydatagridview.Columns.Count;
                for (int i = 0; i < colct; i++)
                {
                    mydatagridview.Columns.Add((i + 1).ToString(), (i + 1).ToString());
                    //mydatagridview.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                mydatagridview.Rows.Add(splitrowdata);
                mydatagridview.Rows[mydatagridview.Rows.Count - 1].HeaderCell.Value = (mydatagridview.Rows.Count).ToString();
            }
        }
    }

    public static void updaterowheader(DataGridView myDataGridView)
    {
        //provide row header title
        for (int j = 0; j < myDataGridView.Rows.Count; j++)
        {
            myDataGridView.Rows[j].HeaderCell.Value = (j + 1).ToString();
        }
    }
    public static DateTime setStartup(string skApplicationName, string skapplicationVersion, string Task, string Remark,  string statusbar1,  string statusbar2, Label lblsbar1, Label lblsbar2)
    {
        //12-apr-2022
        DateTime mytime = DateTime.Now;
        //lblsbar1.Visible = true;
        //lblsbar2.Visible = true;
        lblsbar2.Text = statusbar2;
        lblsbar1.Text = statusbar1 + " < " + mytime.ToString("h:mm:ss") + " >...";
        skWinLib.accesslog(skApplicationName, skapplicationVersion, Task, "Start ;" + Remark, skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
        Cursor.Current = Cursors.WaitCursor;
        lblsbar1.Refresh();
        lblsbar2.Refresh();
        return mytime;

    }

    public static void setCompletion(string skApplicationName, string skapplicationVersion, string Task, string Remark, string statusbar1, string statusbar2, Label lblsbar1, Label lblsbar2, DateTime startTime)
    {
        //12-apr-2022
        //lblsbar1.Visible = true;
        //lblsbar2.Visible = false;        
        skWinLib.worklog(skApplicationName, skapplicationVersion, Task, "Complete ;" + Remark, skTSLib.ModelName, skTSLib.Version, skTSLib.Configuration);
        Cursor.Current = Cursors.Default;
        TimeSpan span = DateTime.Now.Subtract(startTime);
        lblsbar1.Text = "Ready. " + statusbar1 + " < " +  span.Minutes.ToString() + ":" + span.Seconds.ToString() + " >";
        lblsbar2.Text = statusbar2;
        lblsbar1.Refresh();
        lblsbar2.Refresh();
    }

    public static System.Boolean IsNumeric(System.Object Expression)
    {
        if (Expression == null || Expression is DateTime)
            return false;

        if (Expression is Int16 || Expression is Int32 || Expression is Int64 || Expression is Decimal || Expression is Single || Expression is Double || Expression is System.Boolean)
            return true;

        try
        {
            if (Expression is string)
                Double.Parse(Expression as string);
            else
                Double.Parse(Expression.ToString());
            return true;
        }
        catch { } // just dismiss errors but return false
        return false;
    }

    public static System.Double GetNumeric(System.Object Expression)
    {
        if (Expression == null || Expression is DateTime)
            return -1;

        if (Expression is Int16 || Expression is Int32 || Expression is Int64 || Expression is Decimal || Expression is Single || Expression is Double || Expression is System.Boolean)
            return Double.Parse(Expression.ToString());


        try
        {
            if (Expression is string)
            {
                string str_double = string.Empty;
                string chkstring = Expression.ToString();
                foreach (char c in chkstring)
                {
                    if ((int)c >= 48 && (int)c <= 57)
                        str_double = str_double + c.ToString();

                }
                if (str_double != string.Empty)
                    return Double.Parse(str_double);
                else
                    return -1;
            }
            return Double.Parse(Expression.ToString());


        }
        catch { } // just dismiss errors but return false


        return -1;
    }



  

}

