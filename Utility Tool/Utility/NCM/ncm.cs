using System.Runtime.InteropServices;
using System.Security;
using System.Globalization;


//using System.Data.SqlClient;
using MySql.Data.MySqlClient;

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
// Tekla Structures namespaces
using Tekla.Structures;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
//using IfcGuid;

using TSG = Tekla.Structures.Geometry3d;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Solid;
using Tekla.Structures.Model.UI;
using TSMUI = Tekla.Structures.Model.UI;
using T3D = Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.Operations;
//using Tekla.Structures.Drawing;
//using TSD = Tekla.Structures.Drawing;


public class skncm
{
    public static void save()
    {
            
    }
    public static bool IsNodeTextAvailable(TreeNode node, string search_node_text)
    {
        //check whether search_node_text exists in the treenode
        for (int i = 0; i < node.Nodes.Count; i++)
        {
            if (node.Nodes[i].Text.ToUpper().Trim() == search_node_text.ToUpper().Trim())
                return true;
 
        }            
        return false;
    }
        
        
}

