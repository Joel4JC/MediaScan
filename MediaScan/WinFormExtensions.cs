using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MediaScan
{
   public static class WinFormExtensions
    {

       /********************************************************************************************
        * AppendLine
        * 
        * Appends the default line terminator to input string variable value, and displays it in the
        * TextBox referenced in the input variable source.
        * 
        * Returns: Nothing
        * 
        ********************************************************************************************/
       public static void AppendLine(this TextBox source, string value)
        {
            if (source.Text.Length == 0)
                source.Text = value;
            else
                source.AppendText("\r\n" + value);
        } // End of public static void AppendLine
    }
}
