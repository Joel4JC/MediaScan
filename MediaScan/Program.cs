/***************************************************************************************************
 * Revision History:
 * 
 * 1.1  First Production Release. - JAC
 * 1.2  Minor Bug Fixes and added ability to Save Default Values for Data Entry Screen. - JAC
 * 1.21 Fixed updating of saved default values. - JAC
 * 1.22 Corrected spelling errors in messages being displayed. - JAC
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MediaScan
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

       /********************************************************************************************
        * Main
        * 
        * The Program/App begins here.
        * 
        * Returns: Nothing
        *
        ********************************************************************************************/
       static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
