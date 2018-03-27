using System;
using System.IO;
using System.Drawing;
using System.Drawing.Printing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MediaScan
{
    class PrintFile
    {
        private string strPrintFileName = null;
        private uint uintLineCount = 0;
        private const uint uintLinesPerPage = 72;
        private Font PrintFont;
        private StreamReader StreamToPrint;


        /********************************************************************************************
         * PrintFile
         * 
         * Constructor for class PrintFile 
         * 
         * Parameter In: strFileNameIn - FileName to Print
         *               LineCountIn - Number of Lines In the file.
         * 
         * Returns: Nothing
         * 
         *******************************************************************************************/
        public PrintFile(string strFileNameIn, uint LineCountIn)
        {
            strPrintFileName = strFileNameIn.Trim();
            uintLineCount = LineCountIn;

            Printing();
        }


        /********************************************************************************************
         * GetPageCount
         * 
         * Calculate the number of pages in the report to be printed.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private int GetPageCount()
        {
            // Calculate Page Count Rounded Up
            uint uintAns = uintLineCount / uintLinesPerPage;
            if (uintLineCount % uintLinesPerPage != 0)
                return (int)++uintAns;
            else
                return (int)uintAns;
        } // End GetPageCount


        /********************************************************************************************
         * PrintDoc_PrintPage
         * 
         * PrintPage event of the PrintDocument component, this does the actual printing of the file.
         * Occurs when the output to print for the current page is needed.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs ev)
        {
            float LinesPerPage = 0;
            float yPos = 0;
            int count = 0;
            float LeftMargin = ev.MarginBounds.Left;
            float TopMargin = ev.MarginBounds.Top;
            String Line = null;

            // Calculate the number of lines per page.
            LinesPerPage = ev.MarginBounds.Height / PrintFont.GetHeight(ev.Graphics);

            // Iterate over the file, printing each line.
            while (count < LinesPerPage && ((Line = StreamToPrint.ReadLine()) != null))
            {
                yPos = TopMargin + (count * PrintFont.GetHeight(ev.Graphics));
                ev.Graphics.DrawString(Line, PrintFont, Brushes.Black, LeftMargin, yPos, new StringFormat());
                count++;
            }

            // If more lines exist, print another page.
            if (Line != null)
                ev.HasMorePages = true;
            else
                ev.HasMorePages = false;
        } // End of PrintDoc_PrintPage


        /********************************************************************************************
         * Printing
         * 
         * Setsup the printing parameters, Margins, Page Count, etc. Then it displays a dialog popup
         * box asking if the user is sure they want to print the displayed number of pages.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        public void Printing()
        {
            int intPages = 0;
            string strMessage = null;

            try
            {
                StreamToPrint = new StreamReader(strPrintFileName);
                try
                {
                    PrintFont = new Font("Courier New", 8);
                    PrintDocument PrintDoc = new PrintDocument();
                    PrintDoc.PrintPage += new PrintPageEventHandler(PrintDoc_PrintPage);

                    PrintDoc.DefaultPageSettings.Margins.Left = 50;
                    PrintDoc.DefaultPageSettings.Margins.Right = 50;

                    intPages = GetPageCount();

                    strMessage = string.Format("Are you sure you want to print {0} Pages?", intPages);
                    DialogResult DlgResult = MessageBox.Show(strMessage, "Are You Sure", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (DlgResult == DialogResult.Yes)
                    {
                        PrintDialog SelectPrinter = new PrintDialog();
                        SelectPrinter.Document = PrintDoc;
                        if (SelectPrinter.ShowDialog() == DialogResult.OK)
                        {
                            PrintDoc.Print();
                        }
                    }
                }
                finally
                {
                    StreamToPrint.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("PrintFile-M2: " + ex.Message);
            }
        } // End of Printing
    } // End of class PrintFile
} // End of namespace MediaScan
