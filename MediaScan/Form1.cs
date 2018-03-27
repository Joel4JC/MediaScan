/*
 * MediaScan
 * 
 * Purpose: To scan media and create a listing of all the files on the media, including
 *          files found embedded within .Zip files. A report is generated with file
 *          counts, embedded file counts, hidden file counts, system file counts, as
 *          well as, where the media is from, who brought the media in, who scanned
 *          the media and other information. This information is also stored in a
 *          database for future review and analysis.
 *          
 * Version: 1.09, July 19, 2017
 * 
 * Database: dbMediaScan - Microsoft Access Database backend.
 * 
 * NOTE: The original intented use of this program is for a single user. Thus, the
 * Database is Opened at the start of the program and Closed at the end of the program,
 * instead of opening and closing the database for each transaction.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Ionic.Zip;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace MediaScan
{
    public partial class Form1 : Form
    {
        // Program and Database versioning
        private static string strPath = @"C:\MediaScan\";
        private static string strThisAppName = @"MediaScan";
        private static string strThisAppVersion = @"v1.22";
        private static double dblThisAppVersion = 1.22;
        private static string strDisplayName = string.Empty;
        private static string strRequiredDBName = @"dbMediaScan";
        private static string strRequiredDBVersion = @"v1.1";
        private static double dblRequiredDBVersion = 1.1;

        // Output Files
        private static StreamWriter ReportFileOut = null;
        private static StreamWriter NoPathFileOut = null;
        private static StreamWriter LogFileOut = null;
        private static StreamWriter TempFile = null;

        // Output File Names
        private static string strReportFileName = string.Empty;
        private static string strNoPathFileName = string.Empty;  // Name of temp report file with path suppressed
        private static string strLogFileName = string.Empty;
        private static string strTempFileName = string.Empty;

        // Used for suppressing the display of the full file path 
        // in the full report
        private static string strPathToEliminate = string.Empty; // The path name to suppress
        private static bool bDirectoryOnly = false;              // Indicates the user has selected a directory to search

        // List used to store file names and directories as they are discovered
        // and will be used at the end of the reports.
        private static List<string> strHiddenFiles = new List<string>();
        private static List<string> strSystemDirectories = new List<string>();
        private static List<string> strDirectoriesFound = new List<string>();
        private static List<string> strFilesNotUncompressed = new List<string>();

        // Counters
        private static int intFileCount = 0;
        private static int intSysDirCount = 0;
        private static int intSysFileCount = 0;
        private static int intDirCount = 0;
        private static int intZipFileCount = 0;
        private static int intTotEmbeddedZipFiles = 0;
        private static int intTotFilesInZipFiles = 0;
        private static int intNotUncompressCount = 0;

        private static uint uintReportLineCount = 0;
        private static uint uintSummaryLineCount = 0;
        private static int intLineNumber = 0; // File List Line Number in the report

        // Date and Time the program is executed.
        string strDate = String.Empty;
        string strTime = String.Empty;

        // Database
        private Database dbMediaScan = new MediaScan.Database();
        private static int intRecordID = 0; // Media Unique ID in Database
        private static int intDBErrorCode = 0;



        public Form1()
        {
            InitializeComponent();

           // Display version number in the title bar
            strDisplayName = strThisAppName + " " + strThisAppVersion;
            this.Text = strDisplayName;
        }


        /********************************************************************************************
         * ResetCounters
         * 
         * Resets all counters to zero.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void ResetCounters()
        {
            intFileCount = 0;
            intSysDirCount = 0;
            intSysFileCount = 0;
            intDirCount = 0;
            intZipFileCount = 0;
            intTotEmbeddedZipFiles = 0;
            intTotFilesInZipFiles = 0;
            intNotUncompressCount = 0;
            uintReportLineCount = 0;
            uintSummaryLineCount = 0;
            intLineNumber = 0;
        } // End of ResetCounters


        /********************************************************************************************
         * ResetLists
         * 
         * Clears all the lists used in this program.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void ResetLists()
        {
            strHiddenFiles.Clear();
            strSystemDirectories.Clear();
            strDirectoriesFound.Clear();
        } // End of ResetLists



        /********************************************************************************************
         * OpenLogFile
         * 
         * Opens the LogFile.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void OpenLogFile()
        {
           try
           {
              // Open LogFile File
              strLogFileName = strPath + "MediaScanLog.txt";
              LogFileOut = new StreamWriter(strLogFileName);
              LogFileOut.WriteLine(strDate + " @ " + strTime);
           }
           catch
           {
              MessageBox.Show("Form1-M0: Error Opening the LogFile", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
              Application.Exit();
           }
        } // End of OpenLogFile


        /********************************************************************************************
         * CloseLogFile
         * 
         * Closes the LogFile if it is open
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void CloseLogFile()
        {
           try
           {
              if (LogFileOut != null)
                 if (LogFileOut.BaseStream != null)
                    LogFileOut.Close();
           }
           catch (IOException)
           {
              MessageBox.Show("Form1-M1: Error Closing the LogFile On Exit", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
        } // End of Close LogFile


        /********************************************************************************************
         * CloseAndDelete
         * 
         * Closes all the files and Deletes the Report Files and Temp Files.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void CloseAndDelete()
        {
            try
            {
                //Close Report and Log Files
               if (ReportFileOut != null)
                  if (ReportFileOut.BaseStream != null)
                     ReportFileOut.Close();
                // Delete ONLY the Report Files
                if (File.Exists(strReportFileName))
                    File.Delete(strReportFileName);

                // Close and delete temp files
                DeleteTempFiles();
            }
            catch (IOException)
            {
                MessageBox.Show("Form1-M2: Error Closing/Deleting Files On Exit", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogFileOut.WriteLine("Form1-M2: Error Closing/Deleting Files On Exit");
            }
        } // End of CloseAndDelete


        /********************************************************************************************
         * ResetFileNames
         * 
         * Reset all file name strings to empty strings.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void ResetFileNames()
        {
            strReportFileName = string.Empty;
            strNoPathFileName = string.Empty;
            strLogFileName = string.Empty;
            strTempFileName = string.Empty;
        } // End of ResetFileNames


        /********************************************************************************************
         * BlankLine
         * 
         * Creates lank string of length n.
         * 
         * Inputs: n - The size of the blank line to create.
         * 
         * Returns: A string of blanks/spaces of length n
         * 
         ********************************************************************************************/
        private string BlankLine(int n)
        {
            string s = "";
            for (int i = 0; i < n; i++)
            {
                s += " ";
            }
            return (s);
        } // End of private string BlankLine(int n)


       /********************************************************************************************
        * NoPathFileOutWrite
        * 
        * Write the same output as the full report, except without the full file paths
        * 
        * Inputs: strPathIn - Current Path to examine to see if it needs to be modified.
        *         strIndent - Indent string of a specific size for the current file to be printed.
        * 
        * Returns: Nothing
        * 
        ********************************************************************************************/
       private void NoPathFileOutWrite(string strPathIn, string strIndent)
        {
           string strPathRemoved = string.Empty;

           if (strPathToEliminate == strPathIn)
           {
              NoPathFileOut.WriteLine(" ");
           }
           else
           {
              int index = strPathIn.IndexOf(strPathToEliminate);
              strPathRemoved = (index < 0) ? strPathIn : strPathIn.Remove(index, strPathToEliminate.Length);
              if (String.IsNullOrEmpty(strPathRemoved))
                 strPathRemoved = " ";
              if (String.IsNullOrEmpty(strIndent))
                 NoPathFileOut.WriteLine(strPathRemoved);
              else
                 NoPathFileOut.WriteLine(strIndent + strPathRemoved);
           }

        } // End of private void NoPathFileOutWrite


       /********************************************************************************************
        * GetEmbeddedZipMemoryStream
        * 
        * Unzips an embedded zip file to memory, eliminates disk IO, in order to get the zip file's 
        * directory content.
        * 
        * Inputs: zip - zip file to be extracted to memory.
        * 
        * Returns: a pointer to the extract zip file in memory.
        * 
        ********************************************************************************************/
       private MemoryStream GetEmbeddedZipMemoryStream(ZipEntry zip)
        {
            MemoryStream zipMs = new MemoryStream();
            zip.Extract(zipMs);
            zipMs.Seek(0, SeekOrigin.Begin);
            return zipMs;
        } // End of private MemoryStream GetEmbeddedZipMemoryStream(...)


        /********************************************************************************************
         * UnZipToMemory
         * 
         * Extracts the content of a Zip file to memory.
         * 
         * Inputs: ZipFileName - 
         * 
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void UnZipToMemory(string ZipFileName)
        {
            int IndentSize = 0;

            ReportFileOut.WriteLine(ZipFileName);
            NoPathFileOutWrite(ZipFileName, "");
            txtMediaContent.AppendLine(ZipFileName);
            AddFileToDB(ZipFileName, IndentSize);

            try
            {
                using (var zip = ZipFile.Read(ZipFileName))
                {
                    IndentSize += 4;
                    foreach (ZipEntry entry in zip)
                    {
                        GetEmbeddedZip(entry, IndentSize);
                    }
                    IndentSize -= 4;
                }
            }
            catch
            {
                MessageBox.Show("Form1-M3: Exception reading the zip file.");
                LogFileOut.WriteLine("Form1-M3: Exception reading the zip file.");
            }
        } // End of UnZipToMemory


        /********************************************************************************************
         * GetEmbeddedZip
         * 
         * Gets Zip files which are embedded in other Zip files.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/ 
        private void GetEmbeddedZip(ZipEntry EmbeddedFile, int IndentSize)
        {
            string sIndent = " ";

            if (EmbeddedFile.FileName.EndsWith(".zip"))
            {
                intTotEmbeddedZipFiles++;
                intTotFilesInZipFiles++;
                sIndent = BlankLine(IndentSize);
                txtMediaContent.AppendLine(sIndent + EmbeddedFile.FileName);
                AddFileToDB(EmbeddedFile.FileName, IndentSize);
                ReportFileOut.WriteLine(sIndent + EmbeddedFile.FileName);
                NoPathFileOutWrite(EmbeddedFile.FileName, sIndent);
                uintReportLineCount++;

                MemoryStream ZipMS = GetEmbeddedZipMemoryStream(EmbeddedFile);
                IndentSize += 4;
                using (ZipFile Ezip = ZipFile.Read(ZipMS))
                {
                    foreach (ZipEntry EmbeddedEntry in Ezip)
                    {
                        if (EmbeddedEntry.FileName.EndsWith(".zip"))
                        {
                            GetEmbeddedZip(EmbeddedEntry, IndentSize);
                        }
                        else
                        {
                            if (EmbeddedEntry.FileName.EndsWith(".rar"))
                            {
                                txtMediaContent.AppendLine("Uncompression of .rar file is not supported in this version of MediaScan.");
                                strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                            }
                            else if (EmbeddedEntry.FileName.EndsWith(".cab"))
                            {
                                txtMediaContent.AppendLine("Uncompression of .cab file is not supported in this version of MediaScan.");
                                strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                            }
                            else if (EmbeddedEntry.FileName.EndsWith(".gz"))
                            {
                                txtMediaContent.AppendLine("Uncompression of .gz file is not supported in this version of MediaScan");
                                strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                            }

                            intTotFilesInZipFiles++;
                            sIndent = BlankLine(IndentSize);
                            txtMediaContent.AppendLine(sIndent + EmbeddedEntry.FileName);
                            AddFileToDB(EmbeddedEntry.FileName, IndentSize);
                            ReportFileOut.WriteLine(sIndent + EmbeddedEntry.FileName);
                            NoPathFileOutWrite(EmbeddedEntry.FileName, sIndent);
                            uintReportLineCount++;
                        }
                    }
                    IndentSize -= 4;
                }
            }
            else
            {
                if (EmbeddedFile.FileName.EndsWith(".rar"))
                {
                    txtMediaContent.AppendLine("Uncompression of .rar file is not supported in this version of MediaScan.");
                    strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                }
                else if (EmbeddedFile.FileName.EndsWith(".cab"))
                {
                    txtMediaContent.AppendLine("Uncompression of .cab file is not supported in this version of MediaScan.");
                    strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                }
                else if (EmbeddedFile.FileName.EndsWith(".gz"))
                {
                    txtMediaContent.AppendLine("Uncompression of .gz file is not supported in this version of MediaScan");
                    strFilesNotUncompressed.Add(EmbeddedFile.FileName);  // Save file names for later - summary
                }

                intTotFilesInZipFiles++;
                sIndent = BlankLine(IndentSize);
                txtMediaContent.AppendLine(sIndent + EmbeddedFile.FileName);
                AddFileToDB(EmbeddedFile.FileName, IndentSize);
                ReportFileOut.WriteLine(sIndent + EmbeddedFile.FileName);
                NoPathFileOutWrite(EmbeddedFile.FileName, sIndent);
                uintReportLineCount++;
            }
        } // End of private void GetEmbeddedZip(ZipEntry EmbeddedFile)


        /********************************************************************************************
         * isAlphaNumeric
         * 
         * Chesk if the strToCheck is contains only alphanumeric characters. In some cases strToCheck
         * is used as part of a filename.
         * 
         * Returns: True if the string contains only alphanumeric characters, otherwise false.
         * 
         ********************************************************************************************/
        public static Boolean isAlphaNumeric(string strToCheck)
        {
            Regex rg = new Regex(@"^[a-zA-Z0-9\-]*$");
            return rg.IsMatch(strToCheck);
        } // End of static Boolean isAlphaNumeric


        /********************************************************************************************
         * ValidateInputFields
         * 
         * Validates that all of the input fields contains data and that it is valid.
         * 
         * Returns: True if ALL of the fields have valid input data, otherwise false.
         * 
         ********************************************************************************************/
        private bool ValidateInputFields()
        {
            bool bValid = true;

            // Validate all input fields have a value
            if (txtYourName.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter Your Name!");
                bValid = false;
            }

            if (txtUsername.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Can't Find Your Username, Contact Your System Administrator!");
                bValid = false;
            }

            if (txtMediaBroughtInBy.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter Media Brought In By!");
                bValid = false;
            }

            if (txtExistingMediaNum.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter An Existing Media Number!");
                bValid = false;
            }

            if (txtMediaTitle.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter Media Title!");
                bValid = false;
            }

            if (txtFromSystem.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter From System!");
                bValid = false;
            }

            if ( (txtPA03TempNum.Text == string.Empty) || !isAlphaNumeric(txtPA03TempNum.Text) )
            {
                txtMediaContent.AppendLine("Invalid Text For PA03 Temp Number!");
                txtMediaContent.AppendLine("PA03 Temp Number HAS TO BE AlphaNumeric Characters ONLY!");
                bValid = false;
            }

            if (txtToSystem.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter To System!");
                bValid = false;
            }

            if (txtCheckedBy.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter Checked By!");
                bValid = false;
            }

            if (txtDTANum.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter A DTA Number!");
                bValid = false;
            }

            if (txtDriveToScan.Text == string.Empty)
            {
                txtMediaContent.AppendLine("Must Enter A Drive To Scan!");
                bValid = false;
            }

            return bValid;
        } // End of ValidateInputFields


        /********************************************************************************************
         * PrintHeader
         * 
         * Displays the Header information in the App Window, as well as, prints the Header
         * Information on the reports.
         * 
         * Returns: Nothing
         *
         ********************************************************************************************/
        private void PrintHeader()
        {
            LogFileOut.WriteLine("Media Scanned on " + strDate + " at " + strTime);
           
            // Console Display
            txtMediaContent.Clear();
            txtMediaContent.AppendLine(strDisplayName);
            txtMediaContent.AppendLine("Media Scanned on " + strDate + " at " + strTime);
            txtMediaContent.AppendLine(string.Format("{0, -25} {1, -23} {2, -23} {3, -20}", lblYourName.Text, txtYourName.Text, lblUserName.Text, txtUsername.Text));
            txtMediaContent.AppendLine(string.Format("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaBroughtInBy.Text, txtMediaBroughtInBy.Text, lblExistingMediaNum.Text, txtExistingMediaNum.Text));
            txtMediaContent.AppendLine(string.Format("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaTitle.Text, txtMediaTitle.Text, lblFromSystem.Text, txtFromSystem.Text));
            txtMediaContent.AppendLine(string.Format("{0, -25} {1, -23} {2, -23} {3, -20}", lblPA03TempNum.Text, txtPA03TempNum.Text, lblToSystem.Text, txtToSystem.Text));
            txtMediaContent.AppendLine(string.Format("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaCheckedBy.Text, txtCheckedBy.Text, lblDTANum.Text, txtDTANum.Text));
            txtMediaContent.AppendLine(" ");
            txtMediaContent.AppendLine(string.Format("Drive Being Scanned: {0}", txtDriveToScan.Text));

            // Full Report Output
            ReportFileOut.WriteLine(strDisplayName);
            ReportFileOut.WriteLine("Media Scanned on " + strDate + " at " + strTime);
            ReportFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblYourName.Text, txtYourName.Text, lblUserName.Text, txtUsername.Text);
            ReportFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaBroughtInBy.Text, txtMediaBroughtInBy.Text, lblExistingMediaNum.Text, txtExistingMediaNum.Text);
            ReportFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaTitle.Text, txtMediaTitle.Text, lblFromSystem.Text, txtFromSystem.Text);
            ReportFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblPA03TempNum.Text, txtPA03TempNum.Text, lblToSystem.Text, txtToSystem.Text);
            ReportFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaCheckedBy.Text, txtCheckedBy.Text, lblDTANum.Text, txtDTANum.Text);
            ReportFileOut.WriteLine(" ");
            ReportFileOut.WriteLine("Drive Being Scanned: {0}", txtDriveToScan.Text);
            uintReportLineCount += 8;

           // Suppressed Path Report Output
            NoPathFileOut.WriteLine(strDisplayName);
            NoPathFileOut.WriteLine("Media Scanned on " + strDate + " at " + strTime);
            NoPathFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblYourName.Text, txtYourName.Text, lblUserName.Text, txtUsername.Text);
            NoPathFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaBroughtInBy.Text, txtMediaBroughtInBy.Text, lblExistingMediaNum.Text, txtExistingMediaNum.Text);
            NoPathFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaTitle.Text, txtMediaTitle.Text, lblFromSystem.Text, txtFromSystem.Text);
            NoPathFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblPA03TempNum.Text, txtPA03TempNum.Text, lblToSystem.Text, txtToSystem.Text);
            NoPathFileOut.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaCheckedBy.Text, txtCheckedBy.Text, lblDTANum.Text, txtDTANum.Text);
            NoPathFileOut.WriteLine(" ");
            NoPathFileOut.WriteLine("Drive Being Scanned: {0}", txtDriveToScan.Text);

            // Summary Report Output
            TempFile.WriteLine(strDisplayName);
            TempFile.WriteLine("Media Scanned on " + strDate + " at " + strTime);
            TempFile.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblYourName.Text, txtYourName.Text, lblUserName.Text, txtUsername.Text);
            TempFile.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaBroughtInBy.Text, txtMediaBroughtInBy.Text, lblExistingMediaNum.Text, txtExistingMediaNum.Text);
            TempFile.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaTitle.Text, txtMediaTitle.Text, lblFromSystem.Text, txtFromSystem.Text);
            TempFile.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblPA03TempNum.Text, txtPA03TempNum.Text, lblToSystem.Text, txtToSystem.Text);
            TempFile.WriteLine("{0, -25} {1, -23} {2, -23} {3, -20}", lblMediaCheckedBy.Text, txtCheckedBy.Text, lblDTANum.Text, txtDTANum.Text);
            TempFile.WriteLine(" ");
            TempFile.WriteLine("Drive Being Scanned: {0}", txtDriveToScan.Text);
            uintSummaryLineCount += 8;
        } // End of PrintHeader


       /***********************************************************************************************************************
        * 
        * AddFileToDB
        * 
        * Add each file on the media to the database.
        * 
        * Returns: Nothing
        * 
        ***********************************************************************************************************************/
       private void AddFileToDB(string strFullPath, int intIndentLength)
        {
           string strPath = "";
           string strFile = "";

          if (strFullPath != null)
          {
             strPath = Path.GetDirectoryName(strFullPath);
             strFile = Path.GetFileName(strFullPath);

             intLineNumber++;
             intDBErrorCode = dbMediaScan.qryAddFile(intRecordID, intLineNumber, strPath, strFile, intIndentLength);
             if (intDBErrorCode == dbMediaScan.FatalError)
                DBFatalErrorExit();
          }

        } // End of AddFileToDB(string strFullPath, int intIndentLength)


        /**************************************************************************************************************************
         * 
         * DirSearch
         * 
         * Recursively scan all directories found at the specified path. Display each file found that is Not in a System Directory.
         * 
         * Returns: Nothing
         * 
         **************************************************************************************************************************/
        private void DirSearch(string strSourcePath)
        {
            try
            {
                foreach (string strFile in Directory.GetFiles(strSourcePath, "*.*"))
                {
                    FileAttributes ItsAttributes = File.GetAttributes(strFile);
                    if ((ItsAttributes & FileAttributes.System) == FileAttributes.System)
                    {
                        intSysFileCount++;  // System File Count
                    }
                    else
                    {
                        if ((ItsAttributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                            strHiddenFiles.Add(strFile);  // Save Hidden file names for later - summary
                        else
                        {
                            string strExt = Path.GetExtension(strFile);
                            if (strExt == ".zip")
                            {
                                UnZipToMemory(strFile);
                                intZipFileCount++;
                            }
                            else if (strExt == ".rar")
                            {
                                txtMediaContent.AppendLine("Uncompression of .rar file is not supported in this version of MediaScan.");
                                strFilesNotUncompressed.Add(strFile);  // Save file names for later - summary
                            }
                            else if (strExt == ".cab")
                            {
                                txtMediaContent.AppendLine("Uncompression of .cab file is not supported in this version of MediaScan.");
                                strFilesNotUncompressed.Add(strFile);  // Save file names for later - summary
                            }
                            else if (strExt == ".gz")
                            {
                                txtMediaContent.AppendLine("Uncompression of .gz file is not supported in this version of MediaScan");
                                strFilesNotUncompressed.Add(strFile);  // Save file names for later - summary
                            }
                            txtMediaContent.AppendLine(strFile);
                            AddFileToDB(strFile, 0);
                            ReportFileOut.WriteLine(strFile);
                            NoPathFileOutWrite(strFile, "");
                            uintReportLineCount++;
                            intFileCount++;   // Regular file count

                        } // else
                    } // else
                } // foreach

                foreach (string strDirectory in Directory.GetDirectories(strSourcePath))
                {
                    FileAttributes ItsAttributes = File.GetAttributes(strDirectory);
                    if ((ItsAttributes & FileAttributes.System) == FileAttributes.System)
                    {
                        strSystemDirectories.Add(strDirectory); // Save System Directory names for later - summary. Do Not Search!
                    }
                    else
                    {
                        strDirectoriesFound.Add(strDirectory); // Save the Directory name
                        DirSearch(strDirectory); // Recursive Call
                    }
                }

            }
            catch (Exception ex)
            {
                txtMediaContent.AppendLine(ex.Message);
            }
        } // end of DirSearch


        /********************************************************************************************
         * GetSummary
         * 
         * Displays Summary information in the App Window, as well as, prints it on the report.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetSummary(string strSourcePath)
        {
            // Summary Output to Display, Full Report and Temp Summary Report
            txtMediaContent.AppendLine(" ");
            txtMediaContent.AppendLine("------ SUMMARY ------\n");
            ReportFileOut.WriteLine("\n------ SUMMARY ------\n");
            NoPathFileOut.WriteLine("\n------ SUMMARY ------\n");
            TempFile.WriteLine("\n------ SUMMARY ------\n");
            uintReportLineCount++;
            uintSummaryLineCount++;

            // Display all Hidden Files found
            if (strHiddenFiles.Count != 0)
            {
                ReportFileOut.WriteLine("-----------------------------");
                ReportFileOut.WriteLine(" ");
                ReportFileOut.WriteLine("The following Hidden Files were found on the media:");
                ReportFileOut.WriteLine(" ");

                NoPathFileOut.WriteLine("-----------------------------");
                NoPathFileOut.WriteLine(" ");
                NoPathFileOut.WriteLine("The following Hidden Files were found on the media:");
                NoPathFileOut.WriteLine(" ");

                uintReportLineCount += 4;
                foreach (string strFile in strHiddenFiles)
                {
                    ReportFileOut.WriteLine(strFile);
                    NoPathFileOut.WriteLine(strFile);
                    uintReportLineCount++;
                    intFileCount++;
                }
                ReportFileOut.WriteLine(" ");
                NoPathFileOut.WriteLine(" ");
                uintReportLineCount++;
            }

            // Display all System Directories found
            if (strSystemDirectories.Count != 0)
            {
               //Output to full report file
               ReportFileOut.WriteLine("-----------------------------");
               ReportFileOut.WriteLine(" ");
               ReportFileOut.WriteLine("The following System Directories were found on the media and NOT Searched:");
               ReportFileOut.WriteLine(" ");

               NoPathFileOut.WriteLine("-----------------------------");
               NoPathFileOut.WriteLine(" ");
               NoPathFileOut.WriteLine("The following System Directories were found on the media and NOT Searched:");
               NoPathFileOut.WriteLine(" ");

               uintReportLineCount += 4;

               //Output to display
               txtMediaContent.AppendLine("-----------------------------");
               txtMediaContent.AppendLine(" ");
               txtMediaContent.AppendLine("The following System Directories were found on the media and NOT Searched:");
               txtMediaContent.AppendLine(" ");

               //Output to temp summary report file
               TempFile.WriteLine("-----------------------------");
               TempFile.WriteLine(" ");
               TempFile.WriteLine("The following System Directories were found on the media and NOT Searched:");
               TempFile.WriteLine(" ");
               uintSummaryLineCount += 4;

               foreach (string strDir in strSystemDirectories)
               {
                  ReportFileOut.WriteLine(strDir);
                  NoPathFileOutWrite(strDir, "");
                  txtMediaContent.AppendLine(strDir);
                  TempFile.WriteLine(strDir);
                  uintReportLineCount++;
                  uintSummaryLineCount++;
                  intSysDirCount++;
               }

               //Output to full report file
               ReportFileOut.WriteLine(" ");
               ReportFileOut.WriteLine("*** " + intSysDirCount.ToString() + " DIRECTORIES WERE NOT SEARCHED ****");
               ReportFileOut.WriteLine(" ");

               NoPathFileOut.WriteLine(" ");
               NoPathFileOut.WriteLine("*** " + intSysDirCount.ToString() + " DIRECTORIES WERE NOT SEARCHED ****");
               NoPathFileOut.WriteLine(" ");

               uintReportLineCount += 3;

               //Output to display
               txtMediaContent.AppendLine(" ");
               txtMediaContent.AppendLine("*** " + intSysDirCount.ToString() + " DIRECTORIES WERE NOT SEARCHED ****");
               txtMediaContent.AppendLine(" ");

               //Output to temp summary report file
               TempFile.WriteLine(" ");
               TempFile.WriteLine("*** " + intSysDirCount.ToString() + " DIRECTORIES WERE NOT SEARCHED ****");
               TempFile.WriteLine(" ");
               uintSummaryLineCount += 3;

            }

            // Display all Uncompressed Archive files
            if (strFilesNotUncompressed.Count != 0)
            {
               //Output to full report file
               ReportFileOut.WriteLine("-----------------------------");
               ReportFileOut.WriteLine(" ");
               ReportFileOut.WriteLine("The following Archive Files were NOT Uncompressed:");
               ReportFileOut.WriteLine(" ");

               NoPathFileOut.WriteLine("-----------------------------");
               NoPathFileOut.WriteLine(" ");
               NoPathFileOut.WriteLine("The following Archive Files were NOT Uncompressed:");
               NoPathFileOut.WriteLine(" ");

               uintReportLineCount += 4;

               //Output to display
               txtMediaContent.AppendLine("-----------------------------");
               txtMediaContent.AppendLine(" ");
               txtMediaContent.AppendLine("The following Archive Files were NOT Uncompressed:");
               txtMediaContent.AppendLine(" ");

               //Output to temp summary report file
               TempFile.WriteLine("-----------------------------");
               TempFile.WriteLine(" ");
               TempFile.WriteLine("The following Archive Files were NOT Uncompressed:");
               TempFile.WriteLine(" ");
               uintSummaryLineCount += 4;

               foreach (string strNotUncompressed in strFilesNotUncompressed)
               {
                  ReportFileOut.WriteLine(strNotUncompressed);
                  NoPathFileOutWrite(strNotUncompressed, "");
                  txtMediaContent.AppendLine(strNotUncompressed);
                  TempFile.WriteLine(strNotUncompressed);
                  uintReportLineCount++;
                  uintSummaryLineCount++;
                  intNotUncompressCount++;
               }

               //Output to full report file
               ReportFileOut.WriteLine(" ");
               ReportFileOut.WriteLine("*** " + intNotUncompressCount.ToString() + " ARCHIVE FILES NOT UNCOMPRESSED ****");
               ReportFileOut.WriteLine(" ");

               NoPathFileOut.WriteLine(" ");
               NoPathFileOut.WriteLine("*** " + intNotUncompressCount.ToString() + " ARCHIVE FILES NOT UNCOMPRESSED ****");
               NoPathFileOut.WriteLine(" ");

               uintReportLineCount += 3;

               //Output to display
               txtMediaContent.AppendLine(" ");
               txtMediaContent.AppendLine("*** " + intNotUncompressCount.ToString() + " ARCHIVE FILES NOT UNCOMPRESSED ****");
               txtMediaContent.AppendLine(" ");

               //Output to temp summary report file
               TempFile.WriteLine(" ");
               TempFile.WriteLine("*** " + intNotUncompressCount.ToString() + " ARCHIVE FILES NOT UNCOMPRESSED ****");
               TempFile.WriteLine(" ");
               uintSummaryLineCount += 3;

            }

            // Display all Directories searched
            if (strDirectoriesFound.Count != 0)
            {
                foreach (string strDir in strDirectoriesFound)
                {
                    intDirCount++;
                }
            }
            //Output to full report file
            ReportFileOut.WriteLine(@"Total number of Files on '" + strSourcePath + "' are " + intFileCount.ToString());
            ReportFileOut.WriteLine(@"Total number of Directories Searched on '" + strSourcePath + "' are " + intDirCount.ToString());
            ReportFileOut.WriteLine(@"Total number of ZIP Files Searched on '" + strSourcePath + "' are " + intZipFileCount.ToString());
            ReportFileOut.WriteLine(@"In the {0} ZIP Files there are {1} Embedded Zip Files and {2} files in all zip files", intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles);
            ReportFileOut.WriteLine(@"There were a total of {0} Compressed Files that were Not Uncompressed", intNotUncompressCount);

            NoPathFileOut.WriteLine(@"Total number of Files on '" + strSourcePath + "' are " + intFileCount.ToString());
            NoPathFileOut.WriteLine(@"Total number of Directories Searched on '" + strSourcePath + "' are " + intDirCount.ToString());
            NoPathFileOut.WriteLine(@"Total number of ZIP Files Searched on '" + strSourcePath + "' are " + intZipFileCount.ToString());
            NoPathFileOut.WriteLine(@"In the {0} ZIP Files there are {1} Embedded Zip Files and {2} files in all zip files", intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles);
            NoPathFileOut.WriteLine(@"There were a total of {0} Compressed Files that were Not Uncompressed", intNotUncompressCount);

            uintReportLineCount += 5;

            //Output to display
            txtMediaContent.AppendLine(@"Total number of Files on '" + strSourcePath + "' are " + intFileCount.ToString());
            txtMediaContent.AppendLine(@"Total number of Directories Searched on '" + strSourcePath + "' are " + intDirCount.ToString());
            txtMediaContent.AppendLine(@"Total number of ZIP Files Searched on '" + strSourcePath + "' are " + intZipFileCount.ToString());
            txtMediaContent.AppendLine(String.Format(@"In the {0} ZIP Files there are {1} Embedded Zip Files and {2} files in all zip files", intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles));
            txtMediaContent.AppendLine(String.Format(@"There were a total of {0} Compressed Files that were Not Uncompressed", intNotUncompressCount));

            //Output to temp summary report file
            TempFile.WriteLine(@"Total number of Files on '" + strSourcePath + "' are " + intFileCount.ToString());
            TempFile.WriteLine(@"Total number of Directories Searched on '" + strSourcePath + "' are " + intDirCount.ToString());
            TempFile.WriteLine(@"Total number of ZIP Files Searched on '" + strSourcePath + "' are " + intZipFileCount.ToString());
            TempFile.WriteLine(@"In the {0} ZIP Files there are {1} Embedded Zip Files and {2} files in all zip files", intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles);
            TempFile.WriteLine(@"There were a total of {0} Compressed Files that were Not Uncompressed", intNotUncompressCount);
            uintSummaryLineCount += 5;

            if (intSysFileCount != 0)
            {
               ReportFileOut.WriteLine("There were " + intSysFileCount.ToString() + " System files found on the media");
               NoPathFileOut.WriteLine("There were " + intSysFileCount.ToString() + " System files found on the media");
               txtMediaContent.AppendLine("There were " + intSysFileCount.ToString() + " System files found on the media");
               TempFile.WriteLine("There were " + intSysFileCount.ToString() + " System files found on the media");
               uintReportLineCount++;
               uintSummaryLineCount++;
            }
        } // End of GetSummary


       /***********************************************************************************************
        * DeleteTempFiles
        * 
        * Closes and delete temp files only
        * 
        * Returns: Nothing
        * 
        ***********************************************************************************************/
       private static void DeleteTempFiles()
        {
           if (NoPathFileOut != null)
              if (NoPathFileOut.BaseStream != null)
                 NoPathFileOut.Close();
           if (File.Exists(strNoPathFileName))
              File.Delete(strNoPathFileName);
           if (TempFile != null)
              if (TempFile.BaseStream != null)
                 TempFile.Close();
           if (File.Exists(strTempFileName))
              File.Delete(strTempFileName);
        } // End of DeleteTempFiles


        /********************************************************************************************
         * CloseFiles
         * 
         * Closes all the files and the database.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void CloseFiles()
        {
            try
            {
               // Close the database
               
               dbMediaScan.closeDB();

               // Close report and log files.
               if (ReportFileOut != null)
                  if (ReportFileOut.BaseStream != null)
                     ReportFileOut.Close();
               if (NoPathFileOut != null)
                  if (NoPathFileOut.BaseStream != null)
                     NoPathFileOut.Close();
                if (TempFile != null)
                    if (TempFile.BaseStream != null)
                        TempFile.Close();
            }
            catch (IOException)
            {
                MessageBox.Show("Form1-M4: Error Closing/Deleting Files On Exit", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogFileOut.WriteLine("Form1-M4: Error Closing/Deleting Files On Exit");
            }
        } // End of CloseFiles


        /********************************************************************************************
         * DBFatalErrorExit
         * 
         * A call to a database query returned a fatal error code. Notify the user, close all files, 
         * delete all report and temporary files and then exit the program.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void DBFatalErrorExit()
        {
           MessageBox.Show("Database Fatal Error\nApplication Will Terminate", "Severe Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
           LogFileOut.WriteLine("Database Fatal Error\nApplication Will Terminate");
           CloseAndDelete();

           // Close this file last so we can record the above error if it occurs
           CloseLogFile();

           Application.Exit();

        } // End of DBFatalErrorExit


        /********************************************************************************************
         * Form1_Load
         * 
         * Load the App's main Window.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void Form1_Load(object sender, EventArgs e)
        {
            // Get local machine information
            MachineInfo MI = new MachineInfo();
            txtUsername.Text = MI.LoggedInUser;
            txtUsername.Enabled = false;
            txtUsername.TabStop = false;

           int intRows = 0;
           int intDBInfoID = 0;
           string strDBVersion = String.Empty;
           string strDBName = String.Empty;
           double dblDBVersion = 0.0;
           string strAppName = String.Empty;
           string strAppVersion = String.Empty;
           double dblAppVersion = 0.0;
           bool DBOpenedHere = false;
           bool InvalidAppOrDB = false;

           // Get current date and time
           strDate = DateTime.Now.ToString("yyyyMMdd");
           strTime = DateTime.Now.ToString("HHmmss");

           // Open LogFile File
           OpenLogFile();
           LogFileOut.WriteLine(strDisplayName + " Started On " + strDate + " At " + strTime);

           if (dbMediaScan.DBConnState == ConnectionState.Closed)
           {
              dbMediaScan.openDB();
              DBOpenedHere = true;
           }

           txtMediaContent.Clear();
           txtMediaContent.AppendLine(strThisAppName + " " + strThisAppVersion);
           txtMediaContent.AppendLine("Validating Database and Application");
           LogFileOut.WriteLine("Validating Database and Application");
           intRows = dbMediaScan.qryFindDBInfo(out intDBInfoID, out strDBVersion, out strDBName, out strAppName, out strAppVersion, out dblAppVersion,
                     out dblDBVersion);

           if (intRows == 1)
           {
              if (dblDBVersion != dblRequiredDBVersion)
              {
                 MessageBox.Show("Form1-M5: Invalid Database Version", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 LogFileOut.WriteLine("Form1-M5: Invalid Database Version - " + strRequiredDBVersion + " is Required!");
                 InvalidAppOrDB = true;
              }

              if (strDBName != strRequiredDBName)
              { 
                 MessageBox.Show("Form1-M6: Invalid Database Name", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 LogFileOut.WriteLine("Form1-M36: Invalid Database Name");
                 InvalidAppOrDB = true;
              }
                    
              if (strThisAppName != strAppName)
              { 
                 MessageBox.Show("Form1-M7: Invalid Program Name", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 LogFileOut.WriteLine("Form1-M7: Invalid Program Name");
                 InvalidAppOrDB = true;
              }
      
              if (dblThisAppVersion < dblAppVersion)
              { 
                 MessageBox.Show("Form1-M8: Invalid Program Version", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 LogFileOut.WriteLine("Form1-M8: Invalid Program Version - " + strThisAppVersion + " or Greater is Required!");
                 InvalidAppOrDB = true;
              }
           }
           else
           {
              MessageBox.Show("Form1-M9: Corrupt Database", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
              LogFileOut.WriteLine("Form1-M9: Corrupt Database");
              DBFatalErrorExit();
           }

           if ((dbMediaScan.DBConnState == ConnectionState.Open) && DBOpenedHere)
           {
              dbMediaScan.closeDB();
              DBOpenedHere = false;
           }

           if (InvalidAppOrDB)
              DBFatalErrorExit();
           else
           {
              txtMediaContent.AppendLine("Application " + strAppName + ".... Verified");
              txtMediaContent.AppendLine("Minimum Application Version Required " + strAppVersion + ".... Verified");
              txtMediaContent.AppendLine("Database " + strDBName + ".... Verified");
              txtMediaContent.AppendLine("Database Version " + strDBVersion + ".... Verified");
              LogFileOut.WriteLine("Application " + strAppName + ".... Verified");
              LogFileOut.WriteLine("Minimum Application Version " + strAppVersion + ".... Verified");
              LogFileOut.WriteLine("Database " + strDBName + ".... Verified");
              LogFileOut.WriteLine("Database Version " + strDBVersion + ".... Verified");
           }
        } // End of Form1_Load


        /********************************************************************************************
         * btnStart_Click
         * 
         * This is where the scanning and report generation begins.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnStart_Click(object sender, EventArgs e)
        {
            ResetCounters();
            ResetLists();
            CloseFiles();
            DeleteTempFiles();
            ResetFileNames();
            strPathToEliminate = string.Empty; 
            bDirectoryOnly = false;            

            dbMediaScan.openDB();

            // Get current date and time
            strDate = DateTime.Now.ToString("yyyyMMdd");
            strTime = DateTime.Now.ToString("HHmmss");

            txtMediaContent.Clear();
            txtMediaContent.AppendLine(strThisAppName + " " + strThisAppVersion);
            txtMediaContent.AppendLine("Validating....");

            if (!ValidateInputFields())
                return;

            // Determine if we are searching a drive or a specific directory
            // If the drive is not ready, display a warning message.
            try
            {
               FileAttributes Attr = File.GetAttributes(txtDriveToScan.Text);
               if (Attr.HasFlag(FileAttributes.Directory))
               {
                  string result = String.Empty;

                  result = Path.GetFileName(txtDriveToScan.Text.TrimEnd(Path.DirectorySeparatorChar));
                  if (String.IsNullOrEmpty(result))
                  {
                     bDirectoryOnly = false;
                     strPathToEliminate = "";
                  }
                  else
                  {
                     bDirectoryOnly = true;
                     strPathToEliminate = txtDriveToScan.Text;
                  }
               }
            }
            catch (Exception ex)
            {
               MessageBox.Show("Form1-M10: " + ex.Message + "\nInsert Media and/or Try Again");
            }

            // Path + Date + Time + Media Number
            strReportFileName = strPath + strDate + strTime + txtPA03TempNum.Text + ".txt";
            ReportFileOut = new StreamWriter(strReportFileName);
            strNoPathFileName = strPath + strDate + strTime + txtPA03TempNum.Text + "NoPath.txt";
            NoPathFileOut = new StreamWriter(strNoPathFileName);

            // Create Temp File
            strTempFileName = strPath + "Temp" + strDate + strTime + ".txt";
            TempFile = new StreamWriter(strTempFileName);

            PrintHeader();

            // Insert Header Data into the database
            intDBErrorCode = dbMediaScan.qryAddMedia(strDate, strTime, txtYourName.Text, txtUsername.Text, txtMediaBroughtInBy.Text, txtExistingMediaNum.Text,
            txtMediaTitle.Text, txtFromSystem.Text, txtPA03TempNum.Text, txtToSystem.Text, txtCheckedBy.Text, txtDTANum.Text, txtDriveToScan.Text, strReportFileName,
            strTempFileName, intFileCount, intDirCount, intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles, intNotUncompressCount, out intRecordID);

            if (intDBErrorCode == dbMediaScan.FatalError)
               DBFatalErrorExit();

            // Find all files at the specified path
            // Only display files that are NOT System Files or Hidden Files
            // Insert all files found into the Database
            DirSearch(txtDriveToScan.Text);

            GetSummary(txtDriveToScan.Text);

            // Update database with summary information
            intDBErrorCode = dbMediaScan.qryEditMedia(intRecordID, strDate, strTime, txtYourName.Text, txtUsername.Text, txtMediaBroughtInBy.Text, txtExistingMediaNum.Text,
            txtMediaTitle.Text, txtFromSystem.Text, txtPA03TempNum.Text, txtToSystem.Text, txtCheckedBy.Text, txtDTANum.Text, txtDriveToScan.Text, strReportFileName,
            strTempFileName, intFileCount, intDirCount, intZipFileCount, intTotEmbeddedZipFiles, intTotFilesInZipFiles, intNotUncompressCount);

            if (intDBErrorCode == dbMediaScan.FatalError)
               DBFatalErrorExit();

        } // End of btnStart_Click


        /********************************************************************************************
         * btnClear_Click
         * 
         * Clear the contents of the App's Window.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtMediaContent.Clear();
        } // End of btnClear_Click


        /********************************************************************************************
         * btnBrowseFolder_Click
         * 
         * Pops up a Folder selection dialog to allow the user to select the folder/drive he/she
         * would like to scan.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;

            // Show FoderBrowserDialog
            DialogResult Result = folderDlg.ShowDialog();
            if (Result == DialogResult.OK)
            {
                txtDriveToScan.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        } // End of btnBrowseFolder_Click


        /********************************************************************************************
         * btnPrintSummary_Click
         * 
         * Prints a Summary Report only, includes the header and summary information only.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnPrintSummary_Click(object sender, EventArgs e)
        {
            try
            {
                if (TempFile != null)
                {
                    if (TempFile.BaseStream != null)
                        TempFile.Close();
                }
                else
                {
                    MessageBox.Show("Form1-M11: Summary Report Does Not Exist, Press Start To Generate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                new PrintFile(strTempFileName, uintSummaryLineCount);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Form1-M12: " + ex.Message);
            }
        } // End of btnPrintSummary_Click


        /********************************************************************************************
         * btnPrintFull_Click
         * 
         * Prints the Full Report.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnPrintFull_Click(object sender, EventArgs e)
        {
           if (chkSupressPath.Checked && bDirectoryOnly)
           {
              try
              {
                 if (NoPathFileOut != null)
                 {
                    if (NoPathFileOut.BaseStream != null)
                       NoPathFileOut.Close();
                 }
                 else
                 {
                    MessageBox.Show("Form1-M13: Full Report with Suppress Path Does Not Exist, Press Start To Generate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                 }

                 new PrintFile(strNoPathFileName, uintReportLineCount);
              }
              catch (Exception ex)
              {
                 MessageBox.Show("Form1-M14: " + ex.Message);
              }

           }
           else
           {
              try
              {
                 if (ReportFileOut != null)
                 {
                    if (ReportFileOut.BaseStream != null)
                       ReportFileOut.Close();
                 }
                 else
                 {
                    MessageBox.Show("Form1-M15: Full Report Does Not Exist, Press Start To Generate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                 }

                 new PrintFile(strReportFileName, uintReportLineCount);
              }
              catch (Exception ex)
              {
                 MessageBox.Show("Form1-M16: " + ex.Message);
              }
           }
        } // End of btnPrintFull_Click


        /********************************************************************************************
         * btnExit_Click
         * 
         * Exits the program.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnExit_Click(object sender, EventArgs e)
        {
            try
            {
                if (ReportFileOut != null)
                    if (ReportFileOut.BaseStream != null)
                        ReportFileOut.Close();

                DeleteTempFiles();

                this.Close();
            }
            catch (IOException)
            {
                MessageBox.Show("Form1-17: Error Closing/Deleting Files On Exit", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogFileOut.WriteLine("Form1-M17: Error Closing/Deleting Files On Exit");
            }

            // Close this file last so we can record the above error if it occurs
            CloseLogFile();

        } // End of btnExit_Click


        /********************************************************************************************
         * Form1_FormClosing
         * 
         * Closes all files and deletes the temporary files.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (ReportFileOut != null)
                    if (ReportFileOut.BaseStream != null)
                        ReportFileOut.Close();

                DeleteTempFiles();
            }
            catch (IOException)
            {
                MessageBox.Show("Form1-M18: Error Closing/Deleting Files On Exit", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogFileOut.WriteLine("Form1-M18: Error Closing/Deleting Files On Exit");
            }

            // Close this file last so we can record the above error if it occurs
            CloseLogFile();

        } // End of Form1_FormClosing


        /********************************************************************************************
         * btnSaveDefaults_Click
         * 
         * Saves all input fields to the database. Each "Logged In User Name" is allowed one record 
         * in this table.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnSaveDefaults_Click(object sender, EventArgs e)
        {
           int intRecordsFound = 0;
           int intDefaultsID = 0;
           string strUserLoggedInName = String.Empty;
           string strUserName = String.Empty;
           string strMediaBroughtInBy = String.Empty;
           string strExistingMediaNum = String.Empty;
           string strMediaTitle = String.Empty;
           string strFromSystemAgency = String.Empty;
           string strPA03TempNum = String.Empty;
           string strToSystem = String.Empty;
           string strMediaCheckedBy = String.Empty;
           string strDTANumber = String.Empty;
           string strDriveToScan = String.Empty;

           bool DBOpenedHere = false;

           if (dbMediaScan.DBConnState == ConnectionState.Closed)
           {
              dbMediaScan.openDB();
              DBOpenedHere = true;
           }


           // Check to see if Defaults for the current user are already in the database.
           intRecordsFound = dbMediaScan.qryFindUserDefaults(txtUsername.Text, out intDefaultsID, out strUserLoggedInName, out strUserName, 
              out strMediaBroughtInBy, out strExistingMediaNum, out strMediaTitle, out strFromSystemAgency, out strPA03TempNum, out strToSystem, 
              out strMediaCheckedBy, out strDTANumber, out strDriveToScan);

           switch (intRecordsFound)
           {
              case 0: // No records found, add defaults to the database.
                  intDBErrorCode = dbMediaScan.qryAddDefaults(txtUsername.Text, txtYourName.Text, txtMediaBroughtInBy.Text, txtExistingMediaNum.Text,
                  txtMediaTitle.Text, txtFromSystem.Text, txtPA03TempNum.Text, txtToSystem.Text, txtCheckedBy.Text, txtDTANum.Text, txtDriveToScan.Text);

                  if (intDBErrorCode == dbMediaScan.FatalError)
                     DBFatalErrorExit();

                  LogFileOut.WriteLine("Defaults Saved");
                  txtMediaContent.AppendLine("Defaults Saved");
                  break;

              case 1: // Record found, update default values in the database.
                  intDBErrorCode = dbMediaScan.qryEditDefaults(intDefaultsID, txtUsername.Text, txtYourName.Text, txtMediaBroughtInBy.Text, txtExistingMediaNum.Text,
                  txtMediaTitle.Text, txtFromSystem.Text, txtPA03TempNum.Text, txtToSystem.Text, txtCheckedBy.Text, txtDTANum.Text, txtDriveToScan.Text);
                 if (intDBErrorCode == 1)
                 {
                    LogFileOut.WriteLine("Defaults Updated");
                    txtMediaContent.AppendLine("Defaults Updated");
                 }
                 else
                 {
                    LogFileOut.WriteLine("Defaults NOT Updated");
                    txtMediaContent.AppendLine("Defaults NOT Updated");
                 }
                  break;

              default: // Something went wrong!
                  MessageBox.Show("Form1-M19: Error Saving Defaults", "Save Defaults Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  LogFileOut.WriteLine("Form1-M19: Error Saving Defaults");
                  txtMediaContent.AppendLine("Form1-M19: Error Saving Defaults");
                  break;
           }

           if ((dbMediaScan.DBConnState == ConnectionState.Open) && DBOpenedHere)
           {
              dbMediaScan.closeDB();
              DBOpenedHere = false;
           }

        } // End of btnSaveDefaults_Click(object sender, EventArgs e)


        /********************************************************************************************
         * btnLoadDefaults_Click
         * 
         * Loads the previously saved input field information. The saved data is looked up by the  
         * "Logged In User Name".
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void btnLoadDefaults_Click(object sender, EventArgs e)
        {
           int intRecordsFound = 0;
           int intDefaultsID = 0;
           string strUserLoggedInName = String.Empty;
           string strUserName = String.Empty;
           string strMediaBroughtInBy = String.Empty;
           string strExistingMediaNum = String.Empty;
           string strMediaTitle = String.Empty;
           string strFromSystemAgency = String.Empty;
           string strPA03TempNum = String.Empty;
           string strToSystem = String.Empty;
           string strMediaCheckedBy = String.Empty;
           string strDTANumber = String.Empty;
           string strDriveToScan = String.Empty;

           bool DBOpenedHere = false;

           if (dbMediaScan.DBConnState == ConnectionState.Closed)
           {
              dbMediaScan.openDB();
              DBOpenedHere = true;
           }

           // Get Default values for the current user.
           intRecordsFound = dbMediaScan.qryFindUserDefaults(txtUsername.Text, out intDefaultsID, out strUserLoggedInName, out strUserName,
              out strMediaBroughtInBy, out strExistingMediaNum, out strMediaTitle, out strFromSystemAgency, out strPA03TempNum, out strToSystem,
              out strMediaCheckedBy, out strDTANumber, out strDriveToScan);

           switch (intRecordsFound)
           {
              case 0: // No records found for this machine name / user combination.
                 MessageBox.Show("No Default Values Found For " +  txtUsername.Text, "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                 break;

              case 1: // Record found, load the default values to the input fields.
                 txtYourName.Text = strUserName;
                 txtMediaBroughtInBy.Text = strMediaBroughtInBy;
                 txtExistingMediaNum.Text = strExistingMediaNum;
                 txtMediaTitle.Text = strMediaTitle;
                 txtFromSystem.Text = strFromSystemAgency;

                 // Do Not load default value for PA03TempNum! Force user input!
                 txtPA03TempNum.Text = String.Empty;

                 txtToSystem.Text = strToSystem;
                 txtCheckedBy.Text = strMediaCheckedBy;

                 // Do Not load default value for DriveToScan! Force user input!
                 txtDTANum.Text = String.Empty;

                 txtDriveToScan.Text = strDriveToScan;

                 break;

              default: // Something went wrong!
                 MessageBox.Show("Form1-M20: Error Loading Defaults", "Load Defaults Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 LogFileOut.WriteLine("Form1-M20: Error Loading Defaults");
                 break;
           }

           if ((dbMediaScan.DBConnState == ConnectionState.Open) && DBOpenedHere)
           {
              dbMediaScan.closeDB();
              DBOpenedHere = false;
           }

        } // End of btnLoadDefaults_Click(object sender, EventArgs e)

    } //End of partial class Form1
} // End of namespace MediaScan
