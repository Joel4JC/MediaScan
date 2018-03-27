using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace MediaScan
{
    class Database
    {
        OleDbConnection MediaScanConn;
        private static int FatalDBError = -1956;

       public int FatalError
        {
          get
           {
              return FatalDBError;
           }
        }

       public ConnectionState DBConnState
       {
          get
          {
             return MediaScanConn.State;
          }
       }

        /********************************************************************************************
         * Database
         * 
         * Contructor which Creates the Database connection with the Access Database
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        public Database()
        {
            MediaScanConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\MediaScan\dbMediaScan.accdb");
        } // End Database()


        /********************************************************************************************
         * openDB
         * 
         * Opens the Database connection (opens the database).
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        public void openDB()
        {
            if (MediaScanConn.State != ConnectionState.Open)
            {
                MediaScanConn.Open();
            }
        } // End of openDB()


        /********************************************************************************************
         * closeDB
         * 
         * Closes the Database connection (closes the database)
         * 
         * Returns: Nothing
         ********************************************************************************************/
        public void closeDB()
        {
            if (MediaScanConn.State != ConnectionState.Closed)
            {
                MediaScanConn.Close();
            }
        } // End of closeDB()


        /********************************************************************************************
         * qryFindDBInfo
         * 
         * A SQL Query to retrieve the information for a given Logged In User.
         * 
         * Output Paramters: Values for each field in the database record to be returned from the 
         *                   database query.
         * 
         * Returns: Number of rows meeting the search criteria.
         * 
         ********************************************************************************************/
        public int qryFindDBInfo(out int intDBInfoID, out string strDBVersion, out string strDBName, out string strAppName, out string strAppVersion,
           out double dblAppVersion, out double dblDBVersion)
        {
           int Count = 0;
           string strCommand = String.Empty;

           // NOTE: Some copies of dbMediaScan.acdb will have 'ID' and some will have 'DBInfoID'
           strCommand = "SELECT * FROM [tblDBInfo] WHERE [ID]=" + 1; // May need to be changed to 'ID' or 'DBInfoID'

           intDBInfoID = 0;
           strDBVersion = strDBName = strAppName = strAppVersion = String.Empty;
           dblAppVersion = dblDBVersion = 0;

           try
           {
              OleDbCommand MediaScanCmd = new OleDbCommand();
              MediaScanCmd.Connection = MediaScanConn;
              MediaScanCmd.CommandText = strCommand;
              OleDbDataReader MediaScanReader = MediaScanCmd.ExecuteReader();

              while (MediaScanReader.Read())
              {
                 Count++;
                 intDBInfoID = Convert.ToInt32(MediaScanReader["ID"]); // May need to be changed to 'ID' or 'DBInfoID'
                 strDBVersion = (MediaScanReader["DBVersion"].ToString());
                 strDBName = (MediaScanReader["DBName"].ToString());
                 strAppName = (MediaScanReader["AppName"].ToString());
                 strAppVersion = (MediaScanReader["AppVersion"].ToString());
                 dblAppVersion = Convert.ToDouble(MediaScanReader["AppVersionNumber"]);
                 dblDBVersion = Convert.ToDouble(MediaScanReader["DBVersionNumber"]);
              }
           }
           catch (Exception ex)
           {
              MessageBox.Show(ex.Message, "Database-M0: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
              MessageBox.Show(strCommand);
              return (FatalDBError);
           }

           return (Count);
        } // End of qryFindDBInfo(string strUserToFind, out string DefaultsID,...


        /********************************************************************************************
         * qryFindTempNum
         * 
         * A SQL Query to retrieve the information for a given media Temp Number.
         * 
         * Output Paramters: Values for each field in the database record to be returned from the 
         *                   database query.
         * 
         * Returns: Number of rows meeting the search criteria.
         * 
         ********************************************************************************************/
        public int qryFindTempNum(out string MediaID, out string strDateRan, out string strTimeRan, out string strUserN, out string strLogInName, out string strBroughtIn, 
            out string strExistMediaNum, out string strMediaTitle, out string strFromSys, ref string strTempNum, out string strToSys, out string strCheckBy,
            out string strDTA, out string strDriveScan, out string strRptFileName, out string strTmpFileName, out int intTotFiles, out int intTotDir, out int intTotZip,
            out int intTotEmbedZip, out int intTotFilesInZip, out int intTotOther)
        {
            int Count = 0;
            string strCommand = "SELECT * FROM tblMedia WHERE PA03TempNum='" + strTempNum + "'";

            MediaID = strDateRan = strTimeRan = strUserN = strLogInName = strBroughtIn = strExistMediaNum = strMediaTitle = strFromSys = strToSys = strCheckBy = "";
            strDTA = strDriveScan = strRptFileName = strTmpFileName = "";
            intTotFiles = intTotDir = intTotZip = intTotEmbedZip = intTotFilesInZip = intTotOther = 0;

            try
            {
                // openDB();

                OleDbCommand MediaScanCmd = new OleDbCommand();
                MediaScanCmd.Connection = MediaScanConn;
                MediaScanCmd.CommandText = strCommand;
                OleDbDataReader MediaScanReader =MediaScanCmd.ExecuteReader();

                while (MediaScanReader.Read())
                {
                    Count++;
                    MediaID = (MediaScanReader["MeiaID"].ToString());
                    strDateRan = (MediaScanReader["DateRan"].ToString());
                    strTimeRan = (MediaScanReader["TimeRan"].ToString());
                    strUserN = (MediaScanReader["UserName"].ToString());
                    strLogInName = (MediaScanReader["UserLoginName"].ToString());
                    strBroughtIn = (MediaScanReader["MediaBroughtInBy"].ToString());
                    strExistMediaNum = (MediaScanReader["ExistingMediaNum"].ToString());
                    strMediaTitle = (MediaScanReader["MediaTitle"].ToString());
                    strFromSys = (MediaScanReader["FromSystemAgency"].ToString());
                    strTempNum = (MediaScanReader["PA03TempNum"].ToString());
                    strToSys = (MediaScanReader["ToSystem"].ToString());
                    strCheckBy = (MediaScanReader["MediaCheckedBy"].ToString());
                    strDTA = (MediaScanReader["DTANumber"].ToString());
                    strDriveScan = (MediaScanReader["DriveToScan"].ToString());
                    strRptFileName = (MediaScanReader["ReportFileName"].ToString());
                    strTmpFileName = (MediaScanReader["TempFileName"].ToString());
                    intTotFiles = Convert.ToInt32(MediaScanReader["TotalFilesOnMedia"]);
                    intTotDir = Convert.ToInt32(MediaScanReader["TotalDirectoriesOnMedia"]);
                    intTotZip = Convert.ToInt32(MediaScanReader["TotalZipFilesOnMedia"]);
                    intTotEmbedZip = Convert.ToInt32(MediaScanReader["TotalEmbeddedZipFiles"]);
                    intTotFilesInZip = Convert.ToInt32(MediaScanReader["TotalFilesInZipFiles"]);
                    intTotOther = Convert.ToInt32(MediaScanReader["TotalOtherCompressFiles"]);
                }

                // closeDB();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Database-M1: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(strCommand);
            }

            return (Count);
        } // End of qryFindTempNum(out string MediaID, out string strDateRan, out string strTimeRan,...


        /********************************************************************************************
         * qryAddMedia
         * 
         * SQL Query to Add Media information into the database.
         * 
         * Input Paramters: Values for each field in the database record to be added/inserted into
         *                  the database.
         *                  
         * Output Parameter: The MediaID Key Value of the record just inserted into the database.
         *                   This will be used to update the database record when the summary data
         *                   is calculated at the end of the scan process.
         * 
         * Returns: Number of successful rows inserted.
         * 
         ********************************************************************************************/
        public int qryAddMedia(string strDateRan, string strTimeRan, string strUserN, string strLogInName, string strBroughtIn, string strExistMediaNum,
            string strMediaTitle, string strFromSys, string strTempNum, string strToSys, string strCheckBy, string strDTA, string strDriveScan, string strRptFileName,
            string strTmpFileName, int intTotFiles, int intTotDir, int intTotZip, int intTotEmbedZip, int intTotFilesInZipFiles, int intTotOther, out int intMediaID)
        {
            int Rows = 0;
            int intRecordID = 0;
            string strCommand = String.Empty;
            string strDate = String.Empty;
            string strTime = String.Empty;

            // Get Current Date as mm/dd/yyyy, Time as hh:mm:ss 
            DateTime CurrentDate = DateTime.Now;
            strDate = CurrentDate.ToShortDateString();
            strTime = CurrentDate.ToLongTimeString();

            strCommand = "INSERT INTO [tblMedia] ([DateRan], [TimeRan], [UserName], [UserLoginName], [MediaBroughtInBy], [ExistingMediaNum], [MediaTitle], ";
            strCommand += "[FromSystemAgency], [PA03TempNum], [ToSystem], [MediaCheckedBy], [DTANumber], [DriveToScan], [ReportFileName], [TempFileName], ";
            strCommand += "[TotalFilesOnMedia], [TotalDirectoriesOnMedia], [TotalZipFilesOnMedia], [TotalEmbeddedZipFiles], [TotalFilesInZipFiles], [TotalOtherCompressFiles]) ";
            strCommand += "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            
            try
            {
                // openDB();

                OleDbCommand MediaScanCmd = new OleDbCommand();
                MediaScanCmd.Connection = MediaScanConn;
                MediaScanCmd.CommandText = strCommand;
                /* */
                MediaScanCmd.Parameters.AddWithValue("[DateRan]", strDateRan);
                MediaScanCmd.Parameters.AddWithValue("[TimeRan]", strTimeRan);
                MediaScanCmd.Parameters.AddWithValue("[UserName]", strUserN);
                MediaScanCmd.Parameters.AddWithValue("[UserLoginName]", strLogInName);
                MediaScanCmd.Parameters.AddWithValue("[MediaBroughtInBy]", strBroughtIn);
                MediaScanCmd.Parameters.AddWithValue("[ExistingMediaNum]", strExistMediaNum);
                MediaScanCmd.Parameters.AddWithValue("[MediaTitle]", strMediaTitle);
                MediaScanCmd.Parameters.AddWithValue("[FromSystemAgency]", strFromSys);
                MediaScanCmd.Parameters.AddWithValue("[PA03TempNum]", strTempNum);
                MediaScanCmd.Parameters.AddWithValue("[ToSystem]", strToSys);
                MediaScanCmd.Parameters.AddWithValue("[MediaCheckedBy]", strCheckBy);
                MediaScanCmd.Parameters.AddWithValue("[DTANumber]", strDTA);
                MediaScanCmd.Parameters.AddWithValue("[DriveToScan]", strDriveScan);
                MediaScanCmd.Parameters.AddWithValue("[ReportFileName]", strRptFileName);
                MediaScanCmd.Parameters.AddWithValue("[TempFileName]", strTmpFileName);
                MediaScanCmd.Parameters.AddWithValue("[TotalFilesOnMedia]", intTotFiles);
                MediaScanCmd.Parameters.AddWithValue("[TotalDirectoriesOnMedia]", intTotDir);
                MediaScanCmd.Parameters.AddWithValue("[TotalZipFilesOnMedia]", intTotZip);
                MediaScanCmd.Parameters.AddWithValue("[TotalEmbeddedZipFiles]", intTotEmbedZip);
                MediaScanCmd.Parameters.AddWithValue("[TotalFilesInZipFiles]", intTotFilesInZipFiles);
                MediaScanCmd.Parameters.AddWithValue("[TotalOtherCompressFiles]", intTotOther);
                /* */

                Rows = MediaScanCmd.ExecuteNonQuery();

                // Get Record ID
                strCommand = "SELECT @@IDENTITY";
                MediaScanCmd.CommandText = strCommand;
                intRecordID = (int)MediaScanCmd.ExecuteScalar();

                // closeDB();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database-M2: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(strCommand);
                intMediaID = -1;
                return (FatalDBError);
            }
            intMediaID = intRecordID;
            return (Rows);
        } // End of qryAddMedia(string MediaID, string strDateRan, string strTimeRan, string strUserN, string strLogInName,...


        /********************************************************************************************
         * qryEditMedia
         * 
         * SQL Query to update the Media table information.
         * 
         * Input Paramters: Values for each field in the database record to be updated.
         * 
         * Returns: The number of rows updated.
         * 
         ********************************************************************************************/
        public int qryEditMedia(int intMediaID, string strDateRan, string strTimeRan, string strUserN, string strLogInName, string strBroughtIn, string strExistMediaNum,
            string strMediaTitle, string strFromSys, string strTempNum, string strToSys, string strCheckBy, string strDTA, string strDriveScan, string strRptFileName,
            string strTmpFileName, int intTotFiles, int intTotDir, int intTotZip, int intTotEmbedZip, int intTotFilesInZipFiles, int intTotOther)
        {
            int Rows = 0;
            string strCommand = String.Empty;
            string strDate = String.Empty;
            string strTime = String.Empty;

            // Get Current Date as mm/dd/yyyy, Time as hh:mm:ss
            DateTime CurrentDate = DateTime.Now;
            strDate = CurrentDate.ToShortDateString();
            strTime = CurrentDate.ToLongTimeString();

            strCommand = "UPDATE [tblMedia] SET [DateRan]=?, [TimeRan]=?, [UserName]=?, [UserLoginName]=?, [MediaBroughtInBy]=?, [ExistingMediaNum]=?, [MediaTitle]=?, ";
            strCommand += "[FromSystemAgency]=?, [PA03TempNum]=?, [ToSystem]=?, [MediaCheckedBy]=?, [DTANumber]=?, [DriveToScan]=?, [ReportFileName]=?, [TempFileName]=?, ";
            strCommand += "[TotalFilesOnMedia]=?, [TotalDirectoriesOnMedia]=?, [TotalZipFilesOnMedia]=?, [TotalEmbeddedZipFiles]=?, [TotalFilesInZipFiles]=?, ";
            strCommand += "[TotalOtherCompressFiles]=? ";
            strCommand += "WHERE MediaID = ?";

            try
            {
                // openDB();

                OleDbCommand MediaScanCmd = new OleDbCommand();
                MediaScanCmd.Connection = MediaScanConn;
                MediaScanCmd.CommandText = strCommand;
                MediaScanCmd.Parameters.AddWithValue("@p1", strDateRan);
                MediaScanCmd.Parameters.AddWithValue("@p2", strTimeRan);
                MediaScanCmd.Parameters.AddWithValue("@p3", strUserN);
                MediaScanCmd.Parameters.AddWithValue("@p4", strLogInName);
                MediaScanCmd.Parameters.AddWithValue("@p5", strBroughtIn);
                MediaScanCmd.Parameters.AddWithValue("@p6", strExistMediaNum);
                MediaScanCmd.Parameters.AddWithValue("@p7", strMediaTitle);
                MediaScanCmd.Parameters.AddWithValue("@p8", strFromSys);
                MediaScanCmd.Parameters.AddWithValue("@p9", strTempNum);
                MediaScanCmd.Parameters.AddWithValue("@p10", strToSys);
                MediaScanCmd.Parameters.AddWithValue("@p11", strCheckBy);
                MediaScanCmd.Parameters.AddWithValue("@p12", strDTA);
                MediaScanCmd.Parameters.AddWithValue("@p13", strDriveScan);
                MediaScanCmd.Parameters.AddWithValue("@p14", strRptFileName);
                MediaScanCmd.Parameters.AddWithValue("@p15", strTmpFileName);
                MediaScanCmd.Parameters.AddWithValue("@p16", intTotFiles);
                MediaScanCmd.Parameters.AddWithValue("@p17", intTotDir);
                MediaScanCmd.Parameters.AddWithValue("@p18", intTotZip);
                MediaScanCmd.Parameters.AddWithValue("@p19", intTotEmbedZip);
                MediaScanCmd.Parameters.AddWithValue("@p20", intTotFilesInZipFiles);
                MediaScanCmd.Parameters.AddWithValue("@p21", intTotOther);


                MediaScanCmd.Parameters.AddWithValue("@p22", intMediaID);

                Rows = MediaScanCmd.ExecuteNonQuery();

                // closeDB();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database-M3: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(strCommand);
                intMediaID = -1;
                return (FatalDBError);
            }
            return (Rows);
        } // End of qryEditMedia(int MediaID, string strDateRan, string strTimeRan, string strUserN ....


        /********************************************************************************************
         * qryAddFile
         * 
         * SQL Query to Add File information into the database. Each file found in the scan process
         * is added to the database.
         * 
         * Input Paramters: Values for each field in the database record to be added/inserted into
         *                  the database.
         *                  
         * Returns: Number of successful rows inserted.
         * 
         ********************************************************************************************/
        public int qryAddFile(int intMediaID, int intReportLineNumber, string strFilePath, string strFileName, int intIndentSize)
        {
           int Rows = 0;
           string strCommand = String.Empty;

           strCommand = "INSERT INTO [tblFiles] ([MediaID], [ReportLineNumber], [FilePath], [FileName], [IndentSize]) VALUES (?, ?, ?, ?, ?)";

           try
           {
              // openDB();

              OleDbCommand MediaScanCmd = new OleDbCommand();
              MediaScanCmd.Connection = MediaScanConn;
              MediaScanCmd.CommandText = strCommand;
              /* */
              MediaScanCmd.Parameters.AddWithValue("[MediaID]", intMediaID);
              MediaScanCmd.Parameters.AddWithValue("[ReportLineNumber]", intReportLineNumber);
              MediaScanCmd.Parameters.AddWithValue("[FilePath]", strFilePath);
              MediaScanCmd.Parameters.AddWithValue("[FileName]", strFileName);
              MediaScanCmd.Parameters.AddWithValue("[IndentSize]", intIndentSize);
              /* */

              Rows = MediaScanCmd.ExecuteNonQuery();

              // closeDB();
           }
           catch (Exception ex)
           {
              MessageBox.Show(ex.Message, "Database-M4: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
              MessageBox.Show(strCommand);
              return (FatalDBError);
           }

           return (Rows);
        } // End of qryAddFile(int MediaID, string intReportLine...


        /********************************************************************************************
         * qryAddDefaults
         * 
         * SQL Query to Add Defaults values to the input screen, which were previously saved in the
         * database.
         * 
         * Input Paramters: Values for each field in the database record to be added/inserted into
         *                  the database.
         *                  
         * Returns: Number of successful rows inserted.
         * 
         ********************************************************************************************/
        public int qryAddDefaults(string strUserLoggedInName, string strUserName, string strMediaBroughtInBy, string strExistingMediaNum,
           string strMediaTitle, string strFromSystemAgency, string strPA03TempNum, string strToSystem, string strMediaCheckedBy, string strDTANumber,
           string strDriveToScan)
        {
           int Rows = 0;
           string strCommand = String.Empty;

           strCommand = "INSERT INTO [tblDefaults] ([UserLoginName], [UserName], [MediaBroughtInBy], [ExistingMediaNum], [MediaTitle], ";
           strCommand += "[FromSystemAgency], [PA03TempNum], [ToSystem], [MediaCheckedBy], [DTANumber], [DriveToScan]) ";
           strCommand += "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

           try
           {
              // openDB();

              OleDbCommand MediaScanCmd = new OleDbCommand();
              MediaScanCmd.Connection = MediaScanConn;
              MediaScanCmd.CommandText = strCommand;
              /* */
              MediaScanCmd.Parameters.AddWithValue("[UserLoginName]", strUserLoggedInName);
              MediaScanCmd.Parameters.AddWithValue("[UserName]", strUserName);
              MediaScanCmd.Parameters.AddWithValue("[MediaBroughtInBy]", strMediaBroughtInBy);
              MediaScanCmd.Parameters.AddWithValue("[ExistingMediaNum]", strExistingMediaNum);
              MediaScanCmd.Parameters.AddWithValue("[MediaTitle]", strMediaTitle);
              MediaScanCmd.Parameters.AddWithValue("[FromSystemAgency]", strFromSystemAgency);
              MediaScanCmd.Parameters.AddWithValue("[PA03TempNum]", strPA03TempNum);
              MediaScanCmd.Parameters.AddWithValue("[ToSystem]", strToSystem);
              MediaScanCmd.Parameters.AddWithValue("[MediaCheckedBy]", strMediaCheckedBy);
              MediaScanCmd.Parameters.AddWithValue("[DTANumber]", strDTANumber);
              MediaScanCmd.Parameters.AddWithValue("[DriveToScan]", strDriveToScan);
              /* */

              Rows = MediaScanCmd.ExecuteNonQuery();

              // closeDB();
           }
           catch (Exception ex)
           {
              MessageBox.Show(ex.Message, "Database-M5: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
              MessageBox.Show(strCommand);
              return (FatalDBError);
           }

           return (Rows);
        } // End of qryAddDefaults(int MediaID, string intReportLine...


        /********************************************************************************************
         * qryFindUserDefaults
         * 
         * A SQL Query to retrieve the information for a given Logged In User.
         * 
         * Output Paramters: Values for each field in the database record to be returned from the 
         *                   database query.
         * 
         * Returns: Number of rows meeting the search criteria.
         * 
         ********************************************************************************************/
        public int qryFindUserDefaults(string strUserToFind, out int intDefaultsID, out string strUserLoggedInName, out string strUserName, out string strMediaBroughtInBy,
           out string strExistingMediaNum, out string strMediaTitle, out string strFromSystemAgency, out string strPA03TempNum, out string strToSystem,
           out string strMediaCheckedBy, out string strDTANumber, out string strDriveToScan)
        {
           int Count = 0;
           string strCommand = "SELECT * FROM [tblDefaults] WHERE [UserLoginName]='" + strUserToFind + "'";

           intDefaultsID = 0;
           strUserLoggedInName = strUserName = strMediaBroughtInBy = strExistingMediaNum = strMediaTitle = strFromSystemAgency = String.Empty;
           strPA03TempNum = strToSystem = strMediaCheckedBy = strDTANumber = strDriveToScan = String.Empty;

           try
           {
              OleDbCommand MediaScanCmd = new OleDbCommand();
              MediaScanCmd.Connection = MediaScanConn;
              MediaScanCmd.CommandText = strCommand;
              OleDbDataReader MediaScanReader = MediaScanCmd.ExecuteReader();

              while (MediaScanReader.Read())
              {
                 Count++;
                 intDefaultsID = Convert.ToInt32(MediaScanReader["DefaultsID"]);
                 strUserLoggedInName = (MediaScanReader["UserLoginName"].ToString());
                 strUserName = (MediaScanReader["UserName"].ToString());
                 strMediaBroughtInBy = (MediaScanReader["MediaBroughtInBy"].ToString());
                 strExistingMediaNum = (MediaScanReader["ExistingMediaNum"].ToString());
                 strMediaTitle = (MediaScanReader["MediaTitle"].ToString());
                 strFromSystemAgency = (MediaScanReader["FromSystemAgency"].ToString());
                 strPA03TempNum = (MediaScanReader["PA03TempNum"].ToString());
                 strToSystem = (MediaScanReader["ToSystem"].ToString());
                 strMediaCheckedBy = (MediaScanReader["MediaCheckedBy"].ToString());
                 strDTANumber = (MediaScanReader["DTANumber"].ToString());
                 strDriveToScan = (MediaScanReader["DriveToScan"].ToString());
              }
           }
           catch (Exception ex)
           {
              MessageBox.Show(ex.Message, "Database-M6: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
              MessageBox.Show(strCommand);
              return (FatalDBError);
           }

           return (Count);
        } // End of qryFindUserDefaults(string strUserToFind, out string DefaultsID,...


        /********************************************************************************************
         * qryEditDefaults
         * 
         * SQL Query to update the Defaults table information.
         * 
         * Input Paramters: Values for each field in the database record to be updated.
         * 
         * Returns: The number of rows updated.
         * 
         ********************************************************************************************/
        public int qryEditDefaults(int intDefaultsID, string strUserLoggedInName, string strUserName, string strMediaBroughtInBy,
           string strExistingMediaNum, string strMediaTitle, string strFromSystemAgency, string strPA03TempNum, string strToSystem,
           string strMediaCheckedBy, string strDTANumber, string strDriveToScan)
        {
           int Rows = 0;
           string strCommand = String.Empty;
           string strDate = String.Empty;
           string strTime = String.Empty;

           // Get Current Date as mm/dd/yyyy, Time as hh:mm:ss
           DateTime CurrentDate = DateTime.Now;
           strDate = CurrentDate.ToShortDateString();
           strTime = CurrentDate.ToLongTimeString();

           strCommand = "UPDATE [tblDefaults] SET [UserName]=?, [MediaBroughtInBy]=?, [ExistingMediaNum]=?, [MediaTitle]=?, ";
           strCommand += "[FromSystemAgency]=?, [PA03TempNum]=?, [ToSystem]=?, [MediaCheckedBy]=?, [DTANumber]=?, [DriveToScan]=? ";
           strCommand += "WHERE [DefaultsID] = ?";

           try
           {
              // openDB();

              OleDbCommand MediaScanCmd = new OleDbCommand();
              MediaScanCmd.Connection = MediaScanConn;
              MediaScanCmd.CommandText = strCommand;
              MediaScanCmd.Parameters.AddWithValue("@p1", strUserName);
              MediaScanCmd.Parameters.AddWithValue("@p2", strMediaBroughtInBy);
              MediaScanCmd.Parameters.AddWithValue("@p3", strExistingMediaNum);
              MediaScanCmd.Parameters.AddWithValue("@p4", strMediaTitle);
              MediaScanCmd.Parameters.AddWithValue("@p5", strFromSystemAgency);
              MediaScanCmd.Parameters.AddWithValue("@p6", strPA03TempNum);
              MediaScanCmd.Parameters.AddWithValue("@p7", strToSystem);
              MediaScanCmd.Parameters.AddWithValue("@p8", strMediaCheckedBy);
              MediaScanCmd.Parameters.AddWithValue("@p9", strDTANumber);
              MediaScanCmd.Parameters.AddWithValue("@p10", strDriveToScan);


              MediaScanCmd.Parameters.AddWithValue("@p11", intDefaultsID);

              Rows = MediaScanCmd.ExecuteNonQuery();

              // closeDB();
           }
           catch (Exception ex)
           {
              MessageBox.Show(ex.Message, "Database-M7: ", MessageBoxButtons.OK, MessageBoxIcon.Error);
              MessageBox.Show(strCommand);
              intDefaultsID = -1;
              return (FatalDBError);
           }
           return (Rows);
        } // End of qryEditDefaults(int intDefaultsID, string strUserLoggedInName....

       
        /********************************************************************************************
         * releaseDB
         * 
         * Once the database connection is closed, its resources can be released.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        public void releaseDB()
        {
            MediaScanConn.Dispose();
        } // End of releaseDB()

    } // End of class Database
} // End of namespace MediaScan
