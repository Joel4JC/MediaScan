using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace DatabaseUpdater
{
   class Program
   {
      private static string strAppName = @"MediaScan";
      private static string strAppVersion = @"v1.1";
      private static double dblAppVersion = 1.1;
      private static string strDBName = @"dbMediaScan";
      private static string strDBVersion = @"v1.1";
      private static double dblDBVersion = 1.1;

      // Database
      private static OleDbConnection MediaScanConn;
      private static int FatalDBError = -1956;


      static private int FatalError
      {
         get
         {
            return FatalDBError;
         }
      }


      /********************************************************************************************
       * openDB
       * 
       * Opens the Database connection (opens the database).
       * 
       * Returns: Nothing
       * 
       ********************************************************************************************/
      static private void openDB()
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
      static private void closeDB()
      {
         if (MediaScanConn.State != ConnectionState.Closed)
         {
               MediaScanConn.Close();
         }
      } // End of closeDB()

      /********************************************************************************************
       * qryCreateTable
       * 
       * SQL Query to Create a Table in the database. 
       *                  
       * Returns: Number of successful rows inserted.
       * 
       ********************************************************************************************/
      static private int qryCreateTable()
      {
         int Rows = 0;
         string strCommand = String.Empty;

         strCommand = "CREATE TABLE tblDefaults ([DefaultsID] INTEGER IDENTITY(1,1) PRIMARY KEY NOT NULL, [UserLoginName] TEXT(255) UNIQUE NOT NULL, ";
         strCommand += "[UserName] TEXT(255), [MediaBroughtInBy] TEXT(255), [ExistingMediaNum] TEXT(255), [MediaTitle] TEXT(255), ";
         strCommand += "[FromSystemAgency] TEXT(255), [PA03TempNum] TEXT(255), [ToSystem] TEXT(255), [MediaCheckedBy] TEXT(255), ";
         strCommand += "[DTANumber] TEXT(255), [DriveToScan] TEXT(255))";

         try
         {
            OleDbCommand MediaScanCmd = new OleDbCommand();
            MediaScanCmd.Connection = MediaScanConn;
            MediaScanCmd.CommandText = strCommand;
            /* */
            /* */

            Rows = MediaScanCmd.ExecuteNonQuery();

         }
         catch (OleDbException ex)
         {
            for (int idx = 0; idx < ex.Errors.Count; idx++)
            {
               Console.WriteLine("   Error Code: " + ex.ErrorCode);
               Console.WriteLine("      Index #: " + idx);
               Console.WriteLine("Error Message: " + ex.Errors[idx].Message);
               Console.WriteLine("       Native: " + ex.Errors[idx].NativeError.ToString());
               Console.WriteLine("       Source: " + ex.Errors[idx].Source);
               Console.WriteLine("     SQLState: " + ex.Errors[idx].SQLState);
            }

            return (-99);
         }

         return (Rows);
      } // End of qryCreateTable()


      /********************************************************************************************
       * qryEditDBInfo
       * 
       * SQL Query to update the DBInfo table information.
       * 
       * Input Paramters: Values for each field in the database record to be updated.
       * 
       * Returns: The number of rows updated.
       * 
       ********************************************************************************************/
      static private int qryEditDBInfo()
      {
         int Rows = 0;
         string strCommand1 = String.Empty;
         string strCommand2 = String.Empty;
         string strCommand3 = String.Empty;

         string strDate = String.Empty;
         string strTime = String.Empty;

         // Get Current Date as mm/dd/yyyy, Time as hh:mm:ss
         DateTime CurrentDate = DateTime.Now;
         strDate = CurrentDate.ToShortDateString();
         strTime = CurrentDate.ToLongTimeString();

         // First Alter the table then update the information
         strCommand1 = "ALTER TABLE [tblDBInfo] ADD COLUMN [AppVersionNumber] DOUBLE";

         strCommand2 = "ALTER TABLE [tblDBInfo] ADD COLUMN [DBVersionNumber] DOUBLE";

         strCommand3 = "UPDATE [tblDBInfo] SET [DBVersion]=?, [DBName]=?, [AppName]=?, [AppVersion]=?, [AppVersionNumber]=?, [DBVersionNumber]=? ";
         strCommand3 += "WHERE ID = ?"; // 'ID' may need to be changed to 'DBInfoID'

         try
         {
            // openDB();

            OleDbCommand MediaScanCmd = new OleDbCommand();
            MediaScanCmd.Connection = MediaScanConn;

            MediaScanCmd.CommandText = strCommand1;
            Rows = MediaScanCmd.ExecuteNonQuery();

            MediaScanCmd.CommandText = strCommand2;
            Rows = MediaScanCmd.ExecuteNonQuery();

            MediaScanCmd.CommandText = strCommand3;
            MediaScanCmd.Parameters.AddWithValue("@p1", strDBVersion);
            MediaScanCmd.Parameters.AddWithValue("@p2", strDBName);
            MediaScanCmd.Parameters.AddWithValue("@p3", strAppName);
            MediaScanCmd.Parameters.AddWithValue("@p4", strAppVersion);
            MediaScanCmd.Parameters.AddWithValue("@p5", dblAppVersion);
            MediaScanCmd.Parameters.AddWithValue("@p6", dblDBVersion);

            MediaScanCmd.Parameters.AddWithValue("@p7", 1);

            Rows = MediaScanCmd.ExecuteNonQuery();

            // closeDB();
         }
         catch (OleDbException ex)
         {
            for (int idx = 0; idx < ex.Errors.Count; idx++)
            {
               Console.WriteLine("   Error Code: " + ex.ErrorCode);
               Console.WriteLine("      Index #: " + idx);
               Console.WriteLine("Error Message: " + ex.Errors[idx].Message);
               Console.WriteLine("       Native: " + ex.Errors[idx].NativeError.ToString());
               Console.WriteLine("       Source: " + ex.Errors[idx].Source);
               Console.WriteLine("     SQLState: " + ex.Errors[idx].SQLState);
            }

            return (-99);
         }

         return (Rows);
      } // End of qryEditDBInfo()


      static void Main(string[] args)
      {
         int intDBResults;
         MediaScanConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:\MediaScan\dbMediaScan.accdb");

         openDB();

         intDBResults = qryCreateTable();

         if (intDBResults == 0)
            Console.WriteLine("Table Created Successfully!");
         else
            Console.WriteLine("Table Creation FAILED!");

         intDBResults = qryEditDBInfo();

         if (intDBResults == 1)
            Console.WriteLine("Database Info Updated Successfully!");
         else
            Console.WriteLine("Database Info Update FAILED!");

         closeDB();

         Console.WriteLine("\nPress Enter To Continue\n");
         Console.Read();

      } // End of Main
   } // End of class Program
} // End of namespace DatabaseUpdater
