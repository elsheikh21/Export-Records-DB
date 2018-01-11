using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Configuration;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace exportSelected {
    class Program {



        static void Main(string[] args) {
            if( select() > 0 ) {
                if( select1() > 0 ) {
                    writeLog( String.Format( "Successfully done on ({0})", getTime() ) );
                }
            }
        }

        public static int select() {
            int noOfRowsChanged = 0;
            string userID = ConfigurationManager.AppSettings["userID"].ToString();
            string passWord = ConfigurationManager.AppSettings["password"].ToString();
            string dbName = ConfigurationManager.AppSettings["DataBaseName"].ToString();
            string IPAddress = ConfigurationManager.AppSettings["IPAddress"].ToString();
            SqlConnection MSSQLConn = new SqlConnection();
            MSSQLConn.ConnectionString = String.Format( "User ID={0};Initial Catalog={1};Data Source={2}; password={3}", userID, dbName, IPAddress, passWord );
            //this.MSSQLConn.ConnectionString = String.Format( "User ID={0};Initial Catalog={1};Data Source={2}; password={3}", this.username, this.initialCatalog, this.dataSource, this.password );
            //string connectionString = @"Data Source=192.168.1.17,1776; Network Library=DBMSSOCN; Initial Catalog=AssetTracking; User ID=sa; Password=Pass@1234";
            // [AttendanceHis] WHERE Logdate LIKE '%2017%' ORDER BY Logdate DESC
            string selectStatment = "SELECT * FROM [AttendanceHis] WHERE Logdate LIKE '%2017%' ORDER BY Logdate DESC";

            //SqlConnection MSSQLConn= new SqlConnection( MSSQLConn.ConnectionString );
            SqlCommand comm = new SqlCommand( selectStatment, MSSQLConn );

            SqlDataAdapter da = null;
            DataTable dt = new DataTable();

            try {

                MSSQLConn.Open();
                da = new SqlDataAdapter( selectStatment, MSSQLConn );
                noOfRowsChanged = da.Fill( dt );

            } catch( Exception exc ) {

                writeLog( String.Format( "SELECTING ERROR IS DUE TO: {0}", exc.ToString() ) );

            } finally {
                //da.Dispose();
                MSSQLConn.Close();

            }
            if( noOfRowsChanged > 0 ) {
                exportCSV( dt, 1 );
            }
            return noOfRowsChanged;
        }

        public static int select1() {
            int noOfRowsChanged = 0;

            string userID = ConfigurationManager.AppSettings["userID"].ToString();
            string passWord = ConfigurationManager.AppSettings["password"].ToString();
            string dbName = ConfigurationManager.AppSettings["DataBaseName"].ToString();
            string IPAddress = ConfigurationManager.AppSettings["IPAddress"].ToString();
            // HRPDC
            SqlConnection MSSQLConn = new SqlConnection();
            MSSQLConn.ConnectionString = String.Format( "User ID={0};Initial Catalog={1};Data Source={2}; password={3}", userID, dbName, IPAddress, passWord );

            //string connectionString = String.Format( @"Data Source=192.168.1.17\\SQLEXPRESS; Network Library=DBMSSOCN; Initial Catalog=AssetTracking; User ID={0}; Password={1}", userID, passWord );
            // [Error_Records] WHERE line LIKE '%2017%' ORDER BY Line DESC
            string selectStatment = "SELECT * FROM [Error_Records] WHERE line LIKE '%2017%' ORDER BY Line DESC";

            //SqlConnection conn = new SqlConnection( MSSQLConn.ConnectionString );
            SqlCommand comm = new SqlCommand( selectStatment, MSSQLConn );

            SqlDataAdapter da = null;
            DataTable dt = new DataTable();

            try {

                MSSQLConn.Open();
                da = new SqlDataAdapter( selectStatment, MSSQLConn );
                noOfRowsChanged = da.Fill( dt );

            } catch( Exception exc ) {

                writeLog( String.Format( "SELECTING1 ERROR IS DUE TO: {0}", exc.ToString() ) );

            } finally {
                //da.Dispose();
                MSSQLConn.Close();
            }
            if( noOfRowsChanged > 0 ) {
                exportCSV( dt, 2 );
            }
            return noOfRowsChanged;
        }

        public static void writeLog(String text) {
            string dirName = Path.GetDirectoryName( Assembly.GetExecutingAssembly().GetName().CodeBase ) + "\\log.txt";
            string localPath = new Uri( dirName ).LocalPath;
            using( StreamWriter sw = new StreamWriter( localPath, true ) ) {
                sw.Write( text + Environment.NewLine );
            }
        }

        public static void exportCSV(DataTable dt, int which) {
            try {
                string name = String.Empty;
                StreamWriter sw = null;
                string dirName = Path.GetDirectoryName( Assembly.GetExecutingAssembly().GetName().CodeBase ) + "\\Logs\\log.txt";
                string path = new Uri( dirName ).LocalPath;


                string dirName1 = Path.GetDirectoryName( Assembly.GetExecutingAssembly().GetName().CodeBase );
                string path1 = new Uri( dirName1 ).LocalPath;

                if( which == 1 ) {
                    name = String.Format( "{0}\\XXXX.csv", path1 );
                }
                if( which == 2 ) {
                    name = String.Format( "{0}\\YYYY.csv", path1 );
                }
                if( dt.Rows.Count > 0 ) {
                    sw = new StreamWriter( name, false );
                    int iColCount = dt.Columns.Count;
                    for( int i = 0; i < iColCount; i++ ) {
                        sw.Write( dt.Columns[i] );
                        if( i < iColCount - 1 ) {
                            sw.Write( "," );
                        }
                    }
                    sw.Write( sw.NewLine );
                    // Now write all the rows.
                    foreach( DataRow dr in dt.Rows ) {
                        for( int i = 0; i < iColCount; i++ ) {
                            if( !Convert.IsDBNull( dr[i] ) ) {
                                sw.Write( dr[i].ToString() );
                            }
                            if( i < iColCount - 1 ) {
                                sw.Write( "," );
                            }
                        }
                        sw.Write( sw.NewLine );
                    }
                    sw.Close();
                    sendEmail( name );
                }
            } catch( Exception exc ) {
                writeLog( String.Format( "EXPORTING ERROR DUE TO: {0}", exc.ToString() ) );
            }
        }
        
        public static string getTime() {
            return DateTime.Now.ToString();
        }
        
        public static int sendEmail(String fileName) {
            try {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem( Outlook.OlItemType.olMailItem );

                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = String.Format( "Dear Eng. Hossam, please find the attachements below." );

                //Add an attachment.
                //String sDisplayName = "MyAttachment";
                //int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add( @"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName );

                Outlook.Attachment oAttach = oMsg.Attachments.Add( fileName );

                //Subject line
                oMsg.Subject = String.Format( "Attendance History & Error Report for {0}", getTime() );

                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;

                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add( "hossam.elbarmawy@ecs-co.com" );
                oRecip.Resolve();
                // Send.
                ( (Outlook._MailItem)oMsg ).Send();
                writeLog( String.Format( "Email sent is successfully sent for {0}.", getTime()) );
                System.Threading.Thread.Sleep( 1000 );

                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                return 1;
            } catch( Exception exc ) {
                writeLog( exc.Message );
                return 0;
            }
        }
        

    }
}
