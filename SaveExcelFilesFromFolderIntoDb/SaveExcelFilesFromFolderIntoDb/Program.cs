using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
//added below name spaces
using System.IO;

namespace TechBrothersIT.com_CSharp_Tutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            //the datetime and Log folder will be used for error log file in case error occured
            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            string LogFolder = @"D:\Source\";
            try
            {
                //Provide the Source Folder path where excel files are present
                String FolderPath = @"D:\Source\";
                //Provide the table name in which you want to load excel sheet's data
                String TableName = @"tblCustomer";
                //Provide the schema of table
                String SchemaName = @"dbo";


                //Create Connection to SQL Server Database 
                String connString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = connString;


                var directory = new DirectoryInfo(FolderPath);
                FileInfo[] files = directory.GetFiles();

                //Declare and initilize variables
                string fileFullPath = "";


                //Get one Book(Excel file at a time)
                foreach (FileInfo file in files)
                {
                    fileFullPath = FolderPath + "\\" + file.Name;

                    //Create Excel Connection
                    string ConStr;
                    string HDR;
                    HDR = "YES";
                    ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                    OleDbConnection cnn = new OleDbConnection(ConStr);

                    //Get Sheet Name
                    cnn.Open();
                    DataTable dtSheet = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetname;
                    sheetname = "";

                    //Loop through each sheet
                    foreach (DataRow drSheet in dtSheet.Rows)
                    {
                        if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                        {
                            sheetname = drSheet["TABLE_NAME"].ToString();

                            //Get data from Excel Sheet to DataTable
                            OleDbConnection Conn = new OleDbConnection(ConStr);
                            using (Conn)
                            {
                                OleDbCommand oconn = new OleDbCommand("select * from [" + sheetname + "]", Conn);
                                OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                                DataTable dt = new DataTable();
                                adp.Fill(dt);

                                //Load Data from DataTable to SQL Server Table.
                                using (SQLConnection)
                                {
                                    using (SqlBulkCopy BC = new SqlBulkCopy(SQLConnection))
                                    {
                                        BC.DestinationTableName = SchemaName + "." + TableName;
                                        foreach (var column in dt.Columns)
                                            BC.ColumnMappings.Add(column.ToString(), column.ToString());
                                        BC.WriteToServer(dt);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                // Create Log File for Errors
                using (StreamWriter sw = File.CreateText(LogFolder
                    + "\\" + "ErrorLog_" + datetime + ".log"))
                {
                    sw.WriteLine(exception.ToString());
                }

            }

            Console.WriteLine("Podaci spremljeni u bazu!");
        }
    }
}
