using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.Net.Mail;
using DevExpress.XtraReports.UI;
using PayRunIO.CSharp.SDK;
using PayRunIOClassLibrary;

namespace PayRunIOProcessReports
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void UpdateContactDetails(XDocument xdoc)
        {
            string contactsFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Contacts\\";
            string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
            string dataBase = xdoc.Root.Element("Database").Value;
            string userID = xdoc.Root.Element("Username").Value;
            string password = xdoc.Root.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";

            DirectoryInfo dirInfo = new DirectoryInfo(contactsFolder);
            FileInfo[] files = dirInfo.GetFiles("*.csv");
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("_contacts_"))
                {
                    //Get a table of contacts from the csv file.
                    DataTable dtContacts = GetDataTableFromCSVFile(xdoc, file.FullName);
                    //Insert the data into an SQL Database.
                    bool success = InsertDataIntoSQLServerUsingSQLBulkCopy(dtContacts, sqlConnectionString, file.FullName, xdoc);
                    if (success)
                    {
                        //We've successfully written the contact data to a temporary table with the name "tmp_CompanyNo_Contacts". e.g. "tmp_2137_Contacts"
                        //Now Insert / Update the contacts table then delete the table.
                        int x = file.FullName.LastIndexOf("\\") + 1;
                        string companyNo = file.FullName.Substring(x, 4);
                        success = InsertUpdateContacts(xdoc, sqlConnectionString, companyNo);
                        if (success)
                        {
                            //Delete the temporary contacts.
                            DeleteTemporaryContacts(xdoc, sqlConnectionString);
                            //Delete the csv file.
                            file.Delete();
                        }
                    }
                }

            }
        }
        private bool InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvDataTable, string sqlConnectionString, string csvFileName, XDocument xdoc)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;

            using (SqlConnection sqlConnection = new SqlConnection(sqlConnectionString))
            {

                try
                {
                    sqlConnection.Open();
                    // Check if a table exsists
                    bool tableExists;
                    //
                    // Change the csvFileName to SQL table name here JCB TO DO
                    //
                    string tableName;
                    //This is the contacts file we've received from Web Globe it's named in the following format.
                    //CompanyNo_unity_contacts_export_datetimestamp.csv e.g. 1234_unity_contacts_export_20190630100130001.csv
                    //We just need the company number and contacts for the a table name.

                    tableName = "tmpContacts";  // Create a temporary invoices table and an SQL query will create the live one.


                    string sqlStatement = "SELECT COUNT (*) FROM " + tableName;


                    try
                    {
                        using (SqlCommand sqlCommand = new SqlCommand(sqlStatement, sqlConnection))
                        {
                            sqlCommand.ExecuteScalar();
                            tableExists = true;
                        }
                    }
                    catch
                    {
                        tableExists = false;
                    }

                    if (!tableExists)
                    {
                        // Create the table
                        try
                        {
                            textLine = string.Format("About to create tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            sqlStatement = "CREATE TABLE " + tableName + "(";
                            foreach (DataColumn dataColumn in csvDataTable.Columns)
                            {

                                dataColumn.ColumnName = Regex.Replace(dataColumn.ColumnName, "[^A-Za-z0-9]", "");
                                sqlStatement = sqlStatement + dataColumn.ColumnName + " varchar(150),";
                            }
                            sqlStatement = sqlStatement.Remove(sqlStatement.Length - 1, 1) + ")";
                            SqlCommand createTable = new SqlCommand(sqlStatement, sqlConnection);
                            createTable.ExecuteNonQuery();

                            textLine = string.Format("Sucessfully created tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);
                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Failed to create tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            return false;

                        }
                    }
                    try
                    {
                        using (SqlBulkCopy bulkData = new SqlBulkCopy(sqlConnection))
                        {
                            textLine = string.Format("About to bulk write to tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            bulkData.DestinationTableName = tableName;

                            foreach (DataColumn dataColumn in csvDataTable.Columns)
                            {
                                dataColumn.ColumnName = Regex.Replace(dataColumn.ColumnName, "[^A-Za-z0-9]", "");
                                bulkData.ColumnMappings.Add(dataColumn.ToString(), dataColumn.ToString());

                            }
                            //bulkData.BulkCopyTimeout = 600; // 600 seconds
                            bulkData.WriteToServer(csvDataTable);

                            textLine = string.Format("Successfull bulk write to tmpContacts table.");
                            update_Progress(textLine, configDirName, 1);

                            return true;

                        }
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Failed bulk write to tmpContacts table.");
                        update_Progress(textLine, configDirName, 1);

                        return false;

                    }
                }
                catch
                {
                    return false;

                }
                finally
                {
                    sqlConnection.Close();

                }

            }
        }
        private DataTable GetDataTableFromCSVFile(XDocument xdoc, string csvFileName)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            string delimiter = ",";
            DataTable csvDataTable = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csvFileName))
                {
                    csvReader.SetDelimiters(new string[] { delimiter });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    string[] colFields = csvReader.ReadFields();

                    foreach (string column in colFields)
                    {

                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.AllowDBNull = true;
                        //
                        // Check to make sure we don't have two columns with the same name.
                        //
                        try
                        {
                            csvDataTable.Columns.Add(datacolumn);
                        }
                        catch (Exception ex)
                        {
                            //
                            // We do have a column with this name already.
                            //
                            if (ex.ToString().Contains("already belongs to"))
                            {
                                DateTime dateTimeNow = DateTime.Now;
                                DataColumn dataColumnUnique = new DataColumn(column + dateTimeNow);
                                csvDataTable.Columns.Add(dataColumnUnique);
                            }
                            else
                            {
                                textLine = string.Format("Error getting data from csv file.\r\n{0}.\r\n", ex);
                                update_Progress(textLine, configDirName, logOneIn);
                            }

                        }

                    }

                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        int x = fieldData.Count();
                        string[] tableData = new string[x];
                        for (int i = 0; i < x; i++)
                        {
                            tableData[i] = fieldData[i];
                        }


                        csvDataTable.Rows.Add(tableData);
                    }
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting data from csv file.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);

            }
            return csvDataTable;
        }
        private bool InsertUpdateContacts(XDocument xdoc, string sqlConnectionString, string companyNo)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            bool success = false;
            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("InsertUpdateContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    command.ExecuteNonQuery();
                    success = true;
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error inserting/updating contacts.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }


            return success;
        }
        private void DeleteTemporaryContacts(XDocument xdoc, string sqlConnectionString)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("DeleteTemporaryContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.ExecuteNonQuery();

                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error deleting temporary contacts.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }


        }
        private void update_Progress(string textLine, string configDirName, int logOneIn)
        {
            //Get the month and year from today's date
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.Month.ToString().PadLeft(2, '0');
            string homeFolder = configDirName;

            using (StreamWriter sw = new StreamWriter(homeFolder + "Config\\" + "PRtoWG-Log" + year + month + ".txt", true))
            {
                textLine = string.Format(textLine + " - {0}", now);
                sw.WriteLine(textLine);

            }

        }
        private void ProcessReportsFromPayRunIO(XDocument xdoc)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string dataHomeFolder = xdoc.Root.Element("DataHomeFolder").Value;
            bool archive = Convert.ToBoolean(xdoc.Root.Element("Archive").Value);
            string sftpHostName = xdoc.Root.Element("SFTPHostName").Value;
            string user = xdoc.Root.Element("User").Value;
            string passwordFile = softwareHomeFolder + xdoc.Root.Element("PasswordFile").Value;
            string filePrefix = xdoc.Root.Element("FilePrefix").Value;
            int interval = Convert.ToInt32(xdoc.Root.Element("Interval").Value);
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);

            string textLine = null;

            textLine = string.Format("Start processing the reports.");
            update_Progress(textLine, softwareHomeFolder, logOneIn);

            //We'er going to change the way this done. Instead of PR producing folders with the EmployeePeriod & EmployeeYtd report already in them.
            //PR are going to give us an xml file to tell that the payroll has been done and the file will contain enough info to let us produce the required reports.
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            FileInfo[] completedPayrollFiles = prWG.GetAllCompletedPayrollFiles(xdoc);
            foreach (FileInfo completedPayrollFile in completedPayrollFiles)
            {
                ReadProcessCompletedPayrollFile(xdoc, completedPayrollFile);
                //Put in some test for success then archive the file.
                ArchiveCompletedPayrollFile(xdoc, completedPayrollFile);
            }

            ////This is the old method with folders containing the reports.
            //string[] directories = GetAListOfDirectories(xdoc);
            //for (int i = 0; i < directories.Count(); i++)
            //{
            //    try
            //    {
            //        bool success = ProduceReports(xdoc, directories[i]);
            //        if (success)
            //        {
            //            ArchiveDirectory(xdoc, directories[i]);
            //        }


            //    }
            //    catch (Exception ex)
            //    {
            //        textLine = string.Format("Error processing the reports for directory {0}.\r\n{1}.\r\n", directories[i], ex);
            //        update_Progress(textLine, softwareHomeFolder, logOneIn);
            //    }

            //}

            textLine = string.Format("Finished processing the reports.");
            update_Progress(textLine, softwareHomeFolder, logOneIn);
        }
        private void ArchiveCompletedPayrollFile(XDocument xdoc, FileInfo completedPayrollFile)
        {
            string destFileName = completedPayrollFile.FullName.Replace("Outputs", "PE-ArchivedOutputs");

            File.Move(completedPayrollFile.FullName, destFileName);
        }
        private void ReadProcessCompletedPayrollFile(XDocument xdoc, FileInfo completedPayrollFile)
        {
            XmlDocument xmlCompletedPayroll = new XmlDocument();
            xmlCompletedPayroll.Load(completedPayrollFile.FullName);

            //Now extract the necessary data and produce the required reports.
            
            RPParameters rpParameters = new RPParameters();
            foreach (XmlElement parameter in xmlCompletedPayroll.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = GetDateElementByTagFromXml(parameter, "AccountingYearStartDate");
                rpParameters.AccYearEnd = GetDateElementByTagFromXml(parameter, "AccountingYearEndDate");
                rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
            }
            GenerateReportsFromPR(xdoc, rpParameters);

        }
        private void GenerateReportsFromPR(XDocument xdoc, RPParameters rpParameters) 
        {
            //Produce and process Employee Period report.
            //Get the history report
            string rptRef = "EEPERIOD";              //Original report name : "PayescapeEmployeePeriod"
            string parameter1 = "EmployerKey";
            string parameter2 = "TaxYear";
            string parameter3 = "AccPeriodStart";
            string parameter4 = "AccPeriodEnd";
            string parameter5 = "TaxPeriod";
            string parameter6 = "PayScheduleKey";

            //Get the history report
            XmlDocument xmlPeriodReport = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
                                              rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
                                              parameter6, rpParameters.PaySchedule.ToUpper());
            //PrepareStandardReportsOld(xdoc, xmlPeriodReport);
            var tuple = PrepareStandardReports(xdoc, xmlPeriodReport);
            List<RPEmployeePeriod> rpEmployeePeriodList = tuple.Item1;
            List<RPPayComponent> rpPayComponents = tuple.Item2;
            //I don't think the P45 report will be able to be produced from the EmployeePeriod report but I'm leaving it here for now.
            List<P45> p45s = tuple.Item3;
            RPEmployer rpEmployer = tuple.Item4;
            //Put a sort of the pay codes in here
            //I didn't need to sort it in the end because the DevExpress report does it but this is useful code for future reference.
            //rpPayComponents.Sort(delegate (RPPayComponent x, RPPayComponent y)
            //{
            //    if (x.Description == null && y.Description == null) return 0;
            //    else if (x.Description == null) return -1;
            //    else if (y.Description == null) return 1;
            //    else return x.Description.CompareTo(y.Description);
            //});
            //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
            decimal hmrcTotal = CalculateHMRCTotal(rpEmployeePeriodList);
            string hmrcDesc = "[" + hmrcTotal.ToString() + "]";
            //I now have a list of employee with their total for this period & ytd plus addition & deductions
            //I can print payslips from here.
            PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents);

            //Change this to supply rpEmployeePeriodList instead of xmlPeriodReport
            CreateHistoryCSV(xdoc, rpParameters,rpEmployer,rpEmployeePeriodList);


            //Produce and process Employee Ytd report.
            rptRef = "EEYTD";              //Original report name : "PayescapeEmployeeYtd"
            XmlDocument xmlYTDReport = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
                                              rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
                                              parameter6, rpParameters.PaySchedule.ToUpper());
            CreateYTDCSV(xdoc, xmlYTDReport);

            //Produce and process P45s if required.
            rptRef = "P45";
            parameter2 = "EmployeeKey";
            rpParameters.ErRef = "1176";
            string eeRef = "14";
            XmlDocument xmlP45Report = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, eeRef, null,
                                              null, null, null, null, null, null, null);

            //Produce and process P32 if required
            rptRef = "P32S";
            parameter2 = "TaxYear";
            XmlDocument xmlP32Report = RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), null,
                                              null, null, null, null, null, null, null);

            ZipReports(xdoc, rpEmployer, rpParameters, hmrcDesc);
            EmailZippedReports(xdoc, rpEmployer, rpParameters);


        }
        private Tuple<int,int> TupleTest()
        {
            return new Tuple<int,int>(0,0);
        }
        private XmlDocument RunReport(string rptRef, string prm1, string val1, string prm2, string val2, string prm3, string val3,
                                 string prm4, string val4, string prm5, string val5, string prm6, string val6)
        {
            string url = null;
            if (prm2 == null)
            {
                url = prm1 + "=" + val1;

            }
            else if (prm3 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2;

            }
            else if (prm4 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3;

            }
            else if (prm5 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4;

            }
            else if (prm6 == null)
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                            + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4
                            + "&" + prm5 + "=" + val5;

            }
            else
            {
                url = prm1 + "=" + val1 + "&" + prm2 + "=" + val2
                           + "&" + prm3 + "=" + val3 + "&" + prm4 + "=" + val4
                           + "&" + prm5 + "=" + val5 + "&" + prm6 + "=" + val6;
            }
            XmlDocument xmlReport = null;
            try
            {
                //Mark this is the full url = "https://api.test.payrun.io/Report/PayescapeEmployeePeriod/run?EmployerKey=1104&TaxYear=2018&AccPeriodStart=2018/01/01&AccPeriodEnd=2019/03/08&TaxPeriod=49&PayScheduleKey=Weekly"
                var apiHelper = ApiHelper();
                //string testurl = "EmployerKey=1958&TaxYear=2019&AccPeriodStart=2019-04-06&AccPeriodEnd=2020-04-05&TaxPeriod=27&PayScheduleKey=Weekly";
                //xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + testurl);
                xmlReport = apiHelper.GetRawXml("/Report/" + rptRef + "/run?" + url);

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error running a report.\r\n" + ex);
            }
            return xmlReport;
        }
        private RestApiHelper ApiHelper()
        {
            string consumerKey = "1UH6t3ikiWbdxTNT2Dg";                             //Original developer key : "m5lsJMpBnkaJw086zwDw"
            string consumerSecret = "jKUX3lrQUe4KhEiox6IZw8CXnWUdAkyTl1kthR8ayQ";   //Original developer secret : "GHM6x3xLEWujpLC5sGXKQ3r2j14RGI0eoLbab8w415Q"
            string url = "https://api.test.payrun.io";
            RestApiHelper apiHelper = new PayRunIO.CSharp.SDK.RestApiHelper(
                    new PayRunIO.OAuth1.OAuthSignatureGenerator(),
                    consumerKey,
                    consumerSecret,
                    url,
                    "application/xml",
                    "application/xml");
            return apiHelper;
        }
        private void ArchiveDirectory(XDocument xdoc, string directory)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                DateTime now = DateTime.Now;

                int x = directory.LastIndexOf("\\");
                string coNo = directory.Substring(x + 1, 4);
                Directory.CreateDirectory(directory.Replace("Outputs", "PE-ArchivedOutputs"));
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    string destFileName = file.FullName.Replace("Outputs", "PE-ArchivedOutputs");
                    destFileName = destFileName.Replace(".xml", "_" + now.ToString("yyyyMMddHHmmssfff") + ".xml");
                    File.Move(file.FullName, destFileName);

                }

                Directory.Delete(directory);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error archiving the Outputs directory, {0}.\r\n{1}.\r\n", directory, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

        }
        private string[] GetAListOfDirectories(XDocument xdoc)
        {
            string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
            string[] directories = Directory.GetDirectories(path);

            return directories;
        }
        //private FileInfo[] GetAllCompletedPayrollFiles(XDocument xdoc)
        //{
        //    string path = xdoc.Root.Element("DataHomeFolder").Value + "Outputs";
        //    DirectoryInfo folder = new DirectoryInfo(path);
        //    FileInfo[] files = folder.GetFiles("*CompletedPayroll*.xml");

        //    return files;
        //}
        private bool ProduceReports(XDocument xdoc, string directory)
        {
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            bool eePeriodProcessed = false;
            bool eeYtdProcessed = false;
            DirectoryInfo dirInfo = new DirectoryInfo(directory);
            FileInfo[] files = dirInfo.GetFiles("*.xml");
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("EmployeePeriod"))
                {
                    try
                    {
                        ProducePeriodReport(xdoc, file);
                        ProducePDFReports(xdoc, file);
                        eePeriodProcessed = true;
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error producing the employee period reports for file {0}.\r\n{1}.\r\n", file, ex);
                        update_Progress(textLine, configDirName, logOneIn);
                    }

                }
                else if (file.FullName.Contains("EmployeeYtd"))
                {
                    try
                    {
                        ProduceYTDReport(xdoc, file);
                        eeYtdProcessed = true;
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error producing the employee ytd report for file {0}.\r\n{1}.\r\n", file, ex);
                        update_Progress(textLine, configDirName, logOneIn);
                    }
                }

            }
            if (eePeriodProcessed && eeYtdProcessed)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void ProducePeriodReport(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlPeriodReport = new XmlDocument();
            xmlPeriodReport.Load(file.FullName);
            CreateHistoryCSVOld(xdoc, xmlPeriodReport);
        }
        private void ProduceYTDReport(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlPeriodReport = new XmlDocument();
            xmlPeriodReport.Load(file.FullName);
            CreateYTDCSV(xdoc, xmlPeriodReport);
        }
        private void CreateYTDCSV(XDocument xdoc, XmlDocument xmlReport)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            RPParameters rpParameters = new RPParameters();
            foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = GetDateElementByTagFromXml(parameter, "AccountingYearStartDate");
                rpParameters.AccYearEnd = GetDateElementByTagFromXml(parameter, "AccountingYearEndDate");
                rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
            }

            string coNo = rpParameters.ErRef;
            //Write the whole xml file to the folder.
            //string xmlFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            string dirName = outgoingFolder + "\\" + coNo + "\\";
            Directory.CreateDirectory(dirName);
            string xmlFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            StreamWriter sw = new StreamWriter(xmlFileName);
            string xmlStream = xmlReport.InnerXml;
            sw.WriteLine(xmlStream);
            sw.Close();
            //Create csv version and write it to the same folder.
            //string csvFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            string csvFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            bool writeHeader = true;
            using (sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payYTDDetails = new string[41];


                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;
                    bool payRunDate = false;
                    if (GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                    {
                        if (!payRunDate)
                        {
                            rpParameters.PayRunDate = GetDateElementByTagFromXml(employee, "PayRunDate");
                            payRunDate = true;
                        }

                        //If the employee is a leaver before the start date then don't include.
                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        DateTime leavingDate = new DateTime();
                        if (GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            leavingDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        }
                        DateTime periodStartDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "ThisPeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        if (leaver.StartsWith("N"))
                        {
                            include = true;
                        }
                        else if (leavingDate >= periodStartDate)
                        {
                            include = true;
                        }
                    }

                    if (include)
                    {
                        DateTime dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "LastPaymentDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        payYTDDetails[0] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payYTDDetails[1] = GetElementByTagFromXml(employee, "EeRef");     //EeRef without the EE

                        if (GetElementByTagFromXml(employee, "LeavingDate") != null && GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            payYTDDetails[2] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payYTDDetails[2] = "";

                        }
                        payYTDDetails[3] = GetElementByTagFromXml(employee, "Leaver").Substring(0, 1);        //N of No or Y of Yes
                        payYTDDetails[4] = GetElementByTagFromXml(employee, "TaxPrevEmployment");
                        payYTDDetails[5] = GetElementByTagFromXml(employee, "TaxablePayPrevEmployment");
                        payYTDDetails[6] = GetElementByTagFromXml(employee, "TaxThisEmployment");
                        payYTDDetails[7] = GetElementByTagFromXml(employee, "TaxablePayThisEmployment");
                        payYTDDetails[8] = GetElementByTagFromXml(employee, "GrossedUp");
                        payYTDDetails[9] = GetElementByTagFromXml(employee, "GrossedUpTax");
                        payYTDDetails[10] = GetElementByTagFromXml(employee, "NetPayYTD");
                        payYTDDetails[11] = GetElementByTagFromXml(employee, "GrossPayYTD");
                        payYTDDetails[12] = GetElementByTagFromXml(employee, "BenefitInKindYTD");
                        payYTDDetails[13] = GetElementByTagFromXml(employee, "SuperannuationYTD");
                        payYTDDetails[14] = GetElementByTagFromXml(employee, "HolidayPayYTD");
                        payYTDDetails[15] = GetElementByTagFromXml(employee, "ErPensionYTD");
                        payYTDDetails[16] = GetElementByTagFromXml(employee, "EePensionYTD");
                        payYTDDetails[17] = GetElementByTagFromXml(employee, "AeoYTD");
                        if (GetElementByTagFromXml(employee, "StudentLoanStartDate") != null && GetElementByTagFromXml(employee, "StudentLoanStartDate") != "")
                        {
                            dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "StudentLoanStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            payYTDDetails[18] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payYTDDetails[18] = "";

                        }
                        if (GetElementByTagFromXml(employee, "StudentLoanEndDate") != null && GetElementByTagFromXml(employee, "StudentLoanEndDate") != "")
                        {
                            dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "StudentLoanEndDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            payYTDDetails[19] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payYTDDetails[19] = "";

                        }
                        payYTDDetails[20] = GetElementByTagFromXml(employee, "StudentLoanDeductionsYTD");
                        payYTDDetails[21] = GetElementByTagFromXml(employee, "NiLetter");
                        payYTDDetails[22] = GetElementByTagFromXml(employee, "NiableYtd");
                        payYTDDetails[23] = GetElementByTagFromXml(employee, "EarningToLEL");
                        payYTDDetails[24] = GetElementByTagFromXml(employee, "EarningsToSET");
                        payYTDDetails[25] = GetElementByTagFromXml(employee, "EarningsToPET");
                        payYTDDetails[26] = GetElementByTagFromXml(employee, "EarningsToUST");
                        payYTDDetails[27] = GetElementByTagFromXml(employee, "EarningsToAUST");
                        payYTDDetails[28] = GetElementByTagFromXml(employee, "EarningsToUEL");
                        payYTDDetails[29] = GetElementByTagFromXml(employee, "EarningsAboveUEL");
                        payYTDDetails[30] = GetElementByTagFromXml(employee, "EeContributionsPt1");
                        payYTDDetails[31] = GetElementByTagFromXml(employee, "EeContributionsPt2");
                        payYTDDetails[32] = GetElementByTagFromXml(employee, "ErContributions");
                        payYTDDetails[33] = GetElementByTagFromXml(employee, "EeRebate");
                        payYTDDetails[34] = GetElementByTagFromXml(employee, "ErRebate");
                        payYTDDetails[35] = GetElementByTagFromXml(employee, "EeReduction");
                        payYTDDetails[36] = GetElementByTagFromXml(employee, "TaxCode");
                        if (GetElementByTagFromXml(employee, "Week1Month1") == "False")
                        {
                            payYTDDetails[37] = "N";
                        }
                        else
                        {
                            payYTDDetails[37] = "Y";
                        }
                        payYTDDetails[38] = GetElementByTagFromXml(employee, "WeekNumber");
                        payYTDDetails[39] = GetElementByTagFromXml(employee, "MonthNumber");
                        payYTDDetails[40] = GetElementByTagFromXml(employee, "PeriodNumber");
                        //These next few fields get treated like pay codes. Use them if they are not zero.
                        //4 pay components EeNiPaidByEr, EeGuTaxPaidByEr, EeNiLERtoUER & ErNi
                        for (int i = 0; i < 6; i++)
                        {
                            string[] payCodeDetails = new string[8];
                            switch (i)
                            {
                                case 0:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "E";
                                    payCodeDetails[2] = "EeNIPdByEr";
                                    payCodeDetails[3] = "Ee NI Paid By Er";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "EeNiPaidByErAccountsAmount");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "EeNiPaidByErPayeAmount");
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "EeNiPaidByErAccountsUnits");
                                    payCodeDetails[7] = GetElementByTagFromXml(employee, "EeNiPaidByErPayeUnits");
                                    break;
                                case 1:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "E";
                                    payCodeDetails[2] = "GUTax";
                                    payCodeDetails[3] = "Grossed up Tax";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                                    payCodeDetails[7] = GetElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                                    break;
                                case 2:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "T";
                                    payCodeDetails[2] = "NIEeeLERtoUER";
                                    payCodeDetails[3] = "NIEeeLERtoUER-A";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                                    payCodeDetails[7] = GetElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                                    break;
                                case 3:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "T";
                                    payCodeDetails[2] = "NIEr";
                                    payCodeDetails[3] = "NIEr-A";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "ErNiAccountAmount");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "ErNiPayeAmount");
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "ErNiAccountUnit");
                                    payCodeDetails[7] = GetElementByTagFromXml(employee, "ErNiPayeUnit");
                                    break;
                                case 4:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "D";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "PenEr";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "ErPensionYTD");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "ErPensionYTD");
                                    payCodeDetails[6] = "0.00";
                                    payCodeDetails[7] = "0.00";
                                    break;
                                default:
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "D";
                                    payCodeDetails[2] = "PenPostTaxEe";
                                    payCodeDetails[3] = "PenPostTaxEe";
                                    payCodeDetails[4] = GetElementByTagFromXml(employee, "EePensionYTD");
                                    payCodeDetails[5] = GetElementByTagFromXml(employee, "EePensionYTD");
                                    payCodeDetails[6] = "0.00";
                                    payCodeDetails[7] = "0.00";
                                    break;
                            }

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (payCodeDetails[4] == "0.00" && payCodeDetails[5] == "0.00" &&
                                payCodeDetails[6] == "0.00" && payCodeDetails[7] == "0.00")
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayYTDCSV(rpParameters, payYTDDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }
                        }

                        foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                        {
                            foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                            {
                                string[] payCodeDetails = new string[8];
                                payCodeDetails[0] = GetElementByTagFromXml(payCode, "Code");
                                payCodeDetails[1] = GetElementByTagFromXml(payCode, "EarningOrDeduction");
                                payCodeDetails[2] = GetElementByTagFromXml(payCode, "Code");
                                payCodeDetails[3] = GetElementByTagFromXml(payCode, "Description");
                                payCodeDetails[4] = GetElementByTagFromXml(payCode, "AccountsAmount");
                                payCodeDetails[5] = GetElementByTagFromXml(payCode, "PayeAmount");
                                payCodeDetails[6] = GetElementByTagFromXml(payCode, "AccountsUnits");
                                payCodeDetails[7] = GetElementByTagFromXml(payCode, "PayeUnits");

                                //
                                //Check if any of the values are not zero. If so write the first employee record
                                //
                                bool allZeros = false;
                                if (payCodeDetails[4] == "0.00" && payCodeDetails[5] == "0.00" &&
                                    payCodeDetails[6] == "0.00" && payCodeDetails[7] == "0.00")
                                {
                                    allZeros = true;

                                }
                                if (!allZeros)
                                {
                                    //I don't require TAX, NI or PENSION
                                    if (payCodeDetails[0] != "TAX" && payCodeDetails[0] != "NI" && !payCodeDetails[0].StartsWith("PENSION"))
                                    {
                                        if (payCodeDetails[1] == "D")
                                        {
                                            //Deduction so multiply by -1
                                            for (int i = 4; i < 8; i++)
                                            {
                                                payCodeDetails[i] = (Convert.ToDecimal(payCodeDetails[i]) * -1).ToString();
                                            }
                                        }
                                        if (payCodeDetails[0] == "UNPDM")
                                        {
                                            //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                            payCodeDetails[0] = "";// "UNPD£";
                                            payCodeDetails[2] = "UNPD£";
                                        }
                                        else
                                        {
                                            payCodeDetails[0] = "";
                                        }
                                        //Write employee record
                                        WritePayYTDCSV(rpParameters, payYTDDetails, payCodeDetails, sw, writeHeader);
                                        writeHeader = false;
                                    }



                                }

                            }
                        }
                    }


                }

            }

        }
        private void WritePayYTDCSV(RPParameters rpParameters, string[] payYTDDetails, string[] payCodeDetails, StreamWriter sw, bool writeHeader)
        {
            string csvLine = null;
            if (writeHeader)
            {
                string csvHeader = "Co,RunDate,process,Batch,EeRef,LeaveDate,Leaver,Tax Previous Emt," +
                              "Taxable Pay Previous Emt, Tax This Emt, Taxable Pay This Emt,Grossed Up," +
                              "Grossed Up Tax, Net Pay,GrossYTD,Benefit in Kind,Superannuation," +
                              "Holiday Pay, ErPensionYTD, EePensionYTD, AEOYTD, StudentLoanStartDate," +
                              "StudentLoanEndDate, StudentLoanDeductions, NI Letter,Total," +
                              "Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL," +
                              "Ee Contributions Pt1,Ee Contributions Pt2,Er Contributions," +
                              "Ee Rebate,Er Rebate, Ee Reduction,PayCode,det,payCodeValue," +
                              "payCodeDesc,Acc Year Bal,PAYE Year Bal,Acc Year Units," +
                              "PAYE Year Units,Tax Code, Week1/ Month 1,Week Number, Month Number";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch = null;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "TwoWeekly":
                    batch = "M";
                    break;
                case "FourWeekly":
                    batch = "M";
                    break;
                case "Yearly":
                    batch = "M";
                    break;
                default:
                    batch = "W";
                    break;
            }
            if (rpParameters.PaySchedule == "Monthly")
            {
                batch = "M";
            }

            string process = null;
            process = "20" + payYTDDetails[0].Substring(6, 2) + payYTDDetails[0].Substring(3, 2) + payYTDDetails[0].Substring(0, 2) + "01";
            csvLine = csvLine + "\"" + rpParameters.ErRef + "\"" + "," +                                     //Co. Number
                            "\"" + payYTDDetails[0] + "\"" + "," +                                            //Run Date / Last Payment Date
                            "\"" + process + "\"" + "," +                                                     //Process
                            "\"" + batch + "\"" + ",";                                                        //Batch


            //From payYTDDetails[1] (EeRef) to payYTDDetails[35] (EeReduction)
            for (int i = 1; i < 36; i++)
            {
                csvLine = csvLine + "\"" + payYTDDetails[i] + "\"" + ",";
            }
            //From payCodeDetails[0] (PayCode) to payCodeDetails[7] (PAYE Year Units)
            for (int i = 0; i < 8; i++)
            {
                csvLine = csvLine + "\"" + payCodeDetails[i] + "\"" + ",";
            }
            //From payYTDDetails[36] (TaxCode) to payYTDDetails[39] (Month Number)
            for (int i = 36; i < 40; i++)
            {
                csvLine = csvLine + "\"" + payYTDDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }
        private void ProducePDFReports(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlPeriodReport = new XmlDocument();
            xmlPeriodReport.Load(file.FullName);
            PrepareStandardReports(xdoc, xmlPeriodReport);
        }
        private void CreateHistoryCSVOld(XDocument xdoc, XmlDocument xmlReport)
        {
            RPEmployeePeriod rpEmployeePeriod = new RPEmployeePeriod();
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            RPParameters rpParameters = new RPParameters();
            foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = GetDateElementByTagFromXml(parameter, "AccountingYearStartDate");
                rpParameters.AccYearEnd = GetDateElementByTagFromXml(parameter, "AccountingYearEndDate");
                rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
            }
            string coNo = rpParameters.ErRef;
            //Write the whole xml file to the folder.
            //string xmlFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            string dirName = outgoingFolder + "\\" + coNo + "\\";
            Directory.CreateDirectory(dirName);
            string xmlFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            StreamWriter sw = new StreamWriter(xmlFileName);
            string xmlStream = xmlReport.InnerXml;
            sw.WriteLine(xmlStream);
            sw.Close();
            //Create csv version and write it to the same folder.
            string csvFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            bool writeHeader = true;
            using (sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payHistoryDetails = new string[49];

                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;
                    bool payRunDate = false;
                    if (GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                    {
                        if (!payRunDate)
                        {
                            rpParameters.PayRunDate = GetDateElementByTagFromXml(employee, "PayRunDate");
                            payRunDate = true;
                        }
                        //If the employee is a leaver before the start date then don't include.
                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        DateTime leavingDate = new DateTime();
                        if (GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            leavingDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        }
                        DateTime periodStartDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "PeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        if (leaver.StartsWith("N"))
                        {
                            include = true;
                        }
                        else if (leavingDate >= periodStartDate)
                        {
                            include = true;
                        }
                    }

                    if (include)
                    {
                        payHistoryDetails[0] = GetElementByTagFromXml(employee, "EeRef");
                        DateTime dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "PayRunDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        payHistoryDetails[1] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "PeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        payHistoryDetails[2] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "PeriodEndDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        payHistoryDetails[3] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[4] = GetElementByTagFromXml(employee, "PayrollYear");
                        payHistoryDetails[5] = GetElementByTagFromXml(employee, "Gross");
                        payHistoryDetails[6] = GetElementByTagFromXml(employee, "Net");
                        payHistoryDetails[7] = " "; // GetElementByTagFromXml(employee, "DayHours");
                        if (GetElementByTagFromXml(employee, "StudentLoanStartDate") != null && GetElementByTagFromXml(employee, "StudentLoanStartDate") != "")
                        {
                            dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "StudentLoanStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            payHistoryDetails[8] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[8] = "";
                        }
                        if (GetElementByTagFromXml(employee, "StudentLoanEndDate") != null && GetElementByTagFromXml(employee, "StudentLoanEndDate") != "")
                        {
                            dt = DateTime.ParseExact(GetElementByTagFromXml(employee, "StudentLoanEndDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            payHistoryDetails[9] = dt.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[9] = "";
                        }
                        payHistoryDetails[10] = GetElementByTagFromXml(employee, "StudentLoanDeductions");
                        payHistoryDetails[11] = GetElementByTagFromXml(employee, "NiLetter");
                        payHistoryDetails[12] = GetElementByTagFromXml(employee, "CalculationBasis");
                        payHistoryDetails[13] = GetElementByTagFromXml(employee, "Total");
                        payHistoryDetails[14] = GetElementByTagFromXml(employee, "EarningToLEL");
                        payHistoryDetails[15] = GetElementByTagFromXml(employee, "EarningsToSET");
                        payHistoryDetails[16] = GetElementByTagFromXml(employee, "EarningsToPET");
                        payHistoryDetails[17] = GetElementByTagFromXml(employee, "EarningsToUST");
                        payHistoryDetails[18] = GetElementByTagFromXml(employee, "EarningsToAUST");
                        payHistoryDetails[19] = GetElementByTagFromXml(employee, "EarningsToUEL");
                        payHistoryDetails[20] = GetElementByTagFromXml(employee, "EarningsAboveUEL");
                        payHistoryDetails[21] = GetElementByTagFromXml(employee, "EeContributionsPt1");
                        payHistoryDetails[22] = GetElementByTagFromXml(employee, "EeContributionsPt2");
                        payHistoryDetails[23] = GetElementByTagFromXml(employee, "ErContributions");
                        payHistoryDetails[24] = GetElementByTagFromXml(employee, "EeRebate");
                        payHistoryDetails[25] = GetElementByTagFromXml(employee, "ErRebate");
                        payHistoryDetails[26] = GetElementByTagFromXml(employee, "EeReduction");
                        payHistoryDetails[27] = GetElementByTagFromXml(employee, "LeavingDate");
                        payHistoryDetails[28] = GetElementByTagFromXml(employee, "Leaver");
                        payHistoryDetails[29] = GetElementByTagFromXml(employee, "TaxCode");
                        payHistoryDetails[30] = GetElementByTagFromXml(employee, "Week1Month1");
                        if (payHistoryDetails[30] == "False")
                        {
                            payHistoryDetails[30] = "N";
                        }
                        else
                        {
                            payHistoryDetails[30] = "Y";
                        }
                        payHistoryDetails[31] = "0";   // GetElementByTagFromXml(employee, "TaxCodeChangeTypeId");
                        payHistoryDetails[32] = "UNKNOWN"; // GetElementByTagFromXml(employee, "TaxCodeChangeType");
                        payHistoryDetails[33] = GetElementByTagFromXml(employee, "TaxPrevious");
                        payHistoryDetails[34] = GetElementByTagFromXml(employee, "TaxablePayPrevious");
                        payHistoryDetails[35] = GetElementByTagFromXml(employee, "TaxThis");
                        payHistoryDetails[36] = GetElementByTagFromXml(employee, "TaxablePayThisYTD");
                        payHistoryDetails[37] = GetElementByTagFromXml(employee, "HolidayAccruedTd");
                        payHistoryDetails[38] = GetElementByTagFromXml(employee, "ErPensionYTD");
                        payHistoryDetails[39] = GetElementByTagFromXml(employee, "EePensionYTD");
                        payHistoryDetails[40] = GetElementByTagFromXml(employee, "ErPensionTaxPeriod");
                        payHistoryDetails[41] = GetElementByTagFromXml(employee, "EePensionTaxPeriod");
                        payHistoryDetails[42] = GetElementByTagFromXml(employee, "ErPensionPayRunDate");
                        payHistoryDetails[43] = GetElementByTagFromXml(employee, "EePensionPayRunDate");
                        payHistoryDetails[44] = GetElementByTagFromXml(employee, "DirectorshipAppointmentDate");
                        payHistoryDetails[45] = GetElementByTagFromXml(employee, "Director");
                        if (payHistoryDetails[45] == "N")               //Director
                        {
                            //They're not a director
                            payHistoryDetails[12] = "E";                //They're not a director
                        }
                        else
                        {
                            //They're a director
                            if (payHistoryDetails[12] == "Cumulative")  //Calculation basis
                            {
                                payHistoryDetails[12] = "C";            //Calculation Basis is Cumulative and they're a director
                            }
                            else
                            {
                                payHistoryDetails[12] = "N";            //Calculation Basis is Week1Month1 and they're a director
                            }

                        }
                        payHistoryDetails[46] = GetElementByTagFromXml(employee, "EeContributionsTaxPeriodPt1");
                        payHistoryDetails[47] = GetElementByTagFromXml(employee, "EeContributionsTaxPeriodPt2");
                        payHistoryDetails[48] = GetElementByTagFromXml(employee, "ErContributionsTaxPeriod");

                        //Er NI & Er Pension
                        for (int i = 0; i < 2; i++)
                        {
                            string[] payCodeDetails = new string[12];

                            switch (i)
                            {
                                case 0:
                                    payCodeDetails[1] = "NIEr-A";
                                    payCodeDetails[2] = "NIEr";
                                    payCodeDetails[3] = "T";
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "ErContributionsTaxPeriod");
                                    break;
                                default:
                                    payCodeDetails[1] = "PenEr";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "M";
                                    payCodeDetails[6] = GetElementByTagFromXml(employee, "ErPensionTaxPeriod");
                                    break;
                            }
                            payCodeDetails[0] = "0";
                            payCodeDetails[4] = "0";
                            payCodeDetails[5] = "0.00";
                            payCodeDetails[7] = "0.00";
                            payCodeDetails[8] = "0.00";
                            payCodeDetails[9] = "0.00";
                            payCodeDetails[10] = "0";
                            payCodeDetails[11] = "0";

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (payCodeDetails[5] == "0.00" && payCodeDetails[6] == "0.00" &&
                                payCodeDetails[7] == "0.00" && payCodeDetails[8] == "0.00" &&
                                payCodeDetails[9] == "0.00")
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }

                        }


                        foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                        {
                            foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                            {
                                string[] payCodeDetails = new string[12];
                                payCodeDetails = new string[12];
                                payCodeDetails[0] = GetElementByTagFromXml(payCode, "Code");
                                payCodeDetails[1] = GetElementByTagFromXml(payCode, "Description");
                                payCodeDetails[2] = GetElementByTagFromXml(payCode, "Code");
                                payCodeDetails[3] = GetElementByTagFromXml(payCode, "EarningOrDeduction");
                                payCodeDetails[4] = GetElementByTagFromXml(payCode, "Rate");
                                payCodeDetails[5] = GetElementByTagFromXml(payCode, "Units");
                                payCodeDetails[6] = GetElementByTagFromXml(payCode, "Amount");
                                if (payCodeDetails[4] == "0.00")
                                {
                                    payCodeDetails[4] = payCodeDetails[6];  // Make Rate equal to amount if rate is zero.
                                }
                                payCodeDetails[7] = GetElementByTagFromXml(payCode, "AccountsYearBalance");
                                payCodeDetails[8] = GetElementByTagFromXml(payCode, "PayeYearBalance");
                                payCodeDetails[9] = GetElementByTagFromXml(payCode, "AccountsYearUnits");
                                payCodeDetails[10] = GetElementByTagFromXml(payCode, "PayeYearUnits");
                                payCodeDetails[11] = GetElementByTagFromXml(payCode, "PayrollAccrued");
                                switch (payCodeDetails[0]) //PayCode
                                {
                                    case "TAX":
                                        payCodeDetails[0] = "0";
                                        payCodeDetails[1] = payHistoryDetails[29];  // Tax Code
                                        payCodeDetails[2] = payHistoryDetails[29];  // Tax Code
                                        payCodeDetails[3] = "T";                    // Tax    
                                        break;
                                    case "NI":
                                        payCodeDetails[0] = "0";
                                        payCodeDetails[1] = "NIEeeLERtoUER-A";      // Ee NI
                                        payCodeDetails[2] = "NIEeeLERtoUER";        // Ee NI
                                        payCodeDetails[3] = "T";                    // Tax    
                                        break;
                                    case "PENSION":
                                    case "PENSIONSS":
                                    case "PENSIONRAS":
                                        payCodeDetails[0] = "0";
                                        payCodeDetails[1] = "PenPostTaxEe";         // Ee Pension
                                        payCodeDetails[2] = "PenPostTaxEe";         // Ee Pension
                                        break;
                                    default:
                                        payCodeDetails[0] = "";
                                        break;

                                }
                                switch (payCodeDetails[0]) //PayCode
                                {
                                    case "0":
                                        //Multiple by minus 1 to make them positive
                                        payCodeDetails[4] = "0.00";     // (Convert.ToDecimal(payCodeDetails[4]) * -1).ToString();         //Rate
                                        payCodeDetails[6] = (Convert.ToDecimal(payCodeDetails[6]) * -1).ToString();         //Amount
                                        payCodeDetails[7] = "0.00";     // (Convert.ToDecimal(payCodeDetails[7]) * -1).ToString();         //AccountsYearBalance
                                        payCodeDetails[8] = "0.00";     //(Convert.ToDecimal(payCodeDetails[8]) * -1).ToString();         //PayeYearBalance
                                        break;
                                }
                                switch (payCodeDetails[3])      //Earning or Deduction
                                {
                                    case "D":
                                        if (!payCodeDetails[1].StartsWith("PenPostTaxEe"))  //Pensions codes were dealt with above
                                        {
                                            //Multiple by minus 1 to make them positive
                                            payCodeDetails[4] = (Convert.ToDecimal(payCodeDetails[4]) * -1).ToString();         //Rate
                                            payCodeDetails[6] = (Convert.ToDecimal(payCodeDetails[6]) * -1).ToString();         //Amount
                                            payCodeDetails[7] = (Convert.ToDecimal(payCodeDetails[7]) * -1).ToString();         //AccountsYearBalance
                                            payCodeDetails[8] = (Convert.ToDecimal(payCodeDetails[8]) * -1).ToString();         //PayeYearBalance

                                        }
                                        break;
                                    default:
                                        break;
                                }

                                //
                                //Check if any of the values are not zero. If so write the first employee record
                                //
                                bool allZeros = false;
                                if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                    Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                    Convert.ToDecimal(payCodeDetails[9]) == 0)
                                {
                                    allZeros = true;

                                }
                                if (!allZeros)
                                {
                                    //Write employee record
                                    WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                    writeHeader = false;

                                }

                            }
                        }
                    }


                }

            }

        }
        private void CreateHistoryCSV(XDocument xdoc, RPParameters rpParameters, RPEmployer rpEmployer, List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            string coNo = rpParameters.ErRef;
            //Write the whole xml file to the folder.
            //string xmlFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            string dirName = outgoingFolder + "\\" + coNo + "\\";
            Directory.CreateDirectory(dirName);
            //Create csv version and write it to the same folder.
            string csvFileName = outgoingFolder + "\\" + coNo + "\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            bool writeHeader = true;
            using (StreamWriter sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payHistoryDetails = new string[49];

                foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
                {
                    bool include = false;
                    
                    //If the employee is a leaver before the start date then don't include.
                    if (!rpEmployeePeriod.Leaver)
                    {
                        include = true;
                    }
                    else if (rpEmployeePeriod.LeavingDate >= rpEmployeePeriod.PeriodStartDate)
                    {
                        include = true;
                    }
                   
                    if (include)
                    {
                        payHistoryDetails[0] = rpEmployeePeriod.Reference;
                        payHistoryDetails[1] = rpEmployeePeriod.PayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[2] = rpEmployeePeriod.PeriodStartDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[3] = rpEmployeePeriod.PeriodEndDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[4] = rpEmployeePeriod.PayrollYear.ToString();
                        payHistoryDetails[5] = rpEmployeePeriod.Gross.ToString();
                        payHistoryDetails[6] = rpEmployeePeriod.NetPayTP.ToString();
                        payHistoryDetails[7] = rpEmployeePeriod.DayHours.ToString();
                        if (rpEmployeePeriod.StudentLoanStartDate != null)
                        {
                            payHistoryDetails[8] = rpEmployeePeriod.StudentLoanStartDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[8] = "";
                        }
                        if (rpEmployeePeriod.StudentLoanEndDate != null)
                        {
                            payHistoryDetails[9] = rpEmployeePeriod.StudentLoanStartDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[9] = "";
                        }
                        payHistoryDetails[10] = rpEmployeePeriod.StudentLoan.ToString();
                        payHistoryDetails[11] = rpEmployeePeriod.NILetter;
                        payHistoryDetails[12] = rpEmployeePeriod.CalculationBasis;
                        payHistoryDetails[13] = rpEmployeePeriod.TotalPayTP.ToString();     //Was "Total"
                        payHistoryDetails[14] = rpEmployeePeriod.EarningsToLEL.ToString();
                        payHistoryDetails[15] = rpEmployeePeriod.EarningsToSET.ToString();
                        payHistoryDetails[16] = rpEmployeePeriod.EarningsToPET.ToString();
                        payHistoryDetails[17] = rpEmployeePeriod.EarningsToUST.ToString(); ;
                        payHistoryDetails[18] = rpEmployeePeriod.EarningsToAUST.ToString();
                        payHistoryDetails[19] = rpEmployeePeriod.EarningsToUEL.ToString();
                        payHistoryDetails[20] = rpEmployeePeriod.EarningsAboveUEL.ToString();
                        payHistoryDetails[21] = rpEmployeePeriod.EeContributionsPt1.ToString();
                        payHistoryDetails[22] = rpEmployeePeriod.EeContributionsPt2.ToString();
                        payHistoryDetails[23] = rpEmployeePeriod.ErNICYTD .ToString();
                        payHistoryDetails[24] = rpEmployeePeriod.EeRebate.ToString();
                        payHistoryDetails[25] = rpEmployeePeriod.ErRebate.ToString();
                        payHistoryDetails[26] = rpEmployeePeriod.EeReduction.ToString();
                        payHistoryDetails[27] = rpEmployeePeriod.LeavingDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        if(rpEmployeePeriod.Leaver)
                        {
                            payHistoryDetails[28] = "N";
                        }
                        else
                        {
                            payHistoryDetails[28] = "Y";
                        }
                        
                        payHistoryDetails[29] = rpEmployeePeriod.TaxCode.ToString();
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            payHistoryDetails[30] = "Y";
                        }
                        else
                        {
                            payHistoryDetails[30] = "N";
                        }
                        payHistoryDetails[31] = "0";   //rpEmployeePeriod.TaxCodeChangeTypeID;
                        payHistoryDetails[32] = "UNKNOWN"; //rpEmployeePeriod.TaxCodeChangeType;
                        payHistoryDetails[33] = rpEmployeePeriod.TaxPrev.ToString();
                        payHistoryDetails[34] = rpEmployeePeriod.TaxablePayPrevious.ToString();
                        payHistoryDetails[35] = rpEmployeePeriod.TaxThis.ToString();
                        payHistoryDetails[36] = rpEmployeePeriod.TaxablePayYTD.ToString();
                        payHistoryDetails[37] = rpEmployeePeriod.HolidayAccruedYTD.ToString();
                        payHistoryDetails[38] = rpEmployeePeriod.ErPensionYTD.ToString();
                        payHistoryDetails[39] = rpEmployeePeriod.EePensionYTD.ToString();
                        payHistoryDetails[40] = rpEmployeePeriod.ErPensionTP.ToString();
                        payHistoryDetails[41] = rpEmployeePeriod.EePensionTP.ToString();
                        payHistoryDetails[42] = rpEmployeePeriod.ErPensionPayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[43] = rpEmployeePeriod.EePensionPayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        payHistoryDetails[44] = rpEmployeePeriod.DirectorshipAppointmentDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        if(rpEmployeePeriod.Director)
                        {
                            payHistoryDetails[45] = "Y";
                        }
                        else
                        {
                            payHistoryDetails[45] = "N";
                        }
                        if (payHistoryDetails[45] == "N")               //Director
                        {
                            //They're not a director
                            payHistoryDetails[12] = "E";                //They're not a director
                        }
                        else
                        {
                            //They're a director
                            if (payHistoryDetails[12] == "Cumulative")  //Calculation basis
                            {
                                payHistoryDetails[12] = "C";            //Calculation Basis is Cumulative and they're a director
                            }
                            else
                            {
                                payHistoryDetails[12] = "N";            //Calculation Basis is Week1Month1 and they're a director
                            }

                        }
                        payHistoryDetails[46] = rpEmployeePeriod.EeContributionTaxPeriodPt1.ToString();
                        payHistoryDetails[47] = rpEmployeePeriod.EeContributionTaxPeriodPt2.ToString();
                        payHistoryDetails[48] = rpEmployeePeriod.ErNICTP.ToString();

                        //Er NI & Er Pension
                        for (int i = 0; i < 2; i++)
                        {
                            string[] payCodeDetails = new string[12];

                            switch (i)
                            {
                                case 0:
                                    payCodeDetails[1] = "NIEr-A";
                                    payCodeDetails[2] = "NIEr";
                                    payCodeDetails[3] = "T";
                                    payCodeDetails[6] = rpEmployeePeriod.ErNICTP.ToString();
                                    break;
                                default:
                                    payCodeDetails[1] = "PenEr";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "M";
                                    payCodeDetails[6] = rpEmployeePeriod.ErPensionTP.ToString();
                                    break;
                            }
                            payCodeDetails[0] = "0";
                            payCodeDetails[4] = "0";
                            payCodeDetails[5] = "0.00";
                            payCodeDetails[7] = "0.00";
                            payCodeDetails[8] = "0.00";
                            payCodeDetails[9] = "0.00";
                            payCodeDetails[10] = "0";
                            payCodeDetails[11] = "0";

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (payCodeDetails[5] == "0.00" && payCodeDetails[6] == "0.00" &&
                                payCodeDetails[7] == "0.00" && payCodeDetails[8] == "0.00" &&
                                payCodeDetails[9] == "0.00")
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }

                        }
                        foreach(RPAddition rpAddition in rpEmployeePeriod.Additions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[0] = rpAddition.Code;
                            payCodeDetails[1] = rpAddition.Description;
                            payCodeDetails[2] = rpAddition.Code;
                            payCodeDetails[3] = "E"; //Earnings
                            payCodeDetails[4] = rpAddition.Rate.ToString();
                            payCodeDetails[5] = rpAddition.Units.ToString();
                            payCodeDetails[6] = rpAddition.AmountTP.ToString();
                            if (payCodeDetails[4] == "0.00")
                            {
                                payCodeDetails[4] = payCodeDetails[6];  // Make Rate equal to amount if rate is zero.
                            }
                            payCodeDetails[7] = rpAddition.AccountsYearBalance.ToString();
                            payCodeDetails[8] = rpAddition.AmountYTD.ToString();
                            payCodeDetails[9] = rpAddition.AccountsYearUnits.ToString();
                            payCodeDetails[10] = rpAddition.PayeYearUnits.ToString();
                            payCodeDetails[11] = rpAddition.PayrollAccrued.ToString();
                            
                            

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[9]) == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }

                        
                        }
                        foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[0] = rpDeduction.Code;
                            payCodeDetails[1] = rpDeduction.Description;
                            payCodeDetails[2] = rpDeduction.Code;
                            payCodeDetails[3] = "D"; //Earnings
                            payCodeDetails[4] = rpDeduction.Rate.ToString();
                            payCodeDetails[5] = rpDeduction.Units.ToString();
                            payCodeDetails[6] = rpDeduction.AmountTP.ToString();
                            if (payCodeDetails[4] == "0.00")
                            {
                                payCodeDetails[4] = payCodeDetails[6];  // Make Rate equal to amount if rate is zero.
                            }
                            payCodeDetails[7] = rpDeduction.AccountsYearBalance.ToString();
                            payCodeDetails[8] = rpDeduction.AmountYTD.ToString();
                            payCodeDetails[9] = rpDeduction.AccountsYearUnits.ToString();
                            payCodeDetails[10] = rpDeduction.PayeYearUnits.ToString();
                            payCodeDetails[11] = rpDeduction.PayrollAccrued.ToString();
                            switch (payCodeDetails[0]) //PayCode
                            {
                                case "TAX":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[2] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "NI":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "NIEeeLERtoUER-A";      // Ee NI
                                    payCodeDetails[2] = "NIEeeLERtoUER";        // Ee NI
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "PENSION":
                                case "PENSIONSS":
                                case "PENSIONRAS":
                                    payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPostTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPostTaxEe";         // Ee Pension
                                    break;
                                default:
                                    payCodeDetails[0] = "";
                                    break;

                            }
                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (Convert.ToDecimal(payCodeDetails[5]) == 0 && Convert.ToDecimal(payCodeDetails[6]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[7]) == 0 && Convert.ToDecimal(payCodeDetails[8]) == 0 &&
                                Convert.ToDecimal(payCodeDetails[9]) == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //Write employee record
                                WritePayHistoryCSV(rpParameters, payHistoryDetails, payCodeDetails, sw, writeHeader);
                                writeHeader = false;

                            }


                        }
                        
                    }


                }

            }

        }
        private void WritePayHistoryCSV(RPParameters rpParameters, string[] payHistoryDetails, string[] payCodeDetails, StreamWriter sw, bool writeHeader)
        {

            string csvLine = null;
            if (writeHeader)
            {
                string csvHeader = "co,runDate,Period_Start_Date,Period_End_Date,process,PayrollYear," +
                              "EEid,Gross,NetPay,Batch,CheckVoucher,Account,Transit,DeptName," +
                              "CostCentreName,branchName,Days/Hours,StudentLoanStartDate," +
                              "StudentLoanEndDate,StudentLoanDeductions,NI Letter,Calculation Basis," +
                              "Total,Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL,Ee Contributions Pt1," +
                              "Ee Contributions Pt2,Er Contributions,Ee Rebate,Er Rebate,Ee Reduction," +
                              "LeaveDate,Leaver,Tax Code,Week1/Month 1,Tax Code Change Type ID," +
                              "Tax Code Change Type,Tax Previous Emt,Taxable Pay Previous Emt,Tax This Emt," +
                              "Taxable Pay This Emt,PayCode,payCodeDesc,payCodeValue,det,rate,hours,Amount," +
                              "Acc Year Bal,PAYE Year Bal,Acc Year Units,PAYE Year Units,Payroll Accrued";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch = null;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "TwoWeekly":
                    batch = "M";
                    break;
                case "FourWeekly":
                    batch = "M";
                    break;
                case "Yearly":
                    batch = "M";
                    break;
                default:
                    batch = "W";
                    break;
            }
            if (rpParameters.PaySchedule == "Monthly")
            {
                batch = "M";
            }

            string process = null;
            process = "20" + payHistoryDetails[1].Substring(6, 2) + payHistoryDetails[1].Substring(3, 2) + payHistoryDetails[1].Substring(0, 2) + "01";
            csvLine = csvLine + "\"" + rpParameters.ErRef + "\"" + "," +                                   //Co. Number
                            "\"" + payHistoryDetails[1] + "\"" + "," +                                  //Run Date
                            "\"" + payHistoryDetails[2] + "\"" + "," +                                  //Period Start Date
                            "\"" + payHistoryDetails[3] + "\"" + "," +                                  //Period End Date
                            "\"" + process + "\"" + "," +                                               //Process
                            "\"" + payHistoryDetails[4] + "\"" + "," +                                  //Payroll Year
                            "\"" + payHistoryDetails[0] + "\"" + "," +                     //Ee ID
                            "\"" + payHistoryDetails[5] + "\"" + "," +                                  //Gross
                            "\"" + payHistoryDetails[6] + "\"" + "," +                                  //Net
                            "\"" + batch + "\"" + "," +                                                 //batch
                            "\"" + "0" + "\"" + "," +                                                   //CheckVoucher
                            "\"" + "0" + "\"" + "," +                                                   //Account
                            "\"" + "0" + "\"" + "," +                                                   //Transit
                            "\"" + "[Default]" + "\"" + "," +                                           //DeptName
                            "\"" + "[Default]" + "\"" + "," +                                           //CostCentreName
                            "\"" + "[Default]" + "\"" + ",";                                            //branchName

            //From payHistoryDetails[7] (DayHours) to payHistoryDetails[36] (Taxable Pay This)
            for (int i = 7; i < 37; i++)
            {
                csvLine = csvLine + "\"" + payHistoryDetails[i] + "\"" + ",";
            }
            //From payCodeDetails[0] (PayCode) to payCodeDetails[11] (Payroll Accrued)
            for (int i = 0; i < 12; i++)
            {
                csvLine = csvLine + "\"" + payCodeDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }
        private decimal GetDecimalElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            decimal decimalValue = 0;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != "" && element != " " && element != null)
            {
                decimalValue = Convert.ToDecimal(GetElementByTagFromXml(xmlElement, tag));
            }

            return decimalValue;
        }
        private bool GetBooleanElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            bool boolValue = false;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != "" && element != " " && element != null)
            {
                boolValue = Convert.ToBoolean(GetElementByTagFromXml(xmlElement, tag));
            }

            return boolValue;
        }
        private int GetIntElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            int intValue = 0;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != "" && element != " " && element != null)
            {
                intValue = Convert.ToInt32(GetElementByTagFromXml(xmlElement, tag));
            }

            return intValue;
        }
        private DateTime GetDateElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            DateTime dateValue = DateTime.MinValue;
            string element = GetElementByTagFromXml(xmlElement, tag);
            if (element != "" && element != " " && element != null)
            {
                dateValue = Convert.ToDateTime(GetElementByTagFromXml(xmlElement, tag));
            }

            return dateValue;
        }
        private string GetElementByTagFromXml(XmlElement xmlElement, string tag)
        {
            string element = null;
            int i = xmlElement.GetElementsByTagName(tag).Count;
            if (i > 0)
            {
                element = xmlElement.GetElementsByTagName(tag).Item(0).InnerText;
            }
            return element;
        }
        private void PrepareStandardReportsOld(XDocument xdoc, XmlDocument xmlReport)
        {
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;

            try
            {
                RPParameters rpParameters = new RPParameters();
                foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
                {
                    rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                    rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                    rpParameters.AccYearStart = GetDateElementByTagFromXml(parameter, "AccountingYearStartDate");
                    rpParameters.AccYearEnd = GetDateElementByTagFromXml(parameter, "AccountingYearEndDate");
                    rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                    rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
                }
                RPEmployer rpEmployer = new RPEmployer();

                foreach (XmlElement employer in xmlReport.GetElementsByTagName("Employer"))
                {
                    rpEmployer.Name = GetElementByTagFromXml(employer, "Name");
                    rpEmployer.PayeRef = GetElementByTagFromXml(employer, "EmployerPayeRef");
                }

                List<RPEmployeePeriod> rpEmployeePeriodList = new List<RPEmployeePeriod>();
                List<P45> p45s = new List<P45>();
                //Create a list of Pay Code totals for the Payroll Component Analysis report
                List<RPPayComponent> rpPayComponents = new List<RPPayComponent>();

                bool payRunDate = false;
                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;

                    if (GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                    {
                        if (!payRunDate)
                        {
                            rpParameters.PayRunDate = GetDateElementByTagFromXml(employee, "PayRunDate");
                            payRunDate = true;
                        }
                        //If the employee is a leaver before the start date then don't include.
                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        DateTime leavingDate = new DateTime();
                        if (GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            leavingDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        }
                        DateTime periodStartDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "PeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        if (leaver.StartsWith("N"))
                        {
                            include = true;
                        }
                        else if (leavingDate >= periodStartDate)
                        {
                            include = true;
                        }
                    }

                    if (include)
                    {
                        RPEmployeePeriod rpEmployeePeriod = new RPEmployeePeriod();
                        rpEmployeePeriod.Reference = GetElementByTagFromXml(employee, "EeRef");
                        rpEmployeePeriod.Title = GetElementByTagFromXml(employee, "Title");
                        rpEmployeePeriod.Forename = GetElementByTagFromXml(employee, "FirstName");
                        rpEmployeePeriod.Surname = GetElementByTagFromXml(employee, "LastName");
                        rpEmployeePeriod.Fullname = rpEmployeePeriod.Title + " " + rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                        rpEmployeePeriod.RefFullname = rpEmployeePeriod.Reference + " " + rpEmployeePeriod.Fullname;
                        string[] address = new string[6];
                        address[0] = GetElementByTagFromXml(employee, "Address1");
                        address[1] = GetElementByTagFromXml(employee, "Address2");
                        address[2] = GetElementByTagFromXml(employee, "Address3");
                        address[3] = GetElementByTagFromXml(employee, "Address4");
                        address[4] = GetElementByTagFromXml(employee, "Postcode");
                        address[5] = GetElementByTagFromXml(employee, "Country");

                        rpEmployeePeriod.DateOfBirth = GetDateElementByTagFromXml(employee, "DateOfBirth");
                        rpEmployeePeriod.Gender = GetElementByTagFromXml(employee, "Gender");

                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        if (leaver == "Y")
                        {
                            rpEmployeePeriod.Leaver = true;
                        }
                        else
                        {
                            rpEmployeePeriod.Leaver = false;
                        }
                        if (rpEmployeePeriod.Leaver)
                        {
                            rpEmployeePeriod.LeavingDate = GetDateElementByTagFromXml(employee, "LeavingDate");

                        }

                        rpEmployeePeriod.SortCode = GetElementByTagFromXml(employee, "SortCode");
                        rpEmployeePeriod.BankAccNo = GetElementByTagFromXml(employee, "BankAccNo");
                        rpEmployeePeriod.BuildingSocRef = GetElementByTagFromXml(employee, "BuildingSocRef");
                        rpEmployeePeriod.NINumber = GetElementByTagFromXml(employee, "NiNumber");
                        rpEmployeePeriod.NILetter = GetElementByTagFromXml(employee, "NiLetter");
                        rpEmployeePeriod.TaxCode = GetElementByTagFromXml(employee, "TaxCode");
                        rpEmployeePeriod.Week1Month1 = GetBooleanElementByTagFromXml(employee, "Week1Month1");
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            rpEmployeePeriod.TaxCode = rpEmployeePeriod.TaxCode + " W1";
                        }
                        rpEmployeePeriod.Frequency = rpParameters.PaySchedule;
                        rpEmployeePeriod.PaymentMethod = GetElementByTagFromXml(employee, "PayMethod");
                        rpEmployeePeriod.NetPayTP = GetDecimalElementByTagFromXml(employee, "Net");
                        rpEmployeePeriod.NetPayYTD = 0;
                        rpEmployeePeriod.TaxablePayTP = GetDecimalElementByTagFromXml(employee, "TaxablePayThisPeriod");
                        rpEmployeePeriod.TaxablePayYTD = GetDecimalElementByTagFromXml(employee, "TaxablePayThisYTD") + GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TaxablePayPrevious = GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TotalPayTP = 0;
                        rpEmployeePeriod.TotalPayYTD = 0;
                        rpEmployeePeriod.TotalDedTP = 0;
                        rpEmployeePeriod.TotalDedYTD = 0;
                        rpEmployeePeriod.ErNICTP = GetDecimalElementByTagFromXml(employee, "ErContributionsTaxPeriod");
                        rpEmployeePeriod.ErNICYTD = GetDecimalElementByTagFromXml(employee, "ErContributions");
                        rpEmployeePeriod.ErPensionTP = GetDecimalElementByTagFromXml(employee, "ErPensionTaxPeriod");
                        rpEmployeePeriod.ErPensionYTD = GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                        rpEmployeePeriod.EePensionTP = GetDecimalElementByTagFromXml(employee, "EePensionTaxPeriod");
                        rpEmployeePeriod.EePensionYTD = GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                        rpEmployeePeriod.PensionablePay = GetDecimalElementByTagFromXml(employee, "PensionablePay");
                        rpEmployeePeriod.PensionCode = GetElementByTagFromXml(employee, "PensionDetails");
                        if (rpEmployeePeriod.PensionCode != null)
                        {
                            //Just use the part after the last "/".
                            int i = rpEmployeePeriod.PensionCode.LastIndexOf("/") + 1;
                            int j = rpEmployeePeriod.PensionCode.Length;
                            rpEmployeePeriod.PensionCode = rpEmployeePeriod.PensionCode.Substring(i, j - i);
                        }
                        rpEmployeePeriod.ErContributionPercent = GetDecimalElementByTagFromXml(employee, "ErContributionPercent") * 100;
                        rpEmployeePeriod.EeContributionPercent = GetDecimalElementByTagFromXml(employee, "EeContributionPercent") * 100;
                        rpEmployeePeriod.PreTaxAddDed = 0;
                        rpEmployeePeriod.GUCosts = 0;
                        rpEmployeePeriod.AbsencePay = 0;
                        rpEmployeePeriod.HolidayPay = 0;
                        rpEmployeePeriod.PreTaxPension = 0;
                        rpEmployeePeriod.Tax = 0;
                        rpEmployeePeriod.TaxPrev = GetDecimalElementByTagFromXml(employee, "TaxPrevious");
                        rpEmployeePeriod.TaxThis = GetDecimalElementByTagFromXml(employee, "TaxThis");
                        rpEmployeePeriod.NetNI = 0;
                        rpEmployeePeriod.PostTaxAddDed = 0;
                        rpEmployeePeriod.PostTaxPension = 0;
                        rpEmployeePeriod.AOE = 0;
                        rpEmployeePeriod.StudentLoan = 0;

                        List<RPAddition> rpAdditions = new List<RPAddition>();
                        List<RPDeduction> rpDeductions = new List<RPDeduction>();
                        foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                        {
                            foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                            {
                                RPPayComponent rpPayComponent = new RPPayComponent();
                                rpPayComponent.PayCode = GetElementByTagFromXml(payCode, "Code");
                                rpPayComponent.Description = GetElementByTagFromXml(payCode, "Description");
                                rpPayComponent.EeRef = rpEmployeePeriod.Reference;
                                rpPayComponent.Fullname = rpEmployeePeriod.Fullname;
                                rpPayComponent.Surname = rpEmployeePeriod.Surname;
                                rpPayComponent.Rate = GetDecimalElementByTagFromXml(payCode, "Rate");
                                rpPayComponent.UnitsTP = GetDecimalElementByTagFromXml(payCode, "Units");
                                rpPayComponent.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount");
                                rpPayComponent.UnitsYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                rpPayComponent.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                if (rpPayComponent.AmountTP != 0 || rpPayComponent.AmountYTD != 0)
                                {
                                    if (GetElementByTagFromXml(payCode, "IsPayCode") == "true")
                                    {
                                        rpPayComponents.Add(rpPayComponent);
                                    }
                                    //Check for the different pay codes and add to the appropriate total.
                                    switch (rpPayComponent.PayCode)
                                    {
                                        case "HOLPY":
                                        case "HOLIDAY":
                                            rpEmployeePeriod.HolidayPay = rpEmployeePeriod.HolidayPay + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSION":
                                            rpEmployeePeriod.PreTaxPension = rpEmployeePeriod.PreTaxPension + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSIONRAS":
                                        case "PENSIONSS":
                                            rpEmployeePeriod.PostTaxPension = rpEmployeePeriod.PostTaxPension + rpPayComponent.AmountTP;
                                            break;
                                        case "AOE":
                                            rpEmployeePeriod.AOE = rpEmployeePeriod.AOE + rpPayComponent.AmountTP;
                                            break;
                                        case "SLOAN":
                                            rpEmployeePeriod.StudentLoan = rpEmployeePeriod.StudentLoan + rpPayComponent.AmountTP;
                                            break;
                                        case "TAX":
                                            rpEmployeePeriod.Tax = rpEmployeePeriod.Tax + rpPayComponent.AmountTP;
                                            break;
                                        case "NI":
                                            rpEmployeePeriod.NetNI = rpEmployeePeriod.NetNI + rpPayComponent.AmountTP;
                                            break;
                                        case "SAP":
                                        case "SHPP":
                                        case "SMP":
                                        case "SPP":
                                        case "SSP":
                                            rpEmployeePeriod.AbsencePay = rpEmployeePeriod.AbsencePay + rpPayComponent.AmountTP;
                                            break;
                                        default:
                                            rpEmployeePeriod.PreTaxAddDed = rpEmployeePeriod.PreTaxAddDed + rpPayComponent.AmountTP;
                                            break;

                                    }
                                }


                                if (GetElementByTagFromXml(payCode, "EarningOrDeduction") == "E")
                                {
                                    RPAddition rpAddition = new RPAddition();
                                    rpAddition.EeRef = rpEmployeePeriod.Reference;
                                    rpAddition.Code = GetElementByTagFromXml(payCode, "Code");
                                    rpAddition.Description = GetElementByTagFromXml(payCode, "Description");
                                    rpAddition.Rate = GetDecimalElementByTagFromXml(payCode, "Rate");
                                    rpAddition.Units = GetDecimalElementByTagFromXml(payCode, "Units");
                                    rpAddition.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount");
                                    rpAddition.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                    if (rpAddition.AmountTP != 0 || rpAddition.AmountYTD != 0)
                                    {
                                        rpAdditions.Add(rpAddition);
                                        rpEmployeePeriod.TotalPayTP = rpEmployeePeriod.TotalPayTP + rpAddition.AmountTP;
                                        rpEmployeePeriod.TotalPayYTD = rpEmployeePeriod.TotalPayYTD + rpAddition.AmountYTD;
                                    }

                                }
                                else
                                {
                                    RPDeduction rpDeduction = new RPDeduction();
                                    rpDeduction.EeRef = rpEmployeePeriod.Reference;
                                    rpDeduction.Code = GetElementByTagFromXml(payCode, "Code");
                                    rpDeduction.Description = GetElementByTagFromXml(payCode, "Description");
                                    rpDeduction.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount") * -1;
                                    rpDeduction.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance") * -1;
                                    if (rpDeduction.AmountTP != 0 || rpDeduction.AmountYTD != 0)
                                    {
                                        rpDeductions.Add(rpDeduction);
                                        rpEmployeePeriod.TotalDedTP = rpEmployeePeriod.TotalDedTP + rpDeduction.AmountTP;
                                        rpEmployeePeriod.TotalDedYTD = rpEmployeePeriod.TotalDedYTD + rpDeduction.AmountYTD;
                                    }

                                }
                                rpEmployeePeriod.Additions = rpAdditions;
                                rpEmployeePeriod.Deductions = rpDeductions;
                            }//End of for each payCode
                        }//End of for each payCodes
                         //Multiple Tax and NI by -1 to make them positive
                        rpEmployeePeriod.Tax = rpEmployeePeriod.Tax * -1;
                        rpEmployeePeriod.NetNI = rpEmployeePeriod.NetNI * -1;
                        //Create a P45 object if the employee is a leaver
                        if (rpEmployeePeriod.Leaver)
                        {
                            P45 p45 = new P45();
                            p45.ErOfficeNo = rpEmployer.PayeRef.Substring(0, 3);
                            p45.ErRefNo = rpEmployer.PayeRef.Substring(4);
                            p45.NINumber = rpEmployeePeriod.NINumber;
                            p45.Title = rpEmployeePeriod.Title;
                            p45.Surname = rpEmployeePeriod.Surname;
                            p45.FirstNames = rpEmployeePeriod.Forename;
                            p45.LeavingDate = rpEmployeePeriod.LeavingDate;
                            p45.DateOfBirth = rpEmployeePeriod.DateOfBirth;
                            p45.StudentLoansDeductionToContinue = false;  //Need to find out where this comes from!
                            p45.TaxCode = rpEmployeePeriod.TaxCode;
                            p45.Week1Month1 = rpEmployeePeriod.Week1Month1;
                            if (rpParameters.PaySchedule == "Monthly")
                            {
                                p45.MonthNo = rpParameters.TaxPeriod;
                                p45.WeekNo = 0;
                            }
                            else
                            {
                                p45.MonthNo = 0;
                                p45.WeekNo = rpParameters.TaxPeriod;
                            }
                            p45.PayToDate = rpEmployeePeriod.TotalPayYTD; //rpEmployeePeriod.TaxablePayYTD + rpEmployeePeriod.TaxablePayPrevious;
                            p45.TaxToDate = rpEmployeePeriod.TaxThis + rpEmployeePeriod.TaxPrev;
                            p45.PayThis = rpEmployeePeriod.TotalPayYTD - rpEmployeePeriod.TaxablePayPrevious;    //rpEmployeePeriod.TaxablePayYTD;
                            p45.TaxThis = rpEmployeePeriod.TaxThis;
                            p45.EeRef = rpEmployeePeriod.Reference;
                            if (rpEmployeePeriod.Gender == "Male")
                            {
                                p45.IsMale = true;
                            }
                            else
                            {
                                p45.IsMale = false;
                            }
                            p45.Address1 = address[0];
                            p45.Address2 = address[1];
                            p45.Address3 = address[2];
                            p45.Address4 = address[3];
                            p45.Postcode = address[4];
                            p45.Country = address[5];
                            p45.ErName = rpEmployer.Name;
                            p45.ErAddress1 = "19 Island Hill";// rpEmployer.Address1;
                            p45.ErAddress2 = "Dromara Road";// rpEmployer.Address2;
                            p45.ErAddress3 = "Dromore";// rpEmployer.Address3;
                            p45.ErAddress4 = "Co. Down";// rpEmployer.Address4;
                            p45.ErPostcode = "BT25 1HA";// rpEmployer.Postcode;
                            p45.ErCountry = "United Kingdom";// rpEmployer.Country;
                            p45.Now = DateTime.Now;

                            p45s.Add(p45);
                        }
                        //Re-Arrange the employees address so that there are no blank lines shown.
                        address = RemoveBlankAddressLines(address);
                        rpEmployeePeriod.Address1 = address[0];
                        rpEmployeePeriod.Address2 = address[1];
                        rpEmployeePeriod.Address3 = address[2];
                        rpEmployeePeriod.Address4 = address[3];
                        rpEmployeePeriod.Postcode = address[4];
                        rpEmployeePeriod.Country = address[5];
                        rpEmployeePeriodList.Add(rpEmployeePeriod);
                    }//End of for each employee


                }
                //Put a sort of the pay codes in here
                //I didn't need to sort it in the end because the DevExpress report does it but this is useful code for future reference.
                //rpPayComponents.Sort(delegate (RPPayComponent x, RPPayComponent y)
                //{
                //    if (x.Description == null && y.Description == null) return 0;
                //    else if (x.Description == null) return -1;
                //    else if (y.Description == null) return 1;
                //    else return x.Description.CompareTo(y.Description);
                //});
                //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
                decimal hmrcTotal = CalculateHMRCTotal(rpEmployeePeriodList);
                string hmrcDesc = "[" + hmrcTotal.ToString() + "]";
                //I now have a list of employee with their total for this period & ytd plus addition & deductions
                //I can print payslips from here.
                PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents);
                
                ZipReports(xdoc, rpEmployer, rpParameters, hmrcDesc);
                EmailZippedReports(xdoc, rpEmployer, rpParameters);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error preparing reports.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

        }

        private Tuple<List<RPEmployeePeriod> ,List<RPPayComponent>, List<P45>, RPEmployer> PrepareStandardReports(XDocument xdoc, XmlDocument xmlReport)
        {
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;

            List<RPEmployeePeriod> rpEmployeePeriodList = new List<RPEmployeePeriod>();
            List<P45> p45s = new List<P45>();
            //Create a list of Pay Code totals for the Payroll Component Analysis report
            List<RPPayComponent> rpPayComponents = new List<RPPayComponent>();
            RPParameters rpParameters = new RPParameters();
            RPEmployer rpEmployer = new RPEmployer();

            try
            {
                foreach (XmlElement parameter in xmlReport.GetElementsByTagName("Parameters"))
                {
                    rpParameters.ErRef = GetElementByTagFromXml(parameter, "EmployerCode");
                    rpParameters.TaxYear = GetIntElementByTagFromXml(parameter, "TaxYear");
                    rpParameters.AccYearStart = GetDateElementByTagFromXml(parameter, "AccountingYearStartDate");
                    rpParameters.AccYearEnd = GetDateElementByTagFromXml(parameter, "AccountingYearEndDate");
                    rpParameters.TaxPeriod = GetIntElementByTagFromXml(parameter, "TaxPeriod");
                    rpParameters.PaySchedule = GetElementByTagFromXml(parameter, "PaySchedule");
                }
                
                foreach (XmlElement employer in xmlReport.GetElementsByTagName("Employer"))
                {
                    rpEmployer.Name = GetElementByTagFromXml(employer, "Name");
                    rpEmployer.PayeRef = GetElementByTagFromXml(employer, "EmployerPayeRef");
                }

                

                bool payRunDate = false;
                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;

                    if (GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                    {
                        if (!payRunDate)
                        {
                            rpParameters.PayRunDate = GetDateElementByTagFromXml(employee, "PayRunDate");
                            payRunDate = true;
                        }
                        //If the employee is a leaver before the start date then don't include.
                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        DateTime leavingDate = new DateTime();
                        if (GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            leavingDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        }
                        DateTime periodStartDate = DateTime.ParseExact(GetElementByTagFromXml(employee, "PeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        if (leaver.StartsWith("N"))
                        {
                            include = true;
                        }
                        else if (leavingDate >= periodStartDate)
                        {
                            include = true;
                        }
                    }

                    if (include)
                    {
                        RPEmployeePeriod rpEmployeePeriod = new RPEmployeePeriod();
                        rpEmployeePeriod.Reference = GetElementByTagFromXml(employee, "EeRef");
                        rpEmployeePeriod.Title = GetElementByTagFromXml(employee, "Title");
                        rpEmployeePeriod.Forename = GetElementByTagFromXml(employee, "FirstName");
                        rpEmployeePeriod.Surname = GetElementByTagFromXml(employee, "LastName");
                        rpEmployeePeriod.Fullname = rpEmployeePeriod.Title + " " + rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                        rpEmployeePeriod.RefFullname = rpEmployeePeriod.Reference + " " + rpEmployeePeriod.Fullname;
                        string[] address = new string[6];
                        address[0] = GetElementByTagFromXml(employee, "Address1");
                        address[1] = GetElementByTagFromXml(employee, "Address2");
                        address[2] = GetElementByTagFromXml(employee, "Address3");
                        address[3] = GetElementByTagFromXml(employee, "Address4");
                        address[4] = GetElementByTagFromXml(employee, "Postcode");
                        address[5] = GetElementByTagFromXml(employee, "Country");

                        rpEmployeePeriod.SortCode = GetElementByTagFromXml(employee, "SortCode");
                        rpEmployeePeriod.BankAccNo = GetElementByTagFromXml(employee, "BankAccNo");
                        rpEmployeePeriod.DateOfBirth = GetDateElementByTagFromXml(employee, "DateOfBirth");
                        rpEmployeePeriod.Gender = GetElementByTagFromXml(employee, "Gender");
                        rpEmployeePeriod.BuildingSocRef = GetElementByTagFromXml(employee, "BuildingSocRef");
                        rpEmployeePeriod.NINumber = GetElementByTagFromXml(employee, "NiNumber");
                        rpEmployeePeriod.PaymentMethod = GetElementByTagFromXml(employee, "PayMethod");
                        rpEmployeePeriod.PayRunDate = GetDateElementByTagFromXml(employee, "PayRunDate");
                        rpEmployeePeriod.PeriodStartDate = GetDateElementByTagFromXml(employee, "PeriodStartDate");
                        rpEmployeePeriod.PeriodEndDate = GetDateElementByTagFromXml(employee, "PeriodEndDate");
                        rpEmployeePeriod.PayrollYear = GetIntElementByTagFromXml(employee, "PayrollYear");
                        rpEmployeePeriod.Gross = GetDecimalElementByTagFromXml(employee, "Gross");
                        rpEmployeePeriod.NetPayTP = GetDecimalElementByTagFromXml(employee, "Net");
                        rpEmployeePeriod.DayHours = GetIntElementByTagFromXml(employee, "DayHours");
                        rpEmployeePeriod.StudentLoanStartDate = GetDateElementByTagFromXml(employee, "StudentLoanStartDate");
                        rpEmployeePeriod.StudentLoanEndDate = GetDateElementByTagFromXml(employee, "StudentLoanEndDate");
                        rpEmployeePeriod.NILetter = GetElementByTagFromXml(employee, "NiLetter");
                        rpEmployeePeriod.CalculationBasis = GetElementByTagFromXml(employee, "CalculationBasis");
                        //TotalPayTP
                        rpEmployeePeriod.EarningsToLEL = GetDecimalElementByTagFromXml(employee, "EarningsToLEL");
                        rpEmployeePeriod.EarningsToSET = GetDecimalElementByTagFromXml(employee, "EarningsToSET");
                        rpEmployeePeriod.EarningsToPET = GetDecimalElementByTagFromXml(employee, "EarningsToPET");
                        rpEmployeePeriod.EarningsToUST = GetDecimalElementByTagFromXml(employee, "EarningsToUST");
                        rpEmployeePeriod.EarningsToAUST = GetDecimalElementByTagFromXml(employee, "EarningsToAUST");
                        rpEmployeePeriod.EarningsToUEL = GetDecimalElementByTagFromXml(employee, "EarningsToUEL");
                        rpEmployeePeriod.EarningsAboveUEL = GetDecimalElementByTagFromXml(employee, "EarningsAboveUEL");
                        rpEmployeePeriod.EeContributionsPt1 = GetDecimalElementByTagFromXml(employee, "EeContributionsPt1");
                        rpEmployeePeriod.EeContributionsPt2 = GetDecimalElementByTagFromXml(employee, "EeContributions2");
                        rpEmployeePeriod.ErNICYTD = GetDecimalElementByTagFromXml(employee, "ErContributions");
                        rpEmployeePeriod.EeRebate = GetDecimalElementByTagFromXml(employee, "EeRabate");
                        rpEmployeePeriod.ErRebate = GetDecimalElementByTagFromXml(employee, "ErRebate");
                        rpEmployeePeriod.EeReduction = GetDecimalElementByTagFromXml(employee, "EeReduction");
                        string leaver = GetElementByTagFromXml(employee, "Leaver");
                        if (leaver == "Y")
                        {
                            rpEmployeePeriod.Leaver = true;
                        }
                        else
                        {
                            rpEmployeePeriod.Leaver = false;
                        }
                        if (rpEmployeePeriod.Leaver)
                        {
                            rpEmployeePeriod.LeavingDate = GetDateElementByTagFromXml(employee, "LeavingDate");

                        }
                        rpEmployeePeriod.TaxCode = GetElementByTagFromXml(employee, "TaxCode");
                        rpEmployeePeriod.Week1Month1 = GetBooleanElementByTagFromXml(employee, "Week1Month1");
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            rpEmployeePeriod.TaxCode = rpEmployeePeriod.TaxCode + " W1";
                        }
                        rpEmployeePeriod.TaxCodeChangeTypeID = GetElementByTagFromXml(employee, "TaxCodeChangeTypeID");
                        rpEmployeePeriod.TaxCodeChangeType = GetElementByTagFromXml(employee, "TaxCodeChangeType");
                        rpEmployeePeriod.TaxPrev = GetDecimalElementByTagFromXml(employee, "TaxPrevious");
                        rpEmployeePeriod.TaxablePayPrevious = GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TaxThis = GetDecimalElementByTagFromXml(employee, "TaxThis");
                        rpEmployeePeriod.TaxablePayYTD = GetDecimalElementByTagFromXml(employee, "TaxablePayThisYTD") + GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TaxablePayTP = GetDecimalElementByTagFromXml(employee, "TaxablePayThisPeriod");
                        rpEmployeePeriod.HolidayAccruedYTD = GetDecimalElementByTagFromXml(employee, "HolidayAccruedTd");
                        rpEmployeePeriod.ErPensionYTD = GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                        rpEmployeePeriod.EePensionYTD = GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                        rpEmployeePeriod.ErPensionTP = GetDecimalElementByTagFromXml(employee, "ErPensionTaxPeriod");
                        rpEmployeePeriod.EePensionTP = GetDecimalElementByTagFromXml(employee, "EePensionTaxPeriod");
                        rpEmployeePeriod.ErContributionPercent = GetDecimalElementByTagFromXml(employee, "ErContributionPercent") * 100;
                        rpEmployeePeriod.EeContributionPercent = GetDecimalElementByTagFromXml(employee, "EeContributionPercent") * 100;
                        rpEmployeePeriod.PensionablePay = GetDecimalElementByTagFromXml(employee, "PensionablePay");
                        rpEmployeePeriod.ErPensionPayRunDate = GetDateElementByTagFromXml(employee, "ErPensionPayRunDate");
                        rpEmployeePeriod.EePensionPayRunDate = GetDateElementByTagFromXml(employee, "EePensionPayRunDate");
                        rpEmployeePeriod.DirectorshipAppointmentDate = GetDateElementByTagFromXml(employee, "DirectorshipAppointmentDate");
                        rpEmployeePeriod.Director = GetBooleanElementByTagFromXml(employee, "Director");
                        rpEmployeePeriod.EeContributionTaxPeriodPt1 = GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt1");
                        rpEmployeePeriod.EeContributionTaxPeriodPt2 = GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt2");
                        rpEmployeePeriod.ErNICTP = GetDecimalElementByTagFromXml(employee, "ErContributionTaxPeriod");
                        rpEmployeePeriod.Frequency = rpParameters.PaySchedule;
                        rpEmployeePeriod.NetPayYTD = 0;
                        rpEmployeePeriod.TotalPayTP = 0;
                        rpEmployeePeriod.TotalPayYTD = 0;
                        rpEmployeePeriod.TotalDedTP = 0;
                        rpEmployeePeriod.TotalDedYTD = 0;
                        rpEmployeePeriod.ErNICTP = GetDecimalElementByTagFromXml(employee, "ErContributionsTaxPeriod");
                        rpEmployeePeriod.ErNICYTD = GetDecimalElementByTagFromXml(employee, "ErContributions");
                        rpEmployeePeriod.PensionCode = GetElementByTagFromXml(employee, "PensionDetails");
                        if (rpEmployeePeriod.PensionCode != null)
                        {
                            //Just use the part after the last "/".
                            int i = rpEmployeePeriod.PensionCode.LastIndexOf("/") + 1;
                            int j = rpEmployeePeriod.PensionCode.Length;
                            rpEmployeePeriod.PensionCode = rpEmployeePeriod.PensionCode.Substring(i, j - i);
                        }
                        rpEmployeePeriod.PreTaxAddDed = 0;
                        rpEmployeePeriod.GUCosts = 0;
                        rpEmployeePeriod.AbsencePay = 0;
                        rpEmployeePeriod.HolidayPay = 0;
                        rpEmployeePeriod.PreTaxPension = 0;
                        rpEmployeePeriod.Tax = 0;
                        rpEmployeePeriod.NetNI = 0;
                        rpEmployeePeriod.PostTaxAddDed = 0;
                        rpEmployeePeriod.PostTaxPension = 0;
                        rpEmployeePeriod.AOE = 0;
                        rpEmployeePeriod.StudentLoan = 0;

                        List<RPAddition> rpAdditions = new List<RPAddition>();
                        List<RPDeduction> rpDeductions = new List<RPDeduction>();
                        foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                        {
                            foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                            {
                                RPPayComponent rpPayComponent = new RPPayComponent();
                                rpPayComponent.PayCode = GetElementByTagFromXml(payCode, "Code");
                                rpPayComponent.Description = GetElementByTagFromXml(payCode, "Description");
                                rpPayComponent.EeRef = rpEmployeePeriod.Reference;
                                rpPayComponent.Fullname = rpEmployeePeriod.Fullname;
                                rpPayComponent.Surname = rpEmployeePeriod.Surname;
                                rpPayComponent.Rate = GetDecimalElementByTagFromXml(payCode, "Rate");
                                rpPayComponent.UnitsTP = GetDecimalElementByTagFromXml(payCode, "Units");
                                rpPayComponent.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount");
                                rpPayComponent.UnitsYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                rpPayComponent.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                if (rpPayComponent.AmountTP != 0 || rpPayComponent.AmountYTD != 0)
                                {
                                    if (GetElementByTagFromXml(payCode, "IsPayCode") == "true")
                                    {
                                        rpPayComponents.Add(rpPayComponent);
                                    }
                                    //Check for the different pay codes and add to the appropriate total.
                                    switch (rpPayComponent.PayCode)
                                    {
                                        case "HOLPY":
                                        case "HOLIDAY":
                                            rpEmployeePeriod.HolidayPay = rpEmployeePeriod.HolidayPay + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSION":
                                            rpEmployeePeriod.PreTaxPension = rpEmployeePeriod.PreTaxPension + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSIONRAS":
                                        case "PENSIONSS":
                                            rpEmployeePeriod.PostTaxPension = rpEmployeePeriod.PostTaxPension + rpPayComponent.AmountTP;
                                            break;
                                        case "AOE":
                                            rpEmployeePeriod.AOE = rpEmployeePeriod.AOE + rpPayComponent.AmountTP;
                                            break;
                                        case "SLOAN":
                                            rpEmployeePeriod.StudentLoan = rpEmployeePeriod.StudentLoan + rpPayComponent.AmountTP;
                                            break;
                                        case "TAX":
                                            rpEmployeePeriod.Tax = rpEmployeePeriod.Tax + rpPayComponent.AmountTP;
                                            break;
                                        case "NI":
                                            rpEmployeePeriod.NetNI = rpEmployeePeriod.NetNI + rpPayComponent.AmountTP;
                                            break;
                                        case "SAP":
                                        case "SHPP":
                                        case "SMP":
                                        case "SPP":
                                        case "SSP":
                                            rpEmployeePeriod.AbsencePay = rpEmployeePeriod.AbsencePay + rpPayComponent.AmountTP;
                                            break;
                                        default:
                                            rpEmployeePeriod.PreTaxAddDed = rpEmployeePeriod.PreTaxAddDed + rpPayComponent.AmountTP;
                                            break;

                                    }
                                }


                                if (GetElementByTagFromXml(payCode, "EarningOrDeduction") == "E")
                                {
                                    RPAddition rpAddition = new RPAddition();
                                    rpAddition.EeRef = rpEmployeePeriod.Reference;
                                    rpAddition.Code = GetElementByTagFromXml(payCode, "Code");
                                    rpAddition.Description = GetElementByTagFromXml(payCode, "Description");
                                    rpAddition.Rate = GetDecimalElementByTagFromXml(payCode, "Rate");
                                    rpAddition.Units = GetDecimalElementByTagFromXml(payCode, "Units");
                                    rpAddition.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount");
                                    rpAddition.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                    rpAddition.AccountsYearBalance = GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance");
                                    rpAddition.AccountsYearUnits = GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits");
                                    rpAddition.PayeYearUnits = GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                    rpAddition.PayrollAccrued = GetDecimalElementByTagFromXml(payCode, "PayrollAccrued");
                                    if (rpAddition.AmountTP != 0 || rpAddition.AmountYTD != 0)
                                    {
                                        rpAdditions.Add(rpAddition);
                                        rpEmployeePeriod.TotalPayTP = rpEmployeePeriod.TotalPayTP + rpAddition.AmountTP;
                                        rpEmployeePeriod.TotalPayYTD = rpEmployeePeriod.TotalPayYTD + rpAddition.AmountYTD;
                                    }

                                }
                                else
                                {
                                    RPDeduction rpDeduction = new RPDeduction();
                                    rpDeduction.EeRef = rpEmployeePeriod.Reference;
                                    rpDeduction.Code = GetElementByTagFromXml(payCode, "Code");
                                    rpDeduction.Description = GetElementByTagFromXml(payCode, "Description");
                                    rpDeduction.AmountTP = GetDecimalElementByTagFromXml(payCode, "Amount") * -1;
                                    rpDeduction.AmountYTD = GetDecimalElementByTagFromXml(payCode, "PayeYearBalance") * -1;
                                    rpDeduction.AccountsYearBalance = GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance") * -1;
                                    rpDeduction.AccountsYearUnits = GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits") * -1;
                                    rpDeduction.PayeYearUnits = GetDecimalElementByTagFromXml(payCode, "PayeYearUnits") * -1;
                                    rpDeduction.PayrollAccrued = GetDecimalElementByTagFromXml(payCode, "PayrollAccrued") * -1;
                                    if (rpDeduction.AmountTP != 0 || rpDeduction.AmountYTD != 0)
                                    {
                                        rpDeductions.Add(rpDeduction);
                                        rpEmployeePeriod.TotalDedTP = rpEmployeePeriod.TotalDedTP + rpDeduction.AmountTP;
                                        rpEmployeePeriod.TotalDedYTD = rpEmployeePeriod.TotalDedYTD + rpDeduction.AmountYTD;
                                    }

                                }
                                rpEmployeePeriod.Additions = rpAdditions;
                                rpEmployeePeriod.Deductions = rpDeductions;
                            }//End of for each payCode
                        }//End of for each payCodes
                         //Multiple Tax and NI by -1 to make them positive
                        rpEmployeePeriod.Tax = rpEmployeePeriod.Tax * -1;
                        rpEmployeePeriod.NetNI = rpEmployeePeriod.NetNI * -1;
                        //Create a P45 object if the employee is a leaver
                        if (rpEmployeePeriod.Leaver)
                        {
                            P45 p45 = new P45();
                            p45.ErOfficeNo = rpEmployer.PayeRef.Substring(0, 3);
                            p45.ErRefNo = rpEmployer.PayeRef.Substring(4);
                            p45.NINumber = rpEmployeePeriod.NINumber;
                            p45.Title = rpEmployeePeriod.Title;
                            p45.Surname = rpEmployeePeriod.Surname;
                            p45.FirstNames = rpEmployeePeriod.Forename;
                            p45.LeavingDate = rpEmployeePeriod.LeavingDate;
                            p45.DateOfBirth = rpEmployeePeriod.DateOfBirth;
                            p45.StudentLoansDeductionToContinue = false;  //Need to find out where this comes from!
                            p45.TaxCode = rpEmployeePeriod.TaxCode;
                            p45.Week1Month1 = rpEmployeePeriod.Week1Month1;
                            if (rpParameters.PaySchedule == "Monthly")
                            {
                                p45.MonthNo = rpParameters.TaxPeriod;
                                p45.WeekNo = 0;
                            }
                            else
                            {
                                p45.MonthNo = 0;
                                p45.WeekNo = rpParameters.TaxPeriod;
                            }
                            p45.PayToDate = rpEmployeePeriod.TotalPayYTD; //rpEmployeePeriod.TaxablePayYTD + rpEmployeePeriod.TaxablePayPrevious;
                            p45.TaxToDate = rpEmployeePeriod.TaxThis + rpEmployeePeriod.TaxPrev;
                            p45.PayThis = rpEmployeePeriod.TotalPayYTD - rpEmployeePeriod.TaxablePayPrevious;    //rpEmployeePeriod.TaxablePayYTD;
                            p45.TaxThis = rpEmployeePeriod.TaxThis;
                            p45.EeRef = rpEmployeePeriod.Reference;
                            if (rpEmployeePeriod.Gender == "Male")
                            {
                                p45.IsMale = true;
                            }
                            else
                            {
                                p45.IsMale = false;
                            }
                            p45.Address1 = address[0];
                            p45.Address2 = address[1];
                            p45.Address3 = address[2];
                            p45.Address4 = address[3];
                            p45.Postcode = address[4];
                            p45.Country = address[5];
                            p45.ErName = rpEmployer.Name;
                            p45.ErAddress1 = "19 Island Hill";// rpEmployer.Address1;
                            p45.ErAddress2 = "Dromara Road";// rpEmployer.Address2;
                            p45.ErAddress3 = "Dromore";// rpEmployer.Address3;
                            p45.ErAddress4 = "Co. Down";// rpEmployer.Address4;
                            p45.ErPostcode = "BT25 1HA";// rpEmployer.Postcode;
                            p45.ErCountry = "United Kingdom";// rpEmployer.Country;
                            p45.Now = DateTime.Now;

                            p45s.Add(p45);
                        }
                        //Re-Arrange the employees address so that there are no blank lines shown.
                        address = RemoveBlankAddressLines(address);
                        rpEmployeePeriod.Address1 = address[0];
                        rpEmployeePeriod.Address2 = address[1];
                        rpEmployeePeriod.Address3 = address[2];
                        rpEmployeePeriod.Address4 = address[3];
                        rpEmployeePeriod.Postcode = address[4];
                        rpEmployeePeriod.Country = address[5];
                        rpEmployeePeriodList.Add(rpEmployeePeriod);
                    }//End of for each employee


                }
                
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error preparing reports.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, RPEmployer>(rpEmployeePeriodList, rpPayComponents, p45s, rpEmployer);
            
        }
        private void PrintStandardReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters, List<P45> p45s, List<RPPayComponent> rpPayComponents)
        {
            PrintPayslips(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPaymentsDueByMethodReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintComponentAnalysisReport(xdoc, rpPayComponents, rpEmployer, rpParameters);
            PrintPensionContributionsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPayrollRunDetailsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            if (p45s.Count > 0)
            {
                PrintP45s(xdoc, p45s, rpParameters);
            }
        }
        private string[] RemoveBlankAddressLines(string[] oldAddress)
        {
            string[] newAddress = new string[6];
            int x = 0;
            for (int i = 0; i < 6; i++)
            {
                if (oldAddress[i] != "" && oldAddress[i] != " " && oldAddress[i] != null)
                {
                    newAddress[x] = oldAddress[i];
                    x++;
                }
            }
            for (int i = x; i < 6; i++)
            {
                newAddress[i] = "";
            }
            return newAddress;
        }
        private decimal CalculateHMRCTotal(List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            decimal hmrcTotal = 0;
            foreach (RPEmployeePeriod employee in rpEmployeePeriodList)
            {
                hmrcTotal = hmrcTotal + employee.Tax + employee.NetNI + employee.ErNICTP;
            }
            return hmrcTotal;
        }
        private void PrintPayslips(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "Payslip.repx", true);
            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate; //.AccYearEnd;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";

                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PayslipReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);

            }

        }
        private void PrintPaymentsDueByMethodReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PaymentsDueByMethodsReport.repx", true);         //"PaymentsDueByMethodReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PaymentsDueByMethodReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintPensionContributionsReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PensionContributionsReport.repx", true);         //"PensionContributionsReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PensionContributionsReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintComponentAnalysisReport(XDocument xdoc, List<RPPayComponent> rpPayComponents, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "ComponentAnalysisReport.repx", true);         //"ComponentAnalysisReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpPayComponents;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_ComponentAnalysisReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintPayrollRunDetailsReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            string coName = rpEmployer.Name;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            string freq = rpParameters.PaySchedule;
            //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
            var payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
            if (payeMonth <= 0)
            {
                payeMonth += 12;
            }

            //Main payslip report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "PayrollRunDetailsReport.repx", true);         //"PayrollRunDetailsReport.repx"

            report1.Parameters["CmpName"].Value = rpEmployer.Name;
            report1.Parameters["PayeRef"].Value = rpEmployer.PayeRef;
            report1.Parameters["Date"].Value = rpParameters.PayRunDate;
            report1.Parameters["Period"].Value = rpParameters.TaxPeriod;
            report1.Parameters["Freq"].Value = rpParameters.PaySchedule;
            report1.Parameters["PAYEMonth"].Value = payeMonth;
            report1.DataSource = rpEmployeePeriodList;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                //string dirName = "V:\\Payescape\\PayRunIO\\WG\\";
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_PayrollRunDetailsReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void PrintP45s(XDocument xdoc, List<P45> p45s, RPParameters rpParameters)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value + "Programs\\";
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            int taxYear = rpParameters.TaxYear;
            int taxPeriod = rpParameters.TaxPeriod;
            //P45 report
            XtraReport report1 = XtraReport.FromFile(softwareHomeFolder + "P45.repx", true);
            report1.DataSource = p45s;
            //// To show the report designer. You need to uncomment this to design the report.
            //// You also need to comment out the report.ExportToPDF line below
            ////
            bool designMode = false;
            if (designMode)
            {
                report1.ShowDesigner();
                //report1.ShowPreview();

            }
            else
            {
                // Export to pdf file.

                //
                // I'm going to remove spaces from the document name. I'll replace them with dashes
                //
                string dirName = outgoingFolder + "\\" + coNo + "\\";
                Directory.CreateDirectory(dirName);
                string docName = coNo + "_P45ReportFor_TaxYear_" + taxYear + "_Period_" + taxPeriod + ".pdf";

                report1.ExportToPdf(dirName + docName);
                docName = docName.Replace(".pdf", ".xlsx");
                report1.ExportToXlsx(dirName + docName);
            }

        }
        private void ZipReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters, string hmrcDesc)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            //
            // Zip the folder.
            //
            string dateTimeStamp = DateTime.Now.ToString("yyyyMMddhhmmssfff");
            string sourceFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef;
            string zipFileName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef + "_PDF_Reports_" + hmrcDesc + "_" + dateTimeStamp + ".zip";
            string zipFileNameNoPassword = xdoc.Root.Element("DataHomeFolder").Value + "PE-ReportsNoPassword\\" + rpParameters.ErRef + "_PDF_Reports_" + hmrcDesc + "_" + dateTimeStamp + ".zip";
            string password = null;
            password = rpEmployer.Name.ToLower().Replace(" ", "");
            if (password.Length >= 4)
            {
                password = password.Substring(0, 4);
            }
            password = password + rpParameters.ErRef;
            try
            {
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.Password = password;
                    zip.AddDirectory(sourceFolder);
                    zip.Save(zipFileName);
                }
                //Create a copy of the reports with no password for Emer & Mark
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.AddDirectory(sourceFolder);
                    zip.Save(zipFileNameNoPassword);
                }

                DeleteFilesThenFolder(xdoc, sourceFolder);

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error zipping pdf reports for source folder, {0}.\r\n{1}.\r\n", sourceFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }

        }
        private void DeleteFilesThenFolder(XDocument xdoc, string sourceFolder)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(sourceFolder);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
                Directory.Delete(sourceFolder);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error deleting files from source folder, {0}.\r\n{1}.\r\n", sourceFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        private void EmailZippedReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(reportFolder);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    EmailZippedReport(xdoc, file, rpEmployer, rpParameters);
                    file.MoveTo(file.FullName.Replace("PE-Reports", "PE-Reports\\Archive"));
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped pdf reports for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
        }
        private void EmailZippedReport(XDocument xdoc, FileInfo file, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            try
            {
                //
                // Send an email.
                //
                bool validEmailAddress = false;
                //Find amount due to HMRC in the file name.
                int x = file.FullName.LastIndexOf('[');
                int y = file.FullName.LastIndexOf(']');
                string hmrcDesc = file.FullName.Substring(x, y - x);
                hmrcDesc = hmrcDesc.Replace("[", "£");
                DateTime runDate = rpParameters.PayRunDate;
                runDate = runDate.AddMonths(1);
                int day = runDate.Day;
                day = 20 - day;
                runDate = runDate.AddDays(day);
                string dueDate = runDate.ToLongDateString();
                //string emailPassword = "fjbykfgxxkdgclfp"; //fjbykfgxxkdgclfp
                string mailSubject = String.Format("Payroll reports for tax year {0}, pay period {1}.", rpParameters.TaxYear, rpParameters.TaxPeriod);

                string mailBody = String.Format("Hi,\r\n\r\nPlease find attached payroll reports for tax year {0}, pay period {1}.\r\n\r\n" +
                                                "The amount payable to HMRC this month is {2}, this payment is due on or before {3}.\r\n\r\n" +
                                                "Please review and confirm if all is correct.\r\n\r\nKind Regards,\r\n\r\nThe Payescape Team."
                                                , rpParameters.TaxYear, rpParameters.TaxPeriod, hmrcDesc, dueDate);
                // Get currrent day of week.
                DayOfWeek today = DateTime.Today.DayOfWeek;
                string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
                string dataBase = xdoc.Root.Element("Database").Value;
                string userID = xdoc.Root.Element("Username").Value;
                string password = xdoc.Root.Element("Password").Value;
                string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
                //
                //Get the SMTP email settings from the database
                //
                SMTPEmailSettings smtpEmailSettings = new SMTPEmailSettings();
                smtpEmailSettings = GetEmailSettings(xdoc, sqlConnectionString);
                //
                //Get a list of email addresses to send the reports to
                //
                List<string> emailAddresses = new List<string>();
                emailAddresses = GetListOfEmailAddresses(xdoc, sqlConnectionString, rpParameters);
                foreach (string emailAddress in emailAddresses)
                {
                    RegexUtilities regexUtilities = new RegexUtilities();
                    validEmailAddress = regexUtilities.IsValidEmail(emailAddress);
                    if (validEmailAddress)
                    {


                        MailMessage mailMessage = new MailMessage();
                        mailMessage.To.Add(new MailAddress(emailAddress));
                        mailMessage.From = new MailAddress(smtpEmailSettings.FromAddress);
                        //mailMessage.From = new MailAddress("jcborland@jbsoftwareservices.onmicrosoft.com");
                        mailMessage.Subject = mailSubject;
                        mailMessage.Body = mailBody;
                        //mailMessage.Attachments.Add(new Attachment(file.FullName));
                        using (Attachment attachment = new Attachment(file.FullName))
                        {
                            mailMessage.Attachments.Add(attachment);

                            //emailPassword = "@LI20sserluss16:";

                            SmtpClient smtpClient = new SmtpClient();
                            smtpClient.UseDefaultCredentials = smtpEmailSettings.SMTPUserDefaultCredentials;
                            smtpClient.Credentials = new System.Net.NetworkCredential(smtpEmailSettings.SMTPUsername, smtpEmailSettings.SMTPPassword);

                            //smtpClient.Credentials = new System.Net.NetworkCredential("jcborland@jbsoftwareservices.onmicrosoft.com", "JB20soft14");
                            smtpClient.Port = smtpEmailSettings.SMTPPort;
                            smtpClient.Host = smtpEmailSettings.SMTPHost;
                            //smtpClient.Host = "outlook-emeawest4.office365.com";
                            smtpClient.EnableSsl = smtpEmailSettings.SMTPEnableSSL;

                            bool emailSent = false;
                            try
                            {


                                smtpClient.Send(mailMessage);
                                emailSent = true;


                            }
                            catch (Exception ex)
                            {
                                textLine = string.Format("Error sending an email to, {0}.\r\n{1}.\r\n", emailAddress, ex);
                                update_Progress(textLine, configDirName, logOneIn);
                            }

                            if (emailSent)
                            {


                            }
                            else
                            {

                            }
                        }



                    }
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error sending email for file, {0}.\r\n{1}.\r\n", file.FullName, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            finally
            {

            }

        }
        private SMTPEmailSettings GetEmailSettings(XDocument xdoc, string sqlConnectionString)
        {
            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);


            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;
            bool success = false;
            SMTPEmailSettings smtpEmailSettings = new SMTPEmailSettings();
            DataTable dtSMTPEmailSettings = new DataTable();
            //
            //Try using a stored procedure
            //
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectSMTPEmailSettings", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtSMTPEmailSettings);
                }
                success = true;
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting email settings with SQL connection string, {0}.\r\n{1}.\r\n", logConnectionString, ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            if (success)
            {
                //
                //There should only be one record.
                //
                DataRow drSMTPEmailSettings;
                drSMTPEmailSettings = dtSMTPEmailSettings.Rows[0];
                smtpEmailSettings.Body = null;                  //I'm not using this yet. May never use it.
                smtpEmailSettings.FromAddress = drSMTPEmailSettings.ItemArray[0].ToString();
                smtpEmailSettings.SMTPEnableSSL = Convert.ToBoolean(drSMTPEmailSettings.ItemArray[6]);
                smtpEmailSettings.SMTPHost = drSMTPEmailSettings.ItemArray[2].ToString();
                smtpEmailSettings.SMTPPassword = drSMTPEmailSettings.ItemArray[4].ToString();
                smtpEmailSettings.SMTPPort = Convert.ToInt32(drSMTPEmailSettings.ItemArray[5]);
                smtpEmailSettings.SMTPUserDefaultCredentials = Convert.ToBoolean(drSMTPEmailSettings.ItemArray[1]);
                smtpEmailSettings.SMTPUsername = drSMTPEmailSettings.ItemArray[3].ToString();
                smtpEmailSettings.Subject = null;               //I'm not using this yet. May never use it.

                string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;

                textLine = string.Format("Getting SMTP email settings with connection string : {0}.", logConnectionString);
                update_Progress(textLine, configDirName, logOneIn);

                textLine = string.Format("Got SMTP email settings, host is : {0}.", smtpEmailSettings.SMTPHost);
                update_Progress(textLine, configDirName, logOneIn);

            }

            return smtpEmailSettings;

        }
        private List<string> GetListOfEmailAddresses(XDocument xdoc, string sqlConnectionString, RPParameters rpParameters)
        {
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine = null;

            int x = sqlConnectionString.LastIndexOf(";Password=") + 10;
            int y = sqlConnectionString.LastIndexOf(";");
            string logConnectionString = sqlConnectionString.Substring(0, x + 2) + "*********" + sqlConnectionString.Substring(y - 2);

            textLine = string.Format("Start getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            List<string> emailAddresses = new List<string>();
            string companyNo = rpParameters.ErRef;                  //file.FullName.Substring(0, 4);
            DataTable dtEmailAddresses = new DataTable();
            try
            {
                using (var connection = new SqlConnection(sqlConnectionString))
                using (var command = new SqlCommand("SelectPayrollReportsContacts", connection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    connection.Open();
                    command.Parameters.AddWithValue("CompanyNo", companyNo);
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                    sqlDataAdapter.Fill(dtEmailAddresses);
                }
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error getting the list of email addresses.\r\n{0}.\r\n", ex);
                update_Progress(textLine, configDirName, logOneIn);
            }
            foreach (DataRow drEmailAddresses in dtEmailAddresses.Rows)
            {
                emailAddresses.Add(drEmailAddresses.ItemArray[0].ToString());
            }

            textLine = string.Format("Finished getting a list of email addresses with connection string : {0}.", logConnectionString);
            update_Progress(textLine, configDirName, logOneIn);

            return emailAddresses;
        }

        private void btnProduceReports_Click(object sender, EventArgs e)
        {
            string configDirName = "C:\\Payescape\\Service\\Config\\";
            //
            //Read the config file to get the outgoing folder and the timer details.
            //
            XDocument xdoc = new XDocument();
            string dirName = configDirName;
            ReadConfigFile configFile = new ReadConfigFile();
            xdoc = configFile.ConfigRecord(dirName);
            dirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            configDirName = dirName;
            int interval = Convert.ToInt32(xdoc.Root.Element("Interval").Value);
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);


            // Scan the folder and upload file waiting there.
            string textLine = string.Format("Starting from called program (ProcessPayRunIOOutput).");
            update_Progress(textLine, configDirName, 1);

            //Start by updating the contacts table
            UpdateContactDetails(xdoc);

            //Now process the reports
            ProcessReportsFromPayRunIO(xdoc);

            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnProduceReports.PerformClick();
        }
    }
    //public class ReadConfigFile
    //{
    //    //
    //    // Using XDocument instead of XmlReader
    //    //
    //    string fileName = "PayescapeWGtoPR.xml";
    //    string xmlSoftwareHomeFolder = "C:\\Payescape\\Service\\";
    //    string xmlDataHomeFolder = "C:\\Payescape\\Data\\";
    //    string xmlSFTPHostName = "sftp.bluemarblepayroll.com";
    //    string xmlUser = "payescape123";
    //    string xmlPasswordFile = "payescape.ppk";
    //    string xmlInterval = "10";
    //    string xmlLogOneIn = "100";
    //    string xmlOffFrom = "22:30:00";
    //    string xmlOffTo = "00:30:00";
    //    string xmlRunConstantly = "False";
    //    string xmlFilePrefix = "WGtoPR_";
    //    string xmlArchive = "True";
    //    string xmlDataSource = "APPSERVER1\\MSSQL";
    //    string xmlDatabase = "Payescape";
    //    string xmlUsername = "PayrollEngineLogin";
    //    string xmlPassword = "JB20soft14";
    //    XDocument xdoc = new XDocument();

    //    public ReadConfigFile() { }


    //    public XDocument ConfigRecord(string dirName)
    //    {
    //        string fullName = dirName + fileName;
    //        string passwordFile = dirName + xmlPasswordFile;

    //        try
    //        {
    //            bool updateRequired = false;
    //            bool exists = false;
    //            xdoc = XDocument.Load(fullName);
    //            exists = xdoc.Root.Descendants("SoftwareHomeFolder").Any();
    //            if (exists)
    //            {
    //                xmlSoftwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("DataHomeFolder").Any();
    //            if (exists)
    //            {
    //                xmlDataHomeFolder = xdoc.Root.Element("DataHomeFolder").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("OffFrom").Any();
    //            if (exists)
    //            {
    //                xmlOffFrom = xdoc.Root.Element("OffFrom").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("OffTo").Any();
    //            if (exists)
    //            {
    //                xmlOffTo = xdoc.Root.Element("OffTo").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("RunConstantly").Any();
    //            if (exists)
    //            {
    //                xmlRunConstantly = xdoc.Root.Element("RunConstantly").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("SFTPHostName").Any();
    //            if (exists)
    //            {
    //                xmlSFTPHostName = xdoc.Root.Element("SFTPHostName").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("User").Any();
    //            if (exists)
    //            {
    //                xmlUser = xdoc.Root.Element("User").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("PasswordFile").Any();
    //            if (exists)
    //            {
    //                xmlPasswordFile = xdoc.Root.Element("PasswordFile").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("Interval").Any();
    //            if (exists)
    //            {
    //                xmlInterval = xdoc.Root.Element("Interval").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("LogOneIn").Any();
    //            if (exists)
    //            {
    //                xmlLogOneIn = xdoc.Root.Element("LogOneIn").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("FilePrefix").Any();
    //            if (exists)
    //            {
    //                xmlFilePrefix = xdoc.Root.Element("FilePrefix").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            exists = xdoc.Root.Descendants("Archive").Any();
    //            if (exists)
    //            {
    //                xmlArchive = xdoc.Root.Element("Archive").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            if (updateRequired)
    //            {
    //                CreateConfigFile(dirName, fullName);
    //            }
    //            exists = xdoc.Root.Descendants("DataSource").Any();
    //            if (exists)
    //            {
    //                xmlDataSource = xdoc.Root.Element("DataSource").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            if (updateRequired)
    //            {
    //                CreateConfigFile(dirName, fullName);
    //            }
    //            exists = xdoc.Root.Descendants("Database").Any();
    //            if (exists)
    //            {
    //                xmlDatabase = xdoc.Root.Element("Database").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            if (updateRequired)
    //            {
    //                CreateConfigFile(dirName, fullName);
    //            }
    //            exists = xdoc.Root.Descendants("Username").Any();
    //            if (exists)
    //            {
    //                xmlUsername = xdoc.Root.Element("Username").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            if (updateRequired)
    //            {
    //                CreateConfigFile(dirName, fullName);
    //            }
    //            exists = xdoc.Root.Descendants("Password").Any();
    //            if (exists)
    //            {
    //                xmlPassword = xdoc.Root.Element("Password").Value;
    //            }
    //            else
    //            {
    //                updateRequired = true;
    //            }
    //            if (updateRequired)
    //            {
    //                CreateConfigFile(dirName, fullName);
    //            }
    //        }

    //        catch (Exception ex)
    //        {

    //            if (ex.ToString().Contains("Could not find a part of the path") || ex.ToString().Contains("Could not find file"))
    //            {
    //                CreateConfigFile(dirName, fullName);

    //            }
    //            xdoc = XDocument.Load(fullName);
    //        }

    //        return xdoc;
    //    }
    //    private void CreateConfigFile(string dirName, string fullName)
    //    {
    //        // Create Folder and dummy config xml file.
    //        Directory.CreateDirectory(dirName);


    //        // Create a dummy config xml file.
    //        new XDocument
    //            (
    //            new XElement
    //                ("Configuration",
    //                 new XElement("SoftwareHomeFolder", xmlSoftwareHomeFolder),
    //                 new XElement("DataHomeFolder", xmlDataHomeFolder),
    //                 new XElement("Interval", xmlInterval),
    //                 new XElement("LogOneIn", xmlLogOneIn),
    //                 new XElement("RunConstantly", xmlRunConstantly),
    //                 new XElement("OffFrom", xmlOffFrom),
    //                 new XElement("OffTo", xmlOffTo),
    //                 new XElement("SFTPHostName", xmlSFTPHostName),
    //                 new XElement("User", xmlUser),
    //                 new XElement("PasswordFile", xmlPasswordFile),
    //                 new XElement("FilePrefix", xmlFilePrefix),
    //                 new XElement("Archive", xmlArchive),
    //                 new XElement("DataSource", xmlDataSource),
    //                 new XElement("Database", xmlDatabase),
    //                 new XElement("Username", xmlUsername),
    //                 new XElement("Password", xmlPassword)
    //                )
    //               )
    //         .Save(fullName);
    //        xdoc = XDocument.Load(fullName);
    //    }

    //}
    ////Classes for printing of the standard reports

    ////Report (RP) Parameters
    //public class RPParameters
    //{
    //    public string ErRef { get; set; }
    //    public int TaxYear { get; set; }
    //    public DateTime AccYearStart { get; set; }
    //    public DateTime AccYearEnd { get; set; }
    //    public int TaxPeriod { get; set; }
    //    public string PaySchedule { get; set; }
    //    public DateTime PayRunDate { get; set; }


    //    public RPParameters() { }
    //    public RPParameters(string erRef, int taxYear, DateTime accYearStart,
    //                        DateTime accYearEnd, int taxPeriod, string paySchedule, DateTime payRundate)
    //    {
    //        ErRef = erRef;
    //        TaxYear = taxYear;
    //        AccYearStart = accYearStart;
    //        AccYearEnd = accYearEnd;
    //        TaxPeriod = taxPeriod;
    //        PaySchedule = paySchedule;
    //        PayRunDate = payRundate;
    //    }
    //}
    ////Report (RP) Employer
    //public class RPEmployer
    //{
    //    public string Name { get; set; }
    //    public string PayeRef { get; set; }

    //    public RPEmployer() { }
    //    public RPEmployer(string name, string payeRef)
    //    {
    //        Name = name;
    //        PayeRef = payeRef;


    //    }
    //}

    ////Report (RP) Employee
    //public class RPEmployeePeriod
    //{
    //    public string Reference { get; set; }
    //    public string Title { get; set; }
    //    public string Forename { get; set; }
    //    public string Surname { get; set; }
    //    public string Fullname { get; set; }
    //    public string RefFullname { get; set; }
    //    public string Address1 { get; set; }
    //    public string Address2 { get; set; }
    //    public string Address3 { get; set; }
    //    public string Address4 { get; set; }
    //    public string Postcode { get; set; }
    //    public string Country { get; set; }
    //    public string SortCode { get; set; }
    //    public string BankAccNo { get; set; }
    //    public DateTime DateOfBirth { get; set; }
    //    public string Gender { get; set; }
    //    public string BuildingSocRef { get; set; }
    //    public string NINumber { get; set; }
    //    public string PaymentMethod { get; set; }
    //    public DateTime PayRunDate { get; set; }
    //    public DateTime PeriodStartDate { get; set; }
    //    public DateTime PeriodEndDate { get; set; }
    //    public int PayrollYear { get; set; }
    //    public decimal Gross { get; set; }
    //    public decimal NetPayTP { get; set; }
    //    public int DayHours { get; set; }
    //    public DateTime StudentLoanStartDate { get; set; }
    //    public DateTime StudentLoanEndDate { get; set; }
    //    public decimal StudentLoan { get; set; }
    //    public string NILetter { get; set; }
    //    public string CalculationBasis { get; set; }
    //    public decimal TotalPayTP { get; set; }
    //    public decimal EarningsToLEL { get; set; }
    //    public decimal EarningsToSET { get; set; }
    //    public decimal EarningsToPET { get; set; }
    //    public decimal EarningsToUST { get; set; }
    //    public decimal EarningsToAUST { get; set; }
    //    public decimal EarningsToUEL { get; set; }
    //    public decimal EarningsAboveUEL { get; set; }
    //    public decimal EeContributionsPt1 { get; set; }
    //    public decimal EeContributionsPt2 { get; set; }
    //    public decimal ErNICYTD { get; set; }
    //    public decimal EeRebate { get; set; }
    //    public decimal ErRebate { get; set; }
    //    public decimal EeReduction { get; set; }
    //    public DateTime LeavingDate { get; set; }
    //    public bool Leaver { get; set; }
    //    public string TaxCode { get; set; }
    //    public bool Week1Month1 { get; set; }
    //    public string TaxCodeChangeTypeID { get; set; }
    //    public string TaxCodeChangeType { get; set; }
    //    public decimal TaxPrev { get; set; }
    //    public decimal TaxablePayPrevious { get; set; }
    //    public decimal TaxThis { get; set; }
    //    public decimal TaxablePayYTD { get; set; }
    //    public decimal TaxablePayTP { get; set; }
    //    public decimal HolidayAccruedYTD { get; set; }
    //    public decimal ErPensionYTD { get; set; }
    //    public decimal EePensionYTD { get; set; }
    //    public decimal ErPensionTP { get; set; }
    //    public decimal EePensionTP { get; set; }
    //    public decimal ErContributionPercent { get; set; }
    //    public decimal EeContributionPercent { get; set; }
    //    public decimal PensionablePay { get; set; }
    //    public DateTime ErPensionPayRunDate { get; set; }
    //    public DateTime EePensionPayRunDate { get; set; }
    //    public DateTime DirectorshipAppointmentDate { get; set; }
    //    public bool Director { get; set; }
    //    public decimal EeContributionTaxPeriodPt1 { get; set; }
    //    public decimal EeContributionTaxPeriodPt2 { get; set; }
    //    public decimal ErNICTP { get; set; }
    //    public string Frequency { get; set; }
    //    public decimal NetPayYTD { get; set; }
    //    public decimal TotalPayYTD { get; set; }
    //    public decimal TotalDedTP { get; set; }
    //    public decimal TotalDedYTD { get; set; }
    //    public string PensionCode { get; set; }
    //    public decimal PreTaxAddDed { get; set; }
    //    public decimal GUCosts { get; set; }
    //    public decimal AbsencePay { get; set; }
    //    public decimal HolidayPay { get; set; }
    //    public decimal PreTaxPension { get; set; }
    //    public decimal Tax { get; set; }
    //    public decimal NetNI { get; set; }
    //    public decimal PostTaxAddDed { get; set; }
    //    public decimal PostTaxPension { get; set; }
    //    public decimal AOE { get; set; }
    //    public List<RPAddition> Additions { get; set; }
    //    public List<RPDeduction> Deductions { get; set; }
    //    public RPEmployeePeriod() { }
    //    public RPEmployeePeriod(string reference, string title, string forename, string surname, string fullname, string refFullname,
    //                      string address1, string address2, string address3, string address4, string postcode,
    //                      string country, DateTime dateOfBirth, string gender, bool leaver, DateTime leavingDate,
    //                      string niNumber, string niLetter, string taxCode, bool week1Month1, string frequency,
    //                      string paymentMethod, DateTime payRunDate,
    //                      decimal netPayTP, decimal netPayYTD, decimal taxablePayTP, decimal taxablePayYTD,
    //                      decimal taxablePayPrevious, decimal totalPayTP, decimal totalPayYTD, decimal totalDedTP, decimal totalDedYTD,
    //                      decimal erNICTP, decimal erNICYTD, decimal erPensionTP, decimal eePensionTP, decimal erPensionYTD,
    //                      decimal eePensionYTD, decimal pensionablePay, string pensionCode, string sortCode, string bankAccNo, string buildingSocRef,
    //                      decimal erContributionPercent, decimal preTaxAddDed, decimal guCosts, decimal absencePay,
    //                      decimal holidayPay, decimal preTaxPension, decimal tax, decimal taxPrev, decimal taxThis, decimal netNI,
    //                      decimal postTaxAddDed, decimal postTaxPension, decimal aoe, decimal studentLoan,
    //                      decimal eeContributionPercent, List<RPAddition> additions, List<RPDeduction> deductions)
    //    {
    //        Reference = reference;
    //        Title = title;
    //        Forename = forename;
    //        Surname = surname;
    //        Fullname = fullname;
    //        RefFullname = refFullname;
    //        Address1 = address1;
    //        Address2 = address2;
    //        Address3 = address3;
    //        Address4 = address4;
    //        Postcode = postcode;
    //        Country = country;
    //        SortCode = sortCode;
    //        BankAccNo = bankAccNo;
    //        DateOfBirth = dateOfBirth;
    //        Gender = gender;
    //        BuildingSocRef = buildingSocRef;
    //        NINumber = niNumber;
    //        PaymentMethod = paymentMethod;
    //        PayRunDate = PayRunDate;
    //        //PeriodStartDate
    //        //PeriodEndDate
    //        //PayrollYear
    //        //Gross
    //        NetPayTP = netPayTP;
    //        //DayHours
    //        //StudentLoanStartDate
    //        //StundentLoanEndDate
    //        StudentLoan = studentLoan;
    //        NILetter = niLetter;
    //        //CalculationBasis
    //        TotalPayTP = totalPayTP;
    //        //EarningsToLEL
    //        //EarningsToSET
    //        //EarningsToPET
    //        //EarningsToUST
    //        //EarningsToAUST
    //        //EarningsToUEL
    //        //EarningsAboveUel
    //        //EeContributionsPt1
    //        //EeContributionsPt2
    //        ErNICYTD = erNICYTD;
    //        //EeRebate
    //        //ErRebate
    //        //EeReduction
    //        LeavingDate = leavingDate;
    //        Leaver = leaver;
    //        TaxCode = taxCode;
    //        Week1Month1 = week1Month1;
    //        //TaxCodeChangeTypeID
    //        //TaxCodeChangeType
    //        TaxPrev = taxPrev;
    //        TaxablePayPrevious = taxablePayPrevious;
    //        TaxThis = taxThis;
    //        TaxablePayYTD = taxablePayYTD;
    //        TaxablePayTP = taxablePayTP;
    //        //Holiday AccruedTd
    //        ErPensionYTD = erPensionYTD;
    //        EePensionYTD = eePensionYTD;
    //        ErPensionTP = erPensionTP;
    //        EePensionTP = eePensionTP;
    //        ErContributionPercent = erContributionPercent;
    //        EeContributionPercent = eeContributionPercent;
    //        PensionablePay = pensionablePay;
    //        //ErPensionPayRunDate
    //        //EePensionPayRunDate
    //        //DirectorshipAppointmentDate
    //        //Director
    //        //EeContributionsTaxPeriodPt1
    //        //EeContributionsTaxPeriodPt2
    //        ErNICTP = erNICTP;
    //        Frequency = frequency;
    //        NetPayYTD = netPayYTD;
    //        TotalPayTP = totalPayTP;
    //        TotalPayYTD = totalPayYTD;
    //        TotalDedTP = totalDedTP;
    //        TotalDedYTD = totalDedYTD;
    //        PensionCode = pensionCode;
    //        PreTaxAddDed = preTaxAddDed;
    //        GUCosts = guCosts;
    //        AbsencePay = absencePay;
    //        HolidayPay = holidayPay;
    //        PreTaxPension = preTaxPension;
    //        Tax = tax;
    //        NetNI = netNI;
    //        PostTaxAddDed = postTaxAddDed;
    //        PostTaxPension = postTaxPension;
    //        AOE = aoe;
    //        Additions = additions;
    //        Deductions = deductions;
    //    }

    //}
    //public class P45
    //{
    //    public string ErOfficeNo { get; set; }
    //    public string ErRefNo { get; set; }
    //    public string NINumber { get; set; }
    //    public string Title { get; set; }
    //    public string Surname { get; set; }
    //    public string FirstNames { get; set; }
    //    public DateTime LeavingDate { get; set; }
    //    public bool StudentLoansDeductionToContinue { get; set; }
    //    public string TaxCode { get; set; }
    //    public bool Week1Month1 { get; set; }
    //    public int WeekNo { get; set; }
    //    public int MonthNo { get; set; }
    //    public decimal PayToDate { get; set; }
    //    public decimal TaxToDate { get; set; }
    //    public decimal PayThis { get; set; }
    //    public decimal TaxThis { get; set; }
    //    public string EeRef { get; set; }
    //    public bool IsMale { get; set; }
    //    public DateTime DateOfBirth { get; set; }
    //    public string Address1 { get; set; }
    //    public string Address2 { get; set; }
    //    public string Address3 { get; set; }
    //    public string Address4 { get; set; }
    //    public string Postcode { get; set; }
    //    public string Country { get; set; }
    //    public string ErName { get; set; }
    //    public string ErAddress1 { get; set; }
    //    public string ErAddress2 { get; set; }
    //    public string ErAddress3 { get; set; }
    //    public string ErAddress4 { get; set; }
    //    public string ErPostcode { get; set; }
    //    public string ErCountry { get; set; }
    //    public DateTime Now { get; set; }

    //    public P45() { }
    //    public P45(string erOfficeNo, string erRefNo, string niNumber, string title, string surname, string firstNames,
    //                DateTime leavingDate,
    //                bool studentLoansDedustionToContinue, string taxCode, int weekNo, int monthNo,
    //                decimal payToDate, decimal taxToDate, decimal payThis, decimal taxThis, string eeRef, bool isMale,
    //                string erName, string address1,
    //                string address2, string address3, string address4, string postcode, string country,
    //                DateTime dateOfBirth, string erAddress1,
    //                string erAddress2, string erAddress3, string erAddress4, string erPostcode, string erCountry,
    //                DateTime now)


    //    {
    //        ErOfficeNo = erOfficeNo;
    //        ErRefNo = erRefNo;
    //        NINumber = niNumber;
    //        Title = title;
    //        Surname = surname;
    //        FirstNames = firstNames;
    //        LeavingDate = leavingDate;
    //        StudentLoansDeductionToContinue = studentLoansDedustionToContinue;
    //        TaxCode = taxCode;
    //        WeekNo = weekNo;
    //        MonthNo = monthNo;
    //        PayToDate = payToDate;
    //        TaxToDate = taxToDate;
    //        PayThis = payThis;
    //        TaxThis = TaxThis;
    //        EeRef = eeRef;
    //        IsMale = isMale;
    //        DateOfBirth = dateOfBirth;
    //        Address1 = address1;
    //        Address2 = address2;
    //        Address3 = address3;
    //        Address4 = address4;
    //        Postcode = postcode;
    //        Country = country;
    //        ErName = erName;
    //        ErAddress1 = erAddress1;
    //        ErAddress2 = erAddress2;
    //        ErAddress3 = erAddress3;
    //        ErAddress4 = erAddress4;
    //        ErPostcode = erPostcode;
    //        ErCountry = erCountry;
    //        Now = now;
    //    }

    //}

    ////Report (RP) Additions
    //public class RPAddition
    //{
    //    public string EeRef { get; set; }
    //    public string Description { get; set; }
    //    public decimal Rate { get; set; }
    //    public decimal Units { get; set; }
    //    public decimal AmountTP { get; set; }
    //    public decimal AmountYTD { get; set; }
    //    public RPAddition() { }
    //    public RPAddition(string eeRef, string description, decimal rate, decimal units,
    //                       decimal amountTP, decimal amountYTD)
    //    {
    //        EeRef = eeRef;
    //        Description = description;
    //        Rate = rate;
    //        Units = units;
    //        AmountTP = amountTP;
    //        AmountYTD = amountYTD;

    //    }
    //}

    ////Report (RP) Deductions
    //public class RPDeduction
    //{
    //    public string EeRef { get; set; }
    //    public string Description { get; set; }
    //    public decimal AmountTP { get; set; }
    //    public decimal AmountYTD { get; set; }
    //    public RPDeduction() { }
    //    public RPDeduction(string eeRef, string description,
    //                       decimal amountTP, decimal amountYTD)
    //    {
    //        EeRef = eeRef;
    //        Description = description;
    //        AmountTP = amountTP;
    //        AmountYTD = amountYTD;

    //    }
    //}
    //public class RPPayComponent
    //{
    //    public string PayCode { get; set; }
    //    public string Description { get; set; }
    //    public string EeRef { get; set; }
    //    public string Fullname { get; set; }
    //    public string Surname { get; set; }
    //    public decimal Rate { get; set; }
    //    public decimal UnitsTP { get; set; }
    //    public decimal AmountTP { get; set; }
    //    public decimal UnitsYTD { get; set; }
    //    public decimal AmountYTD { get; set; }
    //    public RPPayComponent() { }
    //    public RPPayComponent(string payCode, string description, string eeRef, string fullname,
    //                          string surname, decimal rate, decimal unitsTP, decimal amountTP,
    //                           decimal unitsYTD, decimal amountYTD)
    //    {
    //        PayCode = payCode;
    //        Description = description;
    //        EeRef = eeRef;
    //        Fullname = fullname;
    //        Surname = surname;
    //        Rate = rate;
    //        UnitsTP = unitsTP;
    //        AmountTP = amountTP;
    //        UnitsYTD = unitsYTD;
    //        AmountYTD = amountYTD;

    //    }
    //}
    //public class SMTPEmailSettings
    //{
    //    public string Subject { get; set; }
    //    public string Body { get; set; }
    //    public string FromAddress { get; set; }
    //    public bool SMTPUserDefaultCredentials { get; set; }
    //    public string SMTPUsername { get; set; }
    //    public string SMTPPassword { get; set; }
    //    public int SMTPPort { get; set; }
    //    public string SMTPHost { get; set; }
    //    public bool SMTPEnableSSL { get; set; }
    //    public SMTPEmailSettings() { }
    //    public SMTPEmailSettings(string subject, string body, string fromAddress, bool smtpUserDefaultCredentials,
    //                             string smtpUsername, string smtpPassword, int smtpPort, string smtpHost,
    //                             bool smtpEnableSSL)
    //    {
    //        Subject = subject;
    //        Body = body;
    //        FromAddress = fromAddress;
    //        SMTPUserDefaultCredentials = smtpUserDefaultCredentials;
    //        SMTPUsername = smtpUsername;
    //        SMTPPassword = smtpPassword;
    //        SMTPPort = smtpPort;
    //        SMTPHost = smtpHost;
    //        SMTPEnableSSL = smtpEnableSSL;
    //    }
    //}
    //public class RegexUtilities
    //{
    //    bool invalid = false;

    //    public bool IsValidEmail(string strIn)
    //    {
    //        invalid = false;
    //        if (String.IsNullOrEmpty(strIn))
    //            return false;

    //        // Use IdnMapping class to convert Unicode domain names.
    //        strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper);
    //        if (invalid)
    //            return false;

    //        // Return true if strIn is in valid e-mail format.
    //        return Regex.IsMatch(strIn,
    //               @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
    //               @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
    //               RegexOptions.IgnoreCase);
    //    }
    //    public bool IsValidPostcode(string strIn)
    //    {
    //        invalid = false;
    //        if (String.IsNullOrEmpty(strIn))
    //            return false;

    //        // Return true if strIn is in valid Postcode format.
    //        return Regex.IsMatch(strIn,
    //               "(^gir\\s?0aa$)|(^[a-z-[qvx]](\\d{1,2}|[a-hk-y]\\d{1,2}|\\d[a-hjks-uw]|[a-hk-y]\\d[abehmnprv-y])\\s?\\d[a-z-[cikmov]]{2}$)",
    //               RegexOptions.IgnoreCase);
    //    }
    //    public bool IsValidNINumber(string strIn)
    //    {
    //        invalid = false;
    //        if (String.IsNullOrEmpty(strIn))
    //            return false;

    //        // Return true if strIn is in valid NI Number format.
    //        return Regex.IsMatch(strIn,
    //               @"^([a-zA-Z]){2}( )?([0-9]){2}( )?([0-9]){2}( )?([0-9]){2}( )?([a-zA-Z]){1}?$",
    //               RegexOptions.IgnoreCase);
    //    }
    //    private string DomainMapper(Match match)
    //    {
    //        // IdnMapping class with default property values.
    //        IdnMapping idn = new IdnMapping();

    //        string domainName = match.Groups[2].Value;
    //        try
    //        {
    //            domainName = idn.GetAscii(domainName);
    //        }
    //        catch (ArgumentException)
    //        {
    //            invalid = true;
    //        }
    //        return match.Groups[1].Value + domainName;

    //    }
    //}
}
