using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using PicoXLSX;
using PayRunIOClassLibrary;
using System.Globalization;
using System.Reflection;
using WinSCP;
using System.Text.RegularExpressions;

namespace PayRunIOProcessReports
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
           InitializeComponent();
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

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            
            // Scan the folder and upload file waiting there.
            string textLine = string.Format("Starting from called program (PayRunIOProcessReports).");
            prWG.update_Progress(textLine, configDirName, 1);
            
            
            //Start by updating the contacts table
            prWG.UpdateContactDetails(xdoc);
            
            //Now process the reports
            ProcessReportsFromPayRunIO(xdoc);


            Close();

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

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            textLine = string.Format("Start processing the reports.");
            prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            string originalDirName = "Outputs";
            string archiveDirName = "PE-ArchivedOutputs";
            string[] directories = prWG.GetAListOfDirectories(xdoc, "Outputs");

            for (int i = 0; i < directories.Count(); i++)
            {
                try
                {
                    bool success = ProduceReports(xdoc, directories[i]);
                    if (success)
                    {
                        prWG.ArchiveDirectory(xdoc, directories[i], originalDirName, archiveDirName);
                    }


                }
                catch (Exception ex)
                {
                    textLine = string.Format("Error processing the reports for directory {0}.\r\n{1}.\r\n", directories[i], ex);
                    prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                }

            }

            textLine = string.Format("Finished processing the reports.");
            prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);

            //Transfer the contents of PE-Outgoing up to the Blue Marble SFTP server
            //Each company has it's own folder beneath PE-Outgoing which is just named with their company number.
            textLine = string.Format("Start processing the PE-Outgoing directory.");
            prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);

            originalDirName = "PE-Outgoing";
            archiveDirName = "PE-Outgoing\\Archive";
            directories = prWG.GetAListOfDirectories(xdoc, "PE-Outgoing");

            for (int i = 0; i < directories.Count(); i++)
            {
                if(!directories[i].Contains("Archive"))
                {
                    try
                    {
                        bool success = TransferToBlueMarbleSFTPServer(xdoc, directories[i]);            // Transfer contents of the folder up to Blue Marble SFTP server.//ProduceReports(xdoc, directories[i]);
                        if (success)
                        {
                            try
                            {
                                textLine = string.Format("Trying to archive directory {0}.", directories[i]);
                                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);

                                prWG.ArchiveDirectory(xdoc, directories[i], originalDirName, archiveDirName);
                            }
                            catch(Exception ex)
                            {
                                textLine = string.Format("Error archiving directory {0}.\r\n{1}.\r\n", directories[i], ex);
                                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                            }
                            
                        }


                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error processing PE-Outgoing folder for directory {0}.\r\n{1}.\r\n", directories[i], ex);
                        prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                    }

                }

            }

        }
        private bool TransferToBlueMarbleSFTPServer(XDocument xdoc, string directory)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string dataHomeFolder = xdoc.Root.Element("DataHomeFolder").Value;
            bool archive = Convert.ToBoolean(xdoc.Root.Element("Archive").Value);
            string sftpHostName = xdoc.Root.Element("SFTPHostName").Value;
            string user = xdoc.Root.Element("User").Value;
            bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);
            if(!live)
            {
                user = "payruntest123";//For testing purposes
            }
            
            string passwordFile = softwareHomeFolder + "Programs\\" +xdoc.Root.Element("PasswordFile").Value;
            string filePrefix = xdoc.Root.Element("FilePrefix").Value;
            int interval = Convert.ToInt32(xdoc.Root.Element("Interval").Value);
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);

            string textLine = null;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            textLine = string.Format("Start processing the reports.");
            prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);

            string peOutgoingDirName = dataHomeFolder + "\\PE-Outgoing";
            bool success = true;

            //
            // Don't think Blue Marble can cope with a directory so go through each file individually
            //
            bool isUnity = true;

            string dataSource = xdoc.Root.Element("DataSource").Value;            //"APPSERVER1\\MSSQL";  //"13.69.154.210\\MSSQL";  
            string dataBase = xdoc.Root.Element("Database").Value;
            string userID = xdoc.Root.Element("Username").Value;
            string password = xdoc.Root.Element("Password").Value;
            string sqlConnectionString = "Server=" + dataSource + ";Database=" + dataBase + ";User ID=" + userID + ";Password=" + password + ";";
            int x = directory.LastIndexOf('\\') + 1;
            int companyNo = Convert.ToInt32(directory.Substring(x, 4));

            foreach (var csvFile in Directory.GetFiles(directory))
            {
                // Use SFTP to send the file automatically.
                try
                {
                    PutToSFTP PutToSFTP = new PutToSFTP();

                    string[] sftpReturn = PutToSFTP.SFTPTransfer(csvFile, sftpHostName, user, passwordFile, isUnity);
                    if (sftpReturn[0] == "Success")
                    {
                        textLine = string.Format("Successfully uploaded csv file {0} for client reference : {1}", csvFile, companyNo);
                        prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                    }
                    else
                    {
                        //
                        // SFTP failed
                        //
                        textLine = sftpReturn[1];
                        prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                        success = false;
                    }

                }
                catch
                {
                    textLine = string.Format("Failed to upload zipped csv file {0} for client reference : {1}", csvFile, companyNo);
                    prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
                    success = false;
                }
            }
            return success;
        }
       
        
        private void ProducePeriodReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer,
                                          List<P45> p45s, List<RPPayComponent> rpPayComponents, RPParameters rpParameters,
                                          List<RPPreSamplePayCode> rpPreSamplePayCodes, List<RPPensionContribution> rpPensionContributions)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);

            string textLine = null;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            
            //I now have a list of employee with their total for this period & ytd plus addition & deductions
            //I can print payslips and standard reports from here.
            try
            {
                prWG.PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents, rpPensionContributions);
                if (rpEmployer.P32Required)
                {
                    RPP32Report rpP32Report = CreateP32Report(xdoc, rpEmployer, rpParameters);
                    prWG.PrintP32Report(xdoc, rpP32Report, rpParameters);

                    //var payeMonth = rpParameters.AccYearEnd.Day < 6 ? rpParameters.AccYearEnd.Month - 4 : rpParameters.AccYearEnd.Month - 3;
                    int payeMonth = rpParameters.PayRunDate.Day < 6 ? rpParameters.PayRunDate.Month - 4 : rpParameters.PayRunDate.Month - 3;
                    if (payeMonth <= 0)
                    {
                        payeMonth += 12;
                    }

                    //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
                    decimal hmrcTotal = prWG.CalculateHMRCTotal(rpP32Report, payeMonth);
                    rpEmployer.HMRCDesc = "[" + hmrcTotal.ToString() + "]";
                }
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error printing standard reports.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }
            
            //Produce bank files if necessary
            try
            {
                prWG.ProcessBankAndPensionFiles(xdoc, rpEmployeePeriodList, rpPensionContributions, rpEmployer, rpParameters);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error processing bank reports.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }

            
            //Produce Pre Sample file (XLSX)
            CreatePreSampleXLSX(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, rpPreSamplePayCodes);
            try
            {
                prWG.ZipReports(xdoc, rpEmployer, rpParameters);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error zipping reports.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }
            try
            {
                prWG.EmailZippedReports(xdoc, rpEmployer, rpParameters);
                
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error emailing zipped reports.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }
            try
            {
                prWG.UploadZippedReportsToAmazonS3(xdoc, rpEmployer, rpParameters);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error uploading zipped reports to Amazon S3.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }

        }
        private Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                           List<RPPensionContribution>, RPEmployer, RPParameters> 
                           PrepareStandardReports(XDocument xdoc, XmlDocument xmlReport, RPParameters rpParameters)
        {
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            
            List<RPEmployeePeriod> rpEmployeePeriodList = new List<RPEmployeePeriod>();
            List<P45> p45s = new List<P45>();
            //Create a list of Pay Code totals for the Payroll Component Analysis report
            List<RPPayComponent> rpPayComponents = new List<RPPayComponent>();
            RPEmployer rpEmployer = prWG.GetRPEmployer(xmlReport);
            //Create a list of all possible Pay Codes just from the first employee
            bool preSamplePayCodes = false;
            List<RPPreSamplePayCode> rpPreSamplePayCodes = new List<RPPreSamplePayCode>();
            List<RPPensionContribution> rpPensionContributions = new List<RPPensionContribution>();

            try
            {
                bool gotPayRunDate = false;
                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;

                    string payRunDate = prWG.GetElementByTagFromXml(employee, "PayRunDate");

                    if (payRunDate != "No Pay Run Data Found" && payRunDate != null)
                    {
                        if (!gotPayRunDate)
                        {
                            rpParameters.PayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PayRunDate"));
                            gotPayRunDate = true;
                            
                        }
                        decimal gross = prWG.GetDecimalElementByTagFromXml(employee, "Gross");
                        decimal net = prWG.GetDecimalElementByTagFromXml(employee, "Net");
                        //If the employee is a leaver before the start date then don't include unless they have a gross or net.
                        string leaver = prWG.GetElementByTagFromXml(employee, "Leaver");
                        DateTime leavingDate = new DateTime();
                        if (prWG.GetElementByTagFromXml(employee, "LeavingDate") != "")
                        {
                            leavingDate = DateTime.ParseExact(prWG.GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                        }
                        DateTime periodStartDate = DateTime.ParseExact(prWG.GetElementByTagFromXml(employee, "PeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                        if (leaver.StartsWith("N"))
                        {
                            include = true;
                        }
                        else if (leavingDate >= periodStartDate)
                        {
                            include = true;
                        }
                        else if(gross != 0 || net != 0)
                        {
                            include = true;
                        }

                    }

                    if (include)
                    {
                        RPEmployeePeriod rpEmployeePeriod = new RPEmployeePeriod();
                        rpEmployeePeriod.Reference = prWG.GetElementByTagFromXml(employee, "EeRef");
                        rpEmployeePeriod.Title = prWG.GetElementByTagFromXml(employee, "Title");
                        rpEmployeePeriod.Forename = prWG.GetElementByTagFromXml(employee, "FirstName");
                        rpEmployeePeriod.Surname = prWG.GetElementByTagFromXml(employee, "LastName");
                        rpEmployeePeriod.Fullname = rpEmployeePeriod.Title + " " + rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                        rpEmployeePeriod.RefFullname = rpEmployeePeriod.Reference + " " + rpEmployeePeriod.Fullname;
                        rpEmployeePeriod.SurnameForename = rpEmployeePeriod.Surname + rpEmployeePeriod.Forename;
                        string[] address = new string[6];
                        address[0] = prWG.GetElementByTagFromXml(employee, "Address1");
                        address[1] = prWG.GetElementByTagFromXml(employee, "Address2");
                        address[2] = prWG.GetElementByTagFromXml(employee, "Address3");
                        address[3] = prWG.GetElementByTagFromXml(employee, "Address4");
                        address[4] = prWG.GetElementByTagFromXml(employee, "Postcode");
                        address[5] = prWG.GetElementByTagFromXml(employee, "Country");

                        rpEmployeePeriod.SortCode = prWG.GetElementByTagFromXml(employee, "SortCode");
                        rpEmployeePeriod.BankAccNo = prWG.GetElementByTagFromXml(employee, "BankAccNo");
                        rpEmployeePeriod.DateOfBirth = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "DateOfBirth"));
                        rpEmployeePeriod.Gender = prWG.GetElementByTagFromXml(employee, "Gender");
                        rpEmployeePeriod.BuildingSocRef = prWG.GetElementByTagFromXml(employee, "BuildingSocRef");
                        rpEmployeePeriod.NINumber = prWG.GetElementByTagFromXml(employee, "NiNumber");
                        rpEmployeePeriod.PaymentMethod = prWG.GetElementByTagFromXml(employee, "PayMethod");
                        rpEmployeePeriod.PayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PayRunDate"));
                        rpEmployeePeriod.PeriodStartDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PeriodStartDate"));
                        rpEmployeePeriod.PeriodEndDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PeriodEndDate"));
                        rpEmployeePeriod.PayrollYear = prWG.GetIntElementByTagFromXml(employee, "PayrollYear");
                        rpEmployeePeriod.Gross = prWG.GetDecimalElementByTagFromXml(employee, "Gross");
                        rpEmployeePeriod.NetPayTP = prWG.GetDecimalElementByTagFromXml(employee, "Net");
                        rpEmployeePeriod.DayHours = prWG.GetIntElementByTagFromXml(employee, "DayHours");
                        rpEmployeePeriod.StudentLoanStartDate = prWG.GetDateElementByTagFromXml(employee, "StudentLoanStartDate");
                        rpEmployeePeriod.StudentLoanEndDate = prWG.GetDateElementByTagFromXml(employee, "StudentLoanEndDate");
                        rpEmployeePeriod.NILetter = prWG.GetElementByTagFromXml(employee, "NiLetter");
                        rpEmployeePeriod.CalculationBasis = prWG.GetElementByTagFromXml(employee, "CalculationBasis");
                        rpEmployeePeriod.Total = prWG.GetDecimalElementByTagFromXml(employee, "Total");
                        rpEmployeePeriod.EarningsToLEL = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToLEL");
                        rpEmployeePeriod.EarningsToSET = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToSET");
                        rpEmployeePeriod.EarningsToPET = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToPET");
                        rpEmployeePeriod.EarningsToUST = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToUST");
                        rpEmployeePeriod.EarningsToAUST = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToAUST");
                        rpEmployeePeriod.EarningsToUEL = prWG.GetDecimalElementByTagFromXml(employee, "EarningsToUEL");
                        rpEmployeePeriod.EarningsAboveUEL = prWG.GetDecimalElementByTagFromXml(employee, "EarningsAboveUEL");
                        rpEmployeePeriod.EeContributionsPt1 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionsPt1");
                        rpEmployeePeriod.EeContributionsPt2 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributions2");
                        rpEmployeePeriod.ErNICYTD = prWG.GetDecimalElementByTagFromXml(employee, "ErContributions");
                        rpEmployeePeriod.EeRebate = prWG.GetDecimalElementByTagFromXml(employee, "EeRabate");
                        rpEmployeePeriod.ErRebate = prWG.GetDecimalElementByTagFromXml(employee, "ErRebate");
                        rpEmployeePeriod.EeReduction = prWG.GetDecimalElementByTagFromXml(employee, "EeReduction");
                        string leaver = prWG.GetElementByTagFromXml(employee, "Leaver");
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
                            rpEmployeePeriod.LeavingDate = prWG.GetDateElementByTagFromXml(employee, "LeavingDate");

                        }
                        else
                        {
                            rpEmployeePeriod.LeavingDate = null;
                        }
                        rpEmployeePeriod.TaxCode = prWG.GetElementByTagFromXml(employee, "TaxCode");
                        rpEmployeePeriod.Week1Month1 = prWG.GetBooleanElementByTagFromXml(employee, "Week1Month1");
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            rpEmployeePeriod.TaxCode = rpEmployeePeriod.TaxCode + " W1";
                        }
                        rpEmployeePeriod.TaxCodeChangeTypeID = prWG.GetElementByTagFromXml(employee, "TaxCodeChangeTypeID");
                        rpEmployeePeriod.TaxCodeChangeType = prWG.GetElementByTagFromXml(employee, "TaxCodeChangeType");
                        rpEmployeePeriod.TaxPrev = prWG.GetDecimalElementByTagFromXml(employee, "TaxPrevious");
                        rpEmployeePeriod.TaxablePayPrevious = prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TaxThis = prWG.GetDecimalElementByTagFromXml(employee, "TaxThis");
                        rpEmployeePeriod.TaxablePayYTD = prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayThisYTD") + prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayPrevious");
                        rpEmployeePeriod.TaxablePayTP = prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayThisPeriod");
                        rpEmployeePeriod.HolidayAccruedTd = prWG.GetDecimalElementByTagFromXml(employee, "HolidayAccruedTd");
                        //
                        //AE Assessment
                        //
                        RPAEAssessment rpAEAssessment = new RPAEAssessment();
                        foreach(XmlElement aeAssessment in employee.GetElementsByTagName("AEAssessment"))
                        {
                            rpAEAssessment.Age = prWG.GetIntElementByTagFromXml(aeAssessment, "Age");
                            rpAEAssessment.StatePensionAge = prWG.GetIntElementByTagFromXml(aeAssessment, "StatePensionAge");
                            rpAEAssessment.StatePensionDate = prWG.GetDateElementByTagFromXml(aeAssessment, "StatePensionDate");
                            rpAEAssessment.AssessmentDate = prWG.GetDateElementByTagFromXml(aeAssessment, "AssessmentDate");
                            rpAEAssessment.QualifyingEarnings = prWG.GetDecimalElementByTagFromXml(aeAssessment, "QualifyingEarnings");
                            rpAEAssessment.AssessmentCode = prWG.GetElementByTagFromXml(aeAssessment, "AssessmentCode");
                            rpAEAssessment.AssessmentEvent = prWG.GetElementByTagFromXml(aeAssessment, "AssessmentEvent");
                            rpAEAssessment.AssessmentResult = prWG.GetElementByTagFromXml(aeAssessment, "AssessmentResult");
                            rpAEAssessment.AssessmentOverride = prWG.GetElementByTagFromXml(aeAssessment, "AssessmentOverride");
                            rpAEAssessment.OptOutWindowEndDate = prWG.GetDateElementByTagFromXml(aeAssessment, "OptOutWindowEndDate");
                            rpAEAssessment.ReenrolmentDate = prWG.GetDateElementByTagFromXml(aeAssessment, "ReenrolmentDate");
                            rpAEAssessment.IsMemberOfAlternativePensionScheme = prWG.GetBooleanElementByTagFromXml(aeAssessment, "IsMemberOfAlternativePensionScheme");
                            rpAEAssessment.TaxYear = prWG.GetIntElementByTagFromXml(aeAssessment, "TaxYear");
                            rpAEAssessment.TaxPeriod = prWG.GetIntElementByTagFromXml(aeAssessment, "TaxPeriod");
                        }
                        //Split these strings on capital letters by inserting a space before each capital letter.
                        rpAEAssessment.AssessmentCode = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentCode);
                        rpAEAssessment.AssessmentEvent = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentEvent);
                        rpAEAssessment.AssessmentResult = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentResult);
                        rpAEAssessment.AssessmentOverride = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentOverride);

                        rpEmployeePeriod.AEAssessment = rpAEAssessment;

                        rpEmployeePeriod.ErPensionTotalTP = 0;
                        rpEmployeePeriod.ErPensionTotalYtd = 0;
                        List<RPPensionPeriod> rpPensionPeriods = new List<RPPensionPeriod>();
                        foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                        {
                            RPPensionPeriod rpPensionPeriod = new RPPensionPeriod();
                            rpPensionPeriod.Key = Convert.ToInt32(pension.GetAttribute("Key"));
                            rpPensionPeriod.Code = prWG.GetElementByTagFromXml(pension, "Code");
                            rpPensionPeriod.SchemeName = prWG.GetElementByTagFromXml(pension, "SchemeName");
                            rpPensionPeriod.StartJoinDate = prWG.GetDateElementByTagFromXml(pension, "StartJoinDate");
                            rpPensionPeriod.IsJoiner = prWG.GetBooleanElementByTagFromXml(pension, "IsJoiner");
                            rpPensionPeriod.ProviderEmployerReference = prWG.GetElementByTagFromXml(pension, "ProviderEmployerRef");
                            rpPensionPeriod.EePensionYtd = prWG.GetDecimalElementByTagFromXml(pension, "EePensionYtd");
                            rpPensionPeriod.ErPensionYtd = prWG.GetDecimalElementByTagFromXml(pension, "ErPensionYtd");
                            rpPensionPeriod.PensionablePayYtd = prWG.GetDecimalElementByTagFromXml(pension, "PensionablePayYtd");
                            rpPensionPeriod.EePensionTaxPeriod = prWG.GetDecimalElementByTagFromXml(pension, "EePensionTaxPeriod");
                            rpPensionPeriod.ErPensionTaxPeriod = prWG.GetDecimalElementByTagFromXml(pension, "ErPensionTaxPeriod");
                            rpPensionPeriod.PensionablePayTaxPeriod = prWG.GetDecimalElementByTagFromXml(pension, "PensionablePayTaxPeriod");
                            rpPensionPeriod.EePensionPayRunDate = prWG.GetDecimalElementByTagFromXml(pension, "EePensionPayRunDate");
                            rpPensionPeriod.ErPensionPayRunDate = prWG.GetDecimalElementByTagFromXml(pension, "ErPensionPayRunDate");
                            rpPensionPeriod.PensionablePayPayRunDate = prWG.GetDecimalElementByTagFromXml(pension, "PensionablePayDate");
                            rpPensionPeriod.EeContibutionPercent = prWG.GetDecimalElementByTagFromXml(pension, "EeContributionPercent") * 100;
                            rpPensionPeriod.ErContributionPercent = prWG.GetDecimalElementByTagFromXml(pension, "ErContributionPercent") * 100;
                            rpEmployeePeriod.Frequency = rpParameters.PaySchedule;

                            rpPensionPeriods.Add(rpPensionPeriod);

                            RPPensionContribution rpPensionContribution = new RPPensionContribution();
                            rpPensionContribution.EeRef = rpEmployeePeriod.Reference;
                            rpPensionContribution.Title = rpEmployeePeriod.Title;
                            rpPensionContribution.Forename = rpEmployeePeriod.Forename;
                            rpPensionContribution.Surname = rpEmployeePeriod.Surname;
                            rpPensionContribution.Fullname = rpEmployeePeriod.Fullname;
                            rpPensionContribution.SurnameForename = rpEmployeePeriod.SurnameForename;
                            rpPensionContribution.ForenameSurname = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                            rpPensionContribution.DOB = rpEmployeePeriod.DateOfBirth;

                            RPAddress rpAddress = new RPAddress();
                            rpAddress.Line1 = rpEmployeePeriod.Address1;
                            rpAddress.Line2 = rpEmployeePeriod.Address2;
                            rpAddress.Line3 = rpEmployeePeriod.Address3;
                            rpAddress.Line4 = rpEmployeePeriod.Address4;
                            rpAddress.Postcode = rpEmployeePeriod.Postcode;
                            rpAddress.Country = rpEmployeePeriod.Country;

                            rpPensionContribution.RPAddress = rpAddress;
                            rpPensionContribution.EmailAddress = "";
                            rpPensionContribution.Gender = rpEmployeePeriod.Gender;
                            rpPensionContribution.NINumber = rpEmployeePeriod.NINumber;
                            rpPensionContribution.Freq = rpEmployeePeriod.Frequency;
                            rpPensionContribution.PayRunDate = rpEmployeePeriod.PayRunDate;
                            rpPensionContribution.StartDate = rpEmployeePeriod.PeriodStartDate;
                            rpPensionContribution.EndDate = rpEmployeePeriod.PeriodEndDate;
                            rpPensionContribution.RPPensionPeriod = rpPensionPeriod;

                            rpPensionContributions.Add(rpPensionContribution);

                            rpEmployeePeriod.ErPensionTotalTP = rpEmployeePeriod.ErPensionTotalTP + rpPensionPeriod.ErPensionTaxPeriod;
                            rpEmployeePeriod.ErPensionTotalYtd = rpEmployeePeriod.ErPensionTotalYtd + rpPensionPeriod.ErPensionYtd;
                        }
                        rpEmployeePeriod.Pensions = rpPensionPeriods;

                        rpEmployeePeriod.DirectorshipAppointmentDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "DirectorshipAppointmentDate"));
                        rpEmployeePeriod.Director = prWG.GetBooleanElementByTagFromXml(employee, "Director");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt1 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt1");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt2 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt2");
                        rpEmployeePeriod.ErNICTP = prWG.GetDecimalElementByTagFromXml(employee, "ErContributionTaxPeriod");
                        rpEmployeePeriod.NetPayYTD = 0;
                        rpEmployeePeriod.TotalPayTP = 0;
                        rpEmployeePeriod.TotalPayYTD = 0;
                        rpEmployeePeriod.TotalDedTP = 0;
                        rpEmployeePeriod.TotalDedYTD = 0;
                        rpEmployeePeriod.ErNICTP = prWG.GetDecimalElementByTagFromXml(employee, "ErContributionsTaxPeriod");
                        rpEmployeePeriod.ErNICYTD = prWG.GetDecimalElementByTagFromXml(employee, "ErContributions");
                        rpEmployeePeriod.PensionCode = prWG.GetElementByTagFromXml(employee, "PensionDetails");
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
                        List<RPPayslipDeduction> rpPayslipDeductions = new List<RPPayslipDeduction>();
                        foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                        {
                            foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                            {
                                if (!preSamplePayCodes)
                                {
                                    //Make a list of every possible Pay Code.
                                    RPPreSamplePayCode rpPreSamplePayCode = new RPPreSamplePayCode();
                                    rpPreSamplePayCode.Code = prWG.GetElementByTagFromXml(payCode, "Code");
                                    rpPreSamplePayCode.Description = prWG.GetElementByTagFromXml(payCode, "Description");
                                    rpPreSamplePayCode.InUse = false; //Set them all to false to begin with. If there a value it subsequently get set to true.
                                    rpPreSamplePayCodes.Add(rpPreSamplePayCode);

                                }


                                //Make a list of Pay Codes with values and which have IsPayCode set to true.
                                RPPayComponent rpPayComponent = new RPPayComponent();
                                rpPayComponent.PayCode = prWG.GetElementByTagFromXml(payCode, "Code");
                                rpPayComponent.Description = prWG.GetElementByTagFromXml(payCode, "Description");
                                rpPayComponent.EeRef = rpEmployeePeriod.Reference;
                                rpPayComponent.Fullname = rpEmployeePeriod.Fullname;
                                rpPayComponent.SurnameForename = rpEmployeePeriod.SurnameForename;
                                rpPayComponent.Surname = rpEmployeePeriod.Surname;
                                rpPayComponent.Rate = prWG.GetDecimalElementByTagFromXml(payCode, "Rate");
                                rpPayComponent.UnitsTP = prWG.GetDecimalElementByTagFromXml(payCode, "Units");
                                rpPayComponent.AmountTP = prWG.GetDecimalElementByTagFromXml(payCode, "Amount");
                                rpPayComponent.UnitsYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                rpPayComponent.AmountYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                rpPayComponent.AccountsYearBalance = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance");
                                rpPayComponent.AccountsYearUnits = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits");
                                rpPayComponent.PayrollAccrued = prWG.GetDecimalElementByTagFromXml(payCode, "PayrollAccrued");
                                rpPayComponent.IsTaxable = prWG.GetBooleanElementByTagFromXml(payCode, "IsTaxable");
                                rpPayComponent.IsPayCode = prWG.GetBooleanElementByTagFromXml(payCode, "IsPayCode");
                                rpPayComponent.EarningOrDeduction = prWG.GetElementByTagFromXml(payCode, "EarningOrDeduction");
                                if (rpPayComponent.AmountTP != 0 || rpPayComponent.AmountYTD != 0)
                                {
                                    //Value is not equal to zero so go through the list of Pre Sample codes and mark this one as in use
                                    rpPreSamplePayCodes = MarkPreSampleCodeAsInUse(rpPayComponent.PayCode, rpPreSamplePayCodes);
                                    if (rpPayComponent.IsPayCode)
                                    {
                                        rpPayComponents.Add(rpPayComponent);
                                    }
                                    if(rpPayComponent.PayCode != "TAX" && rpPayComponent.PayCode != "NI" && !rpPayComponent.PayCode.StartsWith("PENSION"))
                                    {
                                        if (rpPayComponent.IsTaxable)
                                        {
                                            rpEmployeePeriod.PreTaxAddDed = rpEmployeePeriod.PreTaxAddDed + rpPayComponent.AmountTP;
                                        }
                                        else
                                        {
                                            rpEmployeePeriod.PostTaxAddDed = rpEmployeePeriod.PostTaxAddDed + rpPayComponent.AmountTP;
                                        }
                                    }

                                    //Check for the different pay codes and add to the appropriate total.
                                    switch (rpPayComponent.PayCode)
                                    {
                                        case "HOLPY":
                                        case "HOLIDAY":
                                            rpEmployeePeriod.HolidayPay = rpEmployeePeriod.HolidayPay + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSIONRAS":
                                        case "PENSIONTAXEX":
                                            rpEmployeePeriod.PostTaxPension = rpEmployeePeriod.PostTaxPension + rpPayComponent.AmountTP;
                                            break;
                                        case "PENSION":
                                        case "PENSIONSS":
                                            rpEmployeePeriod.PreTaxPension = rpEmployeePeriod.PreTaxPension + rpPayComponent.AmountTP;
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
                                            break;

                                    }
                                }


                                //if (prWG.GetElementByTagFromXml(payCode, "EarningOrDeduction") == "E")
                                if (rpPayComponent.EarningOrDeduction == "E" && rpPayComponent.PayCode != "PENSIONSS")
                                {
                                    RPAddition rpAddition = new RPAddition();
                                    rpAddition.EeRef = rpEmployeePeriod.Reference;
                                    rpAddition.Code = rpPayComponent.PayCode;//prWG.GetElementByTagFromXml(payCode, "Code");
                                    //They want Basic pay and Salary to come first. This will only work if they use the following codes!
                                    switch(rpAddition.Code)
                                    {
                                        case "BASCH":
                                            rpAddition.Code = " BASCH";
                                            break;
                                        case "BASIC":
                                            rpAddition.Code = " BASIC";
                                            break;
                                        case "SALRY":
                                            rpAddition.Code = " SALRY";
                                            break;
                                        case "SALARY":
                                            rpAddition.Code = " SALARY";
                                            break;
                                    }

                                    rpAddition.Description = rpPayComponent.Description;//prWG.GetElementByTagFromXml(payCode, "Description");
                                    rpAddition.Rate = rpPayComponent.Rate;//prWG.GetDecimalElementByTagFromXml(payCode, "Rate");
                                    rpAddition.Units = rpPayComponent.UnitsTP;//prWG.GetDecimalElementByTagFromXml(payCode, "Units");
                                    rpAddition.AmountTP = rpPayComponent.AmountTP;//prWG.GetDecimalElementByTagFromXml(payCode, "Amount");
                                    rpAddition.AmountYTD = rpPayComponent.AmountYTD;//prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                    rpAddition.AccountsYearBalance = rpPayComponent.AccountsYearBalance;//prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance");
                                    rpAddition.AccountsYearUnits = rpPayComponent.AccountsYearUnits;//prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits");
                                    rpAddition.PayeYearUnits = rpPayComponent.UnitsYTD;//prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                    rpAddition.PayrollAccrued = rpPayComponent.PayrollAccrued;//prWG.GetDecimalElementByTagFromXml(payCode, "PayrollAccrued");
                                    if (rpAddition.AmountTP != 0)
                                    {
                                        rpAdditions.Add(rpAddition);
                                        
                                    }
                                    rpEmployeePeriod.TotalPayTP = rpEmployeePeriod.TotalPayTP + rpAddition.AmountTP;
                                    rpEmployeePeriod.TotalPayYTD = rpEmployeePeriod.TotalPayYTD + rpAddition.AmountYTD;
                                }
                                else
                                {
                                    //We now need a list of deductions for the PayHistory.csv file and a different one for the payslips.
                                    //Deductions used to create the PayHistory.csv file will use the PayCodes provided in the PR xml file for pensions, for the payslip use the pension list from PR.
                                    RPDeduction rpDeduction = new RPDeduction();
                                    rpDeduction.EeRef = rpEmployeePeriod.Reference;
                                    rpDeduction.Code = rpPayComponent.PayCode;//prWG.GetElementByTagFromXml(payCode, "Code");
                                    //They want Tax then NI, then Pension to come first, then the rest in alphabetical order. This will only work if they use the following codes!
                                    switch (rpDeduction.Code)
                                    {
                                        case "TAX":
                                            rpDeduction.Seq = "00" + rpDeduction.Code;
                                            break;
                                        case "NI":
                                            rpDeduction.Seq = "01" + rpDeduction.Code;
                                            break;
                                        case "PENSION":
                                        case "PENSIONRAS":
                                        case "PENSIONSS":
                                        case "PENSIONTAXEX":
                                            rpDeduction.Seq = "02" + rpDeduction.Code;
                                            break;
                                        default:
                                            rpDeduction.Seq = "99" + rpDeduction.Code;
                                            break;

                                    }
                                    rpDeduction.Description = rpPayComponent.Description;//prWG.GetElementByTagFromXml(payCode, "Description");
                                    rpDeduction.IsTaxable = rpPayComponent.IsTaxable;
                                    rpDeduction.AmountTP = rpPayComponent.AmountTP * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "Amount") * -1;
                                    rpDeduction.AmountYTD = rpPayComponent.AmountYTD * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance") * -1;
                                    rpDeduction.AccountsYearBalance = rpPayComponent.AccountsYearBalance * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance") * -1;
                                    rpDeduction.AccountsYearUnits = rpPayComponent.AccountsYearUnits * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits") * -1;
                                    rpDeduction.PayeYearUnits = rpPayComponent.UnitsYTD * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits") * -1;
                                    rpDeduction.PayrollAccrued = rpPayComponent.PayrollAccrued * -1;//prWG.GetDecimalElementByTagFromXml(payCode, "PayrollAccrued") * -1;
                                    //if (rpDeduction.AmountTP != 0 || rpDeduction.AmountYTD != 0)
                                    if (rpDeduction.AmountTP != 0)
                                    {
                                        rpDeductions.Add(rpDeduction);

                                    }
                                    rpEmployeePeriod.TotalDedTP = rpEmployeePeriod.TotalDedTP + rpDeduction.AmountTP;
                                    rpEmployeePeriod.TotalDedYTD = rpEmployeePeriod.TotalDedYTD + rpDeduction.AmountYTD;


                                    //We now need a list of deductions for the PayHistory.csv file and a different one for the payslips.
                                    //Deductions used to create the PayHistory.csv file will use the PayCodes provided in the PR xml file for pensions, for the payslip use the pension list from PR.
                                    if(!rpDeduction.Code.Contains("PENSION"))
                                    {
                                        RPPayslipDeduction rpPayslipDeduction = new RPPayslipDeduction();
                                        rpPayslipDeduction.EeRef = rpEmployeePeriod.Reference;
                                        rpPayslipDeduction.Code = rpDeduction.Code;
                                        rpPayslipDeduction.Seq = rpDeduction.Seq;
                                        rpPayslipDeduction.Description = rpDeduction.Description;
                                        rpPayslipDeduction.AmountTP = rpDeduction.AmountTP;
                                        rpPayslipDeduction.AmountYTD = rpDeduction.AmountYTD;
                                        //if (rpDeduction.AmountTP != 0 || rpDeduction.AmountYTD != 0)
                                        if (rpPayslipDeduction.AmountTP != 0)
                                        {
                                            rpPayslipDeductions.Add(rpPayslipDeduction);

                                        }
                                    }
                                    

                                }
                                
                            }//End of for each payCode
                            preSamplePayCodes = true;
                        }//End of for each payCodes
                        //
                        //Deductions are only used for the payslip. It's possible that we should using the pension list for the pension elements,
                        //in which case we shouldn't use these pension pay codes but use the rpPensionPeriods list we created above instead
                        //
                        foreach (RPPensionPeriod rpPensionPeriod in rpPensionPeriods)
                        {
                            RPPayslipDeduction rpPayslipDeduction = new RPPayslipDeduction();
                            rpPayslipDeduction.EeRef = rpEmployeePeriod.Reference;
                            rpPayslipDeduction.Seq = "02PENSION";
                            rpPayslipDeduction.Code = "PENSION" + rpPensionPeriod.Code;
                            rpPayslipDeduction.Description = rpPensionPeriod.SchemeName;
                            rpPayslipDeduction.AmountTP = rpPensionPeriod.EePensionTaxPeriod;
                            rpPayslipDeduction.AmountYTD = rpPensionPeriod.EePensionYtd;
                            //if (rpPayslipDeduction.AmountTP != 0 || rpPayslipDeduction.AmountYTD != 0)
                            if (rpPayslipDeduction.AmountTP != 0)
                            {
                                rpPayslipDeductions.Add(rpPayslipDeduction);

                            }
                            
                        }
                        
                        //Sort the list of additions into Code sequence before returning them.
                        rpAdditions.Sort(delegate (RPAddition x, RPAddition y)
                        {
                            if (x.Code == null && y.Code == null) return 0;
                            else if (x.Code == null) return -1;
                            else if (y.Code == null) return 1;
                            else return x.Code.CompareTo(y.Code);
                        });
                        //Sort the list of deductions into Code sequence before returning them.
                        rpDeductions.Sort(delegate (RPDeduction x, RPDeduction y)
                        {
                            if (x.Seq == null && y.Seq == null) return 0;
                            else if (x.Seq == null) return -1;
                            else if (y.Seq == null) return 1;
                            else return x.Seq.CompareTo(y.Seq);
                        });
                        //Sort the list of payslip deductions into Code sequence before returning them.
                        rpPayslipDeductions.Sort(delegate (RPPayslipDeduction x, RPPayslipDeduction y)
                        {
                            if (x.Seq == null && y.Seq == null) return 0;
                            else if (x.Seq == null) return -1;
                            else if (y.Seq == null) return 1;
                            else return x.Seq.CompareTo(y.Seq);
                        });
                        rpEmployeePeriod.Additions = rpAdditions;
                        rpEmployeePeriod.Deductions = rpDeductions;
                        rpEmployeePeriod.PayslipDeductions = rpPayslipDeductions;

                        //Multiple Tax and NI by -1 to make them positive
                        rpEmployeePeriod.Tax = rpEmployeePeriod.Tax * -1;
                        rpEmployeePeriod.NetNI = rpEmployeePeriod.NetNI * -1;
                        //Multiple the Pre-Tax Pension & Post-Tax pension by -1 to make them show as positive on the Payroll Run Details report.
                        rpEmployeePeriod.PreTaxPension = rpEmployeePeriod.PreTaxPension * -1;
                        rpEmployeePeriod.PostTaxPension = rpEmployeePeriod.PostTaxPension * -1;

                        //We also have a list of pay codes which are in use. We will use these to create the Pre Sample xlsx file.
                        //foreach(RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
                        //{
                        //    if(rpPreSamplePayCode.InUse==true)
                        //    {
                        //        //This is one that is in use.
                        //    }
                        //}
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
                            p45.LeavingDate = Convert.ToDateTime(rpEmployeePeriod.LeavingDate);
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
                        address = prWG.RemoveBlankAddressLines(address);
                        rpEmployeePeriod.Address1 = address[0];
                        rpEmployeePeriod.Address2 = address[1];
                        rpEmployeePeriod.Address3 = address[2];
                        rpEmployeePeriod.Address4 = address[3];
                        rpEmployeePeriod.Postcode = address[4];
                        rpEmployeePeriod.Country = address[5];

                        rpEmployeePeriodList.Add(rpEmployeePeriod);
                    }//End of for each employee


                }
                //Sort the list of employees into EeRef sequence before returning them.
                rpEmployeePeriodList.Sort(delegate (RPEmployeePeriod x, RPEmployeePeriod y)
                {
                    if (x.Reference == null && y.Reference == null) return 0;
                    else if (x.Reference == null) return -1;
                    else if (y.Reference == null) return 1;
                    else return x.Reference.CompareTo(y.Reference);
                });
                //Sort the list of pension contributions into Scheme Name,EeRef sequence before returning them.
                rpPensionContributions.Sort(delegate (RPPensionContribution x, RPPensionContribution y)
                {
                    if ((x.RPPensionPeriod.SchemeName + x.EeRef) == null && (y.RPPensionPeriod.SchemeName + y.EeRef) == null) return 0;
                    else if ((x.RPPensionPeriod.SchemeName + x.EeRef) == null) return -1;
                    else if ((y.RPPensionPeriod.SchemeName + y.EeRef) == null) return 1;
                    else return (x.RPPensionPeriod.SchemeName + x.EeRef).CompareTo(y.RPPensionPeriod.SchemeName + y.EeRef);
                });

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error preparing reports.\r\n{0}.\r\n", ex);
                prWG.update_Progress(textLine, configDirName, logOneIn);
            }
            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                                  List<RPPensionContribution>, RPEmployer, RPParameters>
                                  (rpEmployeePeriodList, rpPayComponents, p45s, rpPreSamplePayCodes, rpPensionContributions, rpEmployer, rpParameters);

        }
        private string SplitStringOnCapitalLetters(string input)
        {
            string output = null;
            if(input != null)
            {
                var r = new Regex(@"
                                        (?<=[A-Z])(?=[A-Z][a-z]) |
                                         (?<=[^A-Z])(?=[A-Z]) |
                                         (?<=[A-Za-z])(?=[^A-Za-z])", RegexOptions.IgnorePatternWhitespace);


                output = r.Replace(input, " ");
            }
            
            return output;
        }
        
        private List<RPPreSamplePayCode> MarkPreSampleCodeAsInUse(string payCode, List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            foreach(RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
            {
                if(rpPreSamplePayCode.Code == payCode)
                {
                    rpPreSamplePayCode.InUse = true;
                }
            }
            return rpPreSamplePayCodes;
        }
        public bool ProduceReports(XDocument xdoc, string directory)
        {
            //Old method going through directories created by PR
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            bool eePeriodProcessed = false;
            bool eeYtdProcessed = false;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            DirectoryInfo dirInfo = new DirectoryInfo(directory);
            FileInfo[] files = dirInfo.GetFiles("*.xml");
            //We haven't got the correct payroll run date in the "EmployeeYtd" report so I'm going use the RPParameters from the "EmployeePeriod" report.
            //I'm just a bit concerned of the order they will come in. Hopefully always alphabetical.
            RPParameters rpParameters = null;
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("EmployeePeriod"))
                {
                    List<RPEmployeePeriod> rpEmployeePeriodList = null;
                    List<RPPayComponent> rpPayComponents = null;
                    List<P45> p45s = null;
                    List<RPPreSamplePayCode> rpPreSamplePayCodes = null;
                    List<RPPensionContribution> rpPensionContributions = null;
                    RPEmployer rpEmployer = null;
                    
                    try
                    {
                        var tuple = PreparePeriodReport(xdoc, file);
                        rpEmployeePeriodList = tuple.Item1;
                        rpPayComponents = tuple.Item2;
                        p45s = tuple.Item3;
                        rpPreSamplePayCodes = tuple.Item4;
                        rpPensionContributions = tuple.Item5;
                        rpEmployer = tuple.Item6;
                        rpParameters = tuple.Item7;
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error preparing the employee period reports for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.update_Progress(textLine, configDirName, logOneIn);
                    }
                    try
                    {
                        prWG.CreateHistoryCSV(xdoc, rpParameters, rpEmployer, rpEmployeePeriodList);
                    }
                    catch(Exception ex)
                    {
                        textLine = string.Format("Error creating the history csv file for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.update_Progress(textLine, configDirName, logOneIn);
                    }

                    try
                    {
                        ProducePeriodReports(xdoc, rpEmployeePeriodList, rpEmployer, p45s, rpPayComponents, rpParameters, rpPreSamplePayCodes, rpPensionContributions);

                        eePeriodProcessed = true;
                    }   
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error producing the employee period reports for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.update_Progress(textLine, configDirName, logOneIn);
                    }
                   
                }
                else if (file.FullName.Contains("EmployeeYtd"))
                {
                    try
                    {
                        var tuple = prWG.PrepareYTDReport(xdoc, file);
                        List<RPEmployeeYtd> rpEmployeeYtdList = tuple.Item1;
                        //I'm going to use the RPParameters from the "EmployeePeriod" report for now at least.
                        //RPParameters rpParameters = tuple.Item2;
                        prWG.CreateYTDCSV(xdoc, rpEmployeeYtdList, rpParameters);
                        eeYtdProcessed = true;
                    }
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error producing the employee ytd report for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.update_Progress(textLine, configDirName, logOneIn);
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
        private Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                           List<RPPensionContribution>, RPEmployer, RPParameters> 
                           PreparePeriodReport(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlPeriodReport = new XmlDocument();
            xmlPeriodReport.Load(file.FullName);
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = prWG.GetRPParameters(xmlPeriodReport);
            //2
            var tuple = PrepareStandardReports(xdoc, xmlPeriodReport, rpParameters);
            List<RPEmployeePeriod> rpEmployeePeriodList = tuple.Item1;
            List<RPPayComponent> rpPayComponents = tuple.Item2;
            //I don't think the P45 report will be able to be produced from the EmployeePeriod report but I'm leaving it here for now.
            List<P45> p45s = tuple.Item3;
            List<RPPreSamplePayCode> rpPreSamplePayCodes = tuple.Item4;
            List<RPPensionContribution> rpPensionContributions = tuple.Item5;
            RPEmployer rpEmployer = tuple.Item6;
            rpParameters = tuple.Item7;
            
            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                                  List<RPPensionContribution>, RPEmployer, RPParameters>
                                  (rpEmployeePeriodList, rpPayComponents, p45s, rpPreSamplePayCodes, rpPensionContributions, rpEmployer, rpParameters);

        }
        private RPP32Report CreateP32Report(XDocument xdoc, RPEmployer rpEmplopyer, RPParameters rpParameters)
        {
            RPP32Report rpP32Report = null;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            XmlDocument p32ReportXml = new XmlDocument();

            bool test = false;
            if(test)
            {
                p32ReportXml.Load("C:\\Payescape\\Data\\Save\\P32.xml");
            }
            else
            {
                p32ReportXml = prWG.GetP32Report(rpParameters);
            }

            rpP32Report = PrepareP32SummaryReport(xdoc, p32ReportXml, rpParameters, prWG);

            return rpP32Report;
        }
        private RPP32Report PrepareP32SummaryReport(XDocument xdoc, XmlDocument p32ReportXml, RPParameters rpParameters, PayRunIOWebGlobeClass prWG)
        {
            RPP32Report rpP32Report = new RPP32Report();
            foreach (XmlElement header in p32ReportXml.GetElementsByTagName("Header"))
            {
                rpP32Report.EmployerName = prWG.GetElementByTagFromXml(header, "EmployerName");
                rpP32Report.EmployerPayeRef = prWG.GetElementByTagFromXml(header, "EmployerPayeRef");
                rpP32Report.PaymentRef = prWG.GetElementByTagFromXml(header, "PaymentRef");
                rpP32Report.TaxYear = prWG.GetIntElementByTagFromXml(header, "TaxYear");
                rpP32Report.TaxYearStartDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(header, "TaxYearStart"));
                rpP32Report.TaxYearEndDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(header, "TaxYearEndDate"));
                rpP32Report.AnnualEmploymentAllowance = prWG.GetIntElementByTagFromXml(header, "AnnualEmploymentAllowance");
            }
            bool addToList = false;
            bool annualTotalRequired = false;
            List<RPP32ReportMonth> rpP32ReportMonths = new List<RPP32ReportMonth>();
            foreach(XmlElement reportMonth in p32ReportXml.GetElementsByTagName("ReportMonth"))
            {
                RPP32ReportMonth rpP32ReportMonth = new RPP32ReportMonth();
                rpP32ReportMonth.PeriodNo = Convert.ToInt32(reportMonth.GetAttribute("Period"));
                rpP32ReportMonth.RPPeriodNo = rpP32ReportMonth.PeriodNo.ToString();
                rpP32ReportMonth.RPPeriodText = "Month " + rpP32ReportMonth.PeriodNo.ToString();
                rpP32ReportMonth.PeriodName = reportMonth.GetAttribute("RootNodeName");

                RPP32Breakdown rpP32Breakdown = new RPP32Breakdown();
                List<RPP32Schedule> rpP32Schedules = new List<RPP32Schedule>();

                foreach (XmlElement paySchedule in reportMonth.GetElementsByTagName("PaySchedule"))
                {
                    RPP32Schedule rpP32Schedule = new RPP32Schedule();
                    rpP32Schedule.PayScheduleName = paySchedule.GetAttribute("Name");
                    rpP32Schedule.PayScheduleFrequency = paySchedule.GetAttribute("Frequency");
                    List<RPP32PayRun> rpP32PayRuns = new List<RPP32PayRun>();
                    foreach(XmlElement payRun in paySchedule.GetElementsByTagName("PayRun"))
                    {
                        RPP32PayRun rpP32PayRun = new RPP32PayRun();
                        rpP32PayRun.PayDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(payRun, "PayDate"));
                        rpP32PayRun.IncomeTax = prWG.GetDecimalElementByTagFromXml(payRun, "IncomeTax");
                        rpP32PayRun.StudentLoan = prWG.GetDecimalElementByTagFromXml(payRun, "StudentLoan");
                        rpP32PayRun.PostGraduateLoan = prWG.GetDecimalElementByTagFromXml(payRun, "PostGraduateLoan");
                        rpP32PayRun.NetIncomeTax = prWG.GetDecimalElementByTagFromXml(payRun, "NetIncomeTax");
                        rpP32PayRun.GrossNICs = prWG.GetDecimalElementByTagFromXml(payRun, "GrossNICs");

                        rpP32PayRuns.Add(rpP32PayRun);
                    }
                    if(rpP32PayRuns.Count > 0)
                    {
                        rpP32Schedule.RPP32PayRuns = rpP32PayRuns;

                        
                    }
                    try
                    {
                        rpP32Schedules.Add(rpP32Schedule);
                        
                    }
                    catch(Exception ex)
                    {

                    }
                    
                }
                try
                {
                    rpP32Breakdown.RPP32Schedules = rpP32Schedules;

                    rpP32ReportMonth.RPP32Breakdown = rpP32Breakdown;
                }
                catch(Exception ex)
                {

                }
                

                RPP32Summary rpP32Summary = new RPP32Summary();

                foreach(XmlElement summary in reportMonth.GetElementsByTagName("Summary"))
                {
                    rpP32Summary.Tax = prWG.GetDecimalElementByTagFromXml(summary, "Tax");
                    rpP32Summary.StudentLoan= prWG.GetDecimalElementByTagFromXml(summary, "StudentLoan");
                    rpP32Summary.PostGraduateLoan = prWG.GetDecimalElementByTagFromXml(summary, "PostGraduateLoan");
                    rpP32Summary.NetTax = prWG.GetDecimalElementByTagFromXml(summary, "NetTax");
                    rpP32Summary.EmployerNI = prWG.GetDecimalElementByTagFromXml(summary, "EmployerNI");
                    rpP32Summary.EmployeeNI = prWG.GetDecimalElementByTagFromXml(summary, "EmployeeNI");
                    rpP32Summary.GrossNICs = prWG.GetDecimalElementByTagFromXml(summary, "GrossNICs");
                    rpP32Summary.SmpRecovered = prWG.GetDecimalElementByTagFromXml(summary, "SmpRecovered");
                    rpP32Summary.SmpComp = prWG.GetDecimalElementByTagFromXml(summary, "SmpComp");
                    rpP32Summary.SppRecovered = prWG.GetDecimalElementByTagFromXml(summary, "SppRecovered");
                    rpP32Summary.SppComp = prWG.GetDecimalElementByTagFromXml(summary, "SppComp");
                    rpP32Summary.ShppRecovered = prWG.GetDecimalElementByTagFromXml(summary, "ShppRecovered");
                    rpP32Summary.ShppComp = prWG.GetDecimalElementByTagFromXml(summary, "ShppComp");
                    rpP32Summary.SapRecovered = prWG.GetDecimalElementByTagFromXml(summary, "SapRecovered");
                    rpP32Summary.SapComp = prWG.GetDecimalElementByTagFromXml(summary, "SapComp");
                    rpP32Summary.TotalDeductions = rpP32Summary.EmploymentAllowance + rpP32Summary.SmpComp + rpP32Summary.SmpRecovered + rpP32Summary.SppComp + 
                                                   rpP32Summary.SppRecovered + rpP32Summary.SapComp + rpP32Summary.SapRecovered + rpP32Summary.ShppComp + 
                                                   rpP32Summary.ShppRecovered;
                    rpP32Summary.AppLevy = 0;
                    rpP32Summary.CisDeducted = prWG.GetDecimalElementByTagFromXml(summary, "CisDeducted");
                    rpP32Summary.CisSuffered = prWG.GetDecimalElementByTagFromXml(summary, "CisSuffered");
                    rpP32Summary.NetNICs = prWG.GetDecimalElementByTagFromXml(summary, "NetNICs");
                    rpP32Summary.EmploymentAllowance = prWG.GetDecimalElementByTagFromXml(summary, "EmploymentAllowance");
                    rpP32Summary.AmountDue = prWG.GetDecimalElementByTagFromXml(summary, "AmountDue");
                    rpP32Summary.AmountPaid = prWG.GetDecimalElementByTagFromXml(summary, "AmountPaid");
                    rpP32Summary.RemainingBalance = prWG.GetDecimalElementByTagFromXml(summary, "RemainingBalance");
                }

                rpP32ReportMonth.RPP32Summary = rpP32Summary;

                //If any of the values are not zero add the P32 period to the list
                addToList = CheckIfNotZero(rpP32ReportMonth);

                if (addToList)
                {
                    rpP32ReportMonths.Add(rpP32ReportMonth);
                    annualTotalRequired = true;
                }
                
            }
            rpP32Report.RPP32ReportMonths = rpP32ReportMonths;

            if (annualTotalRequired)
            {
                RPP32ReportMonth rpP32ReportMonth = new RPP32ReportMonth();
                rpP32ReportMonth.PeriodNo = 13;
                rpP32ReportMonth.RPPeriodNo = "";
                rpP32ReportMonth.RPPeriodText = "Year " + rpP32Report.TaxYear.ToString();
                rpP32ReportMonth.PeriodName = "Annual total";

                //There is no breakdown for the annual total so just add a null one.
                RPP32Breakdown rpP32Breakdown = new RPP32Breakdown();
                rpP32ReportMonth.RPP32Breakdown = rpP32Breakdown;

                RPP32Summary rpP32Summary = new RPP32Summary();

                foreach (XmlElement annualTotal in p32ReportXml.GetElementsByTagName("AnnualTotal"))
                {
                    rpP32Summary.Tax = prWG.GetDecimalElementByTagFromXml(annualTotal, "Tax");
                    rpP32Summary.StudentLoan = prWG.GetDecimalElementByTagFromXml(annualTotal, "StudentLoan");
                    rpP32Summary.PostGraduateLoan = prWG.GetDecimalElementByTagFromXml(annualTotal, "PostGraduateLoan");
                    rpP32Summary.NetTax = prWG.GetDecimalElementByTagFromXml(annualTotal, "NetTax");
                    rpP32Summary.EmployerNI = prWG.GetDecimalElementByTagFromXml(annualTotal, "EmployerNI");
                    rpP32Summary.EmployeeNI = prWG.GetDecimalElementByTagFromXml(annualTotal, "EmployeeNI");
                    rpP32Summary.GrossNICs = prWG.GetDecimalElementByTagFromXml(annualTotal, "GrossNICs");
                    rpP32Summary.SmpRecovered = prWG.GetDecimalElementByTagFromXml(annualTotal, "SmpRecovered");
                    rpP32Summary.SmpComp = prWG.GetDecimalElementByTagFromXml(annualTotal, "SmpComp");
                    rpP32Summary.SppRecovered = prWG.GetDecimalElementByTagFromXml(annualTotal, "SppRecovered");
                    rpP32Summary.SppComp = prWG.GetDecimalElementByTagFromXml(annualTotal, "SppComp");
                    rpP32Summary.ShppRecovered = prWG.GetDecimalElementByTagFromXml(annualTotal, "ShppRecovered");
                    rpP32Summary.ShppComp = prWG.GetDecimalElementByTagFromXml(annualTotal, "ShppComp");
                    rpP32Summary.SapRecovered = prWG.GetDecimalElementByTagFromXml(annualTotal, "SapRecovered");
                    rpP32Summary.SapComp = prWG.GetDecimalElementByTagFromXml(annualTotal, "SapComp");
                    rpP32Summary.CisDeducted = prWG.GetDecimalElementByTagFromXml(annualTotal, "CisDeducted");
                    rpP32Summary.CisSuffered = prWG.GetDecimalElementByTagFromXml(annualTotal, "CisSuffered");
                    rpP32Summary.NetNICs = prWG.GetDecimalElementByTagFromXml(annualTotal, "NetNICs");
                    rpP32Summary.EmploymentAllowance = prWG.GetDecimalElementByTagFromXml(annualTotal, "EmploymentAllowance");
                    rpP32Summary.AmountDue = prWG.GetDecimalElementByTagFromXml(annualTotal, "AmountDue");
                    rpP32Summary.AmountPaid = prWG.GetDecimalElementByTagFromXml(annualTotal, "AmountPaid");
                    rpP32Summary.RemainingBalance = prWG.GetDecimalElementByTagFromXml(annualTotal, "RemainingBalance");
                }
                rpP32ReportMonth.RPP32Summary = rpP32Summary;

                rpP32Report.RPP32ReportMonths.Add(rpP32ReportMonth);
            }

            
            return rpP32Report;
        }
        private bool CheckIfNotZero(RPP32ReportMonth rpP32ReportMonth)
        {
            //Compare all the decimal fields to see if any are non zero using if
            //if(rpP32Period.Tax != 0 || rpP32Period.StudentLoan != 0 || rpP32Period.PostGraduateLoan != 0 || rpP32Period.NetTax != 0 || rpP32Period.EmployerNI != 0
            //    || rpP32Period.EmployeeNI != 0 || rpP32Period.GrossNICs != 0 || rpP32Period.SmpRecovered != 0 || rpP32Period.SmpComp != 0 || rpP32Period.SppRecovered !=0
            //    || rpP32Period.SppComp != 0 || rpP32Period.ShppRecovered !=0 || rpP32Period.ShppComp != 0 || rpP32Period.SapRecovered != 0 || rpP32Period.SapComp != 0
            //    || rpP32Period.CisDeducted != 0 | rpP32Period.CisSuffered != 0 || rpP32Period.NetNICs != 0 || rpP32Period.EmploymentAllowance !=0
            //    || rpP32Period.AmountDue != 0 || rpP32Period.AmountPaid != 0 || rpP32Period.RemainingBalance != 0)
            //{
            //    return true;
            //}

            //Compare all the decimal fields to see if any are non zero using reflection
            //foreach (PropertyInfo pi in rpP32ReportMonth.GetType().GetProperties() )
            //{
            //    if(pi.PropertyType==typeof(decimal))
            //    {
            //        decimal value = (decimal)pi.GetValue(rpP32ReportMonth);
            //        if(value != 0)
            //        {
            //           return true;
            //        }
            //    }
            //}
            //Compare all the decimal fields (in the Summary) to see if any are non zero using reflection
            foreach (PropertyInfo pi in rpP32ReportMonth.RPP32Summary.GetType().GetProperties())
            {
                if (pi.PropertyType == typeof(decimal))
                {
                    decimal value = (decimal)pi.GetValue(rpP32ReportMonth.RPP32Summary);
                    if (value != 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        private void CreatePreSampleXLSX(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            //Create a list of the required columns.
            List<string> reqCol = new List<string>();
            reqCol.Add("EeRef");
            reqCol.Add("Name");
            reqCol.Add("Dept");
            reqCol.Add("CostCentre");
            reqCol.Add("Branch");
            reqCol.Add("Status");
            reqCol.Add("TaxCode");
            reqCol.Add("NILetter");
            reqCol.Add("PreTaxAddDed");
            reqCol.Add("GrossedUpTaxThisRun");
            reqCol.Add("EeNIPdByEr");
            reqCol.Add("GUStudentLoan");
            reqCol.Add("GUNIReduction");
            reqCol.Add("PenPreTaxEeGU");
            reqCol.Add("TotalAbsencePay");
            reqCol.Add("HolidayPay");
            reqCol.Add("PenPreTaxEe");
            reqCol.Add("TaxablePay");
            reqCol.Add("Tax");
            reqCol.Add("NI");
            reqCol.Add("PostTaxAddDed");
            reqCol.Add("PostTaxPension");
            reqCol.Add("AOE");
            reqCol.Add("StudentLoan");
            reqCol.Add("NetPay");
            reqCol.Add("ErNI");
            reqCol.Add("PenEr");
            reqCol.Add("TotalGrossUp");
            
            RPEmployeePeriod rpEmployeePeriod = rpEmployeePeriodList.First();

            foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
            {
                reqCol.Add(rpAddition.Description);
            }
            foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
            {
                reqCol.Add(rpDeduction.Description);
            }

            //Need to count how many columns we are going to need
            string[] headings = new string[reqCol.Count()];
            int i = 0;
            foreach (string col in reqCol)
            {
                headings[i] = col.ToString();
            }
            //Create a workbook.
            Workbook workbook = new Workbook("X:\\Payescape\\PayRunIO\\PreSample.xlsx", "Pre Sample");
            //Write the headings.
            foreach (string heading in headings)
            {
                workbook.CurrentWorksheet.AddNextCell(heading);
            }
            //Move to the next row.
            workbook.CurrentWorksheet.GoToNextRow();
            //Now create a sample data line.
            //foreach (string column in columns)
            //{
            //    workbook.CurrentWorksheet.AddNextCell(column);
            //}
            //Save the workbook.
            workbook.Save();
        }
        private void CreatePreSampleXLSX(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList,
                                       RPEmployer rpEmployer, RPParameters rpParameters, List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            //Create a list of the required fixed columns.
            List<string> fixCol = new List<string>();
            fixCol = CreateListOfFixedColumns();

            //Create a list of the required variable columns.
            List<string> varCol = new List<string>();
            varCol = CreateListOfVariableColumns(rpPreSamplePayCodes);

            //Create a workbook.
            string workBookName = outgoingFolder + "\\" + coNo + "\\PreSample.xlsx";
            Workbook workbook = new Workbook(workBookName, "Pre Sample");
            foreach (string col in fixCol)
            {
                workbook.CurrentWorksheet.AddNextCell(col);
            }

            foreach (string col in varCol)
            {
                workbook.CurrentWorksheet.AddNextCell(col);
            }
            
            //Now for each employee create a row and add in the values for each column
            foreach(RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                if(rpEmployeePeriod.Reference=="81")
                {

                }
                workbook.CurrentWorksheet.GoToNextRow();

                workbook = CreateFixedWorkbookColumns(workbook, rpEmployeePeriod);
                workbook = CreateVariableWorkbookColumns(workbook, rpEmployeePeriod, varCol);
                
            }
            
            workbook.Save();
        }
        private List<string> CreateListOfFixedColumns()
        {
            //Create a list of the required fixed columns.
            List<string> fixCol = new List<string>();
            fixCol.Add("EeRef");
            fixCol.Add("Name");
            fixCol.Add("Dept");
            fixCol.Add("CostCentre");
            fixCol.Add("Branch");
            fixCol.Add("Status");
            fixCol.Add("TaxCode");
            fixCol.Add("NILetter");
            fixCol.Add("PreTaxAddDed");
            fixCol.Add("GrossedUpTaxThisRun");
            fixCol.Add("EeNIPdByEr");
            fixCol.Add("GUStudentLoan");
            fixCol.Add("GUNIReduction");
            fixCol.Add("PenPreTaxEeGU");
            fixCol.Add("TotalAbsencePay");
            fixCol.Add("HolidayPay");
            fixCol.Add("PenPreTaxEe");
            fixCol.Add("TaxablePay");
            fixCol.Add("Tax");
            fixCol.Add("NI");
            fixCol.Add("PostTaxAddDed");
            fixCol.Add("PostTaxPension");
            fixCol.Add("AEO");
            fixCol.Add("StudentLoan");
            fixCol.Add("NetPay");
            fixCol.Add("ErNI");
            fixCol.Add("PenEr");
            fixCol.Add("TotalGrossUp");

            return fixCol;
        }
        private List<string> CreateListOfVariableColumns(List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            //Create a list of the required variable columns.
            List<string> varCol = new List<string>();

            foreach (RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
            {
                if(rpPreSamplePayCode.Code != "TAX" && rpPreSamplePayCode.Code != "NI")
                {
                    if (rpPreSamplePayCode.InUse)
                    {
                        varCol.Add(rpPreSamplePayCode.Description);
                    }
                }
            }

            return varCol;
        }
        private Workbook CreateFixedWorkbookColumns(Workbook workbook, RPEmployeePeriod rpEmployeePeriod)
        {
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Reference);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Fullname);
            workbook.CurrentWorksheet.AddNextCell("Department");
            workbook.CurrentWorksheet.AddNextCell("Cost Centre");
            workbook.CurrentWorksheet.AddNextCell("Branch");
            workbook.CurrentWorksheet.AddNextCell("Calc");
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxCode);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NILetter);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxAddDed);
            workbook.CurrentWorksheet.AddNextCell(0.00);//GrossedUpTaxThisRun
            workbook.CurrentWorksheet.AddNextCell(0.00);//EeNIPdByEr
            workbook.CurrentWorksheet.AddNextCell(0.00);//GUStudentLoan
            workbook.CurrentWorksheet.AddNextCell(0.00);//GUNIReduction
            workbook.CurrentWorksheet.AddNextCell(0.00);//PenPreTaxEeGU
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AbsencePay);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.HolidayPay);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxPension);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxablePayTP);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Tax);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetNI);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxAddDed);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxPension);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AOE);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.StudentLoan);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetPayTP);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErNICTP);

            decimal erPensionTP = 0;
            foreach(RPPensionPeriod pensionPeriod in rpEmployeePeriod.Pensions)
            {
                erPensionTP = erPensionTP + pensionPeriod.ErPensionTaxPeriod;
            }
            workbook.CurrentWorksheet.AddNextCell(erPensionTP);
            workbook.CurrentWorksheet.AddNextCell(0.00);//TotalGrossUP
            
            return workbook;
        }
        private Workbook CreateVariableWorkbookColumns(Workbook workbook, RPEmployeePeriod rpEmployeePeriod, List<string> varCol)
        {
            foreach (string col in varCol)
            {
                //Add in the variable additions.
                bool colFound = false;
                foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
                {
                    if (col == rpAddition.Description)
                    {
                        workbook.CurrentWorksheet.AddNextCell(rpAddition.AmountTP);
                        colFound = true;
                        break;
                    }
                    
                }
                //If the column has not been found in additions check the variable deductions.
                if(!colFound)
                {
                    foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                    {
                        if (col == rpDeduction.Description)
                        {
                            workbook.CurrentWorksheet.AddNextCell(rpDeduction.AmountTP);
                            colFound = true;
                            break;
                        }

                    }
                    //If the column hasn't been found in additions or deduction set it to zero.
                    if (!colFound)
                    {
                        workbook.CurrentWorksheet.AddNextCell(0.00m);
                    }
                }
                
                

            }

            return workbook;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            btnProduceReports.PerformClick();
        }
        class PutToSFTP
        {
            public static string[] SFTPTransfer(string dataAddress, string strHostName, string strUserName, string strSSHPrivateKeyPath, bool isUnity)
            {
                //For locking of files transfer them up with a suffix of _filepart and when the transfer is complete remove the suffix.
                //I'll only set this to true if we start having problems with files being processed before they are fully uploaded.
                bool lockFiles = false;
                string suffix = "_filepart";
                string[] sftpReturn = new string[2];
                try
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = strHostName,    //"trans.bluemarblepayroll.com",
                        UserName = strUserName,    //"payescapetest",
                        Password = null,
                        PortNumber = 22,
                        SshHostKeyFingerprint = "ssh-rsa 2048 22:5f:d5:de:80:1d:52:69:72:55:3d:38:17:53:24:aa", //Old server  SshHostKeyFingerprint = "ssh-rsa 2048 f9:9e:38:ae:8d:55:d6:5d:f2:b3:63:67:e1:e4:d1:e1",
                        //JCBJCB
                        SshPrivateKeyPath = strSSHPrivateKeyPath    //"X:/jim/Documents/Payescape/Contracts/SFTP Private Key File/payescape.ppk"
                    };
                    //JCB TODO
                    using (Session session = new Session())
                    {
                        // Connect
                        session.Open(sessionOptions);

                        // Upload files
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;
                        transferOptions.ResumeSupport.State = TransferResumeSupportState.Off;
                        transferOptions.PreserveTimestamp = false;
                        transferOptions.FilePermissions = null; //This is the default

                        TransferOperationResult transferResult;
                        string outPath = dataAddress;

                        string destPath;
                        if (isUnity)
                        {
                            destPath = "../incoming/";
                        }
                        else
                        {
                            destPath = "../payescape/";
                        }

                        if (lockFiles)
                        {
                            transferResult = session.PutFiles(outPath, (destPath + "*.*" + suffix), false, transferOptions);

                        }
                        else
                        {
                            transferResult = session.PutFiles(outPath, destPath, false, transferOptions);

                        }


                        // Throw on any error
                        transferResult.Check();

                        //Rename uploaded files
                        if (lockFiles)
                        {
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                string finalName = transfer.Destination.Substring(0, transfer.Destination.Length - suffix.Length);
                                session.MoveFile(transfer.Destination, finalName);
                            }

                        }

                    }

                    sftpReturn[0] = "Success";
                    sftpReturn[1] = "Upload to SFTP Server successful.";
                    return sftpReturn;
                }
                catch (Exception ex)
                {

                    sftpReturn[0] = "Failure";
                    sftpReturn[1] = "Upload to SFTP Server Failed" + ex;
                    return sftpReturn;
                }
            }
        }
    }
    
}
