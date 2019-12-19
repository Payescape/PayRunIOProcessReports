﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using PicoXLSX;
using PayRunIOClassLibrary;
using System.Globalization;

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

            //
            //We'er going to change the way this done. Instead of PR producing folders with the EmployeePeriod & EmployeeYtd report already in them.
            //PR are going to give us an xml file to tell that the payroll has been done and the file will contain enough info to let us produce the required reports.
            //

            //FileInfo[] completedPayrollFiles = prWG.GetAllCompletedPayrollFiles(xdoc);
            //foreach (FileInfo completedPayrollFile in completedPayrollFiles)
            //{
            //    ReadProcessCompletedPayrollFile(xdoc, completedPayrollFile);
            //    //Put in some test for success then archive the file.
                
            //    prWG.ArchiveCompletedPayrollFile(xdoc, completedPayrollFile);
            //}



            //
            //This is the old method with folders containing the reports.
            //
            string[] directories = prWG.GetAListOfDirectories(xdoc);
            for (int i = 0; i < directories.Count(); i++)
            {
                try
                {
                    bool success = ProduceReports(xdoc, directories[i]);
                    if (success)
                    {
                        prWG.ArchiveDirectory(xdoc, directories[i]);
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
        }
        //private void ReadProcessCompletedPayrollFile(XDocument xdoc, FileInfo completedPayrollFile)
        //{
        //    PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

        //    XmlDocument xmlCompletedPayroll = new XmlDocument();
        //    xmlCompletedPayroll.Load(completedPayrollFile.FullName);

        //    //Now extract the necessary data and produce the required reports.

        //    RPParameters rpParameters = new RPParameters();
        //    foreach (XmlElement parameter in xmlCompletedPayroll.GetElementsByTagName("Parameters"))
        //    {
        //        rpParameters.ErRef = prWG.GetElementByTagFromXml(parameter, "EmployerCode");
        //        rpParameters.TaxYear = prWG.GetIntElementByTagFromXml(parameter, "TaxYear");
        //        rpParameters.AccYearStart = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(parameter, "AccountingYearStartDate"));
        //        rpParameters.AccYearEnd = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(parameter, "AccountingYearEndDate"));
        //        rpParameters.TaxPeriod = prWG.GetIntElementByTagFromXml(parameter, "TaxPeriod");
        //        rpParameters.PaySchedule = prWG.GetElementByTagFromXml(parameter, "PaySchedule");
        //    }
        //    GenerateReportsFromPR(xdoc, rpParameters);

        //}
        //private void GenerateReportsFromPR(XDocument xdoc, RPParameters rpParameters)
        //{
        //    PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
        //    //Produce and process Employee Period report. From this one report we are producing 5/6 standard pdf reports.
        //    //Get the history report
        //    string rptRef = "EEPERIOD";              //Original report name : "PayescapeEmployeePeriod"
        //    string parameter1 = "EmployerKey";
        //    string parameter2 = "TaxYear";
        //    string parameter3 = "AccPeriodStart";
        //    string parameter4 = "AccPeriodEnd";
        //    string parameter5 = "TaxPeriod";
        //    string parameter6 = "PayScheduleKey";

        //    //
        //    //Get the history report
        //    //

        //    XmlDocument xmlPeriodReport = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
        //                                      rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
        //                                      parameter6, rpParameters.PaySchedule.ToUpper());

        //    var tuple = ProcessPeriodReport(xdoc, xmlPeriodReport, rpParameters);
        //    RPEmployer rpEmployer = tuple.Item1;
        //    rpParameters = tuple.Item2;
            
        //    //
        //    //Produce and process Employee Ytd report.
        //    //

        //    rptRef = "EEYTD";              //Original report name : "PayescapeEmployeeYtd"
        //    XmlDocument xmlYTDReport = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
        //                                      rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
        //                                      parameter6, rpParameters.PaySchedule.ToUpper());

        //    prWG.ProcessYtdReport(xdoc, xmlYTDReport, rpParameters);

        //    //
        //    //Produce and process P45s if required. It is intended that PR will provide a list of employees who require a P45 within the completed payroll file.
        //    //

        //    //rptRef = "P45";
        //    //parameter2 = "EmployeeKey";
        //    //rpParameters.ErRef = "1176";
        //    //string eeRef = "14";
        //    //XmlDocument xmlP45Report = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, eeRef, null,
        //    //                                  null, null, null, null, null, null, null);

        //    //
        //    //Produce and process P32 if required. If the next pay run date gives us a different tax month than the current run date then we need to produce a P32 report.
        //    //
        //    bool p32Required = prWG.CheckIfP32Required(rpParameters);
        //    if(p32Required)
        //    {
        //        rptRef = "P32SUM";
        //        parameter2 = "TaxYear";
        //        XmlDocument xmlP32Report = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), null,
        //                                          null, null, null, null, null, null, null);
        //    }
            
        //    //
        //    //All the reports have been produced, so now just zip them and email them.
        //    //
        //    prWG.ZipReports(xdoc, rpEmployer, rpParameters);
        //    prWG.EmailZippedReports(xdoc, rpEmployer, rpParameters);


        //}
        //private Tuple<RPEmployer, RPParameters> ProcessPeriodReport(XDocument xdoc, XmlDocument xmlPeriodReport, RPParameters rpParameters)
        //{
        //    PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

        //    var tuple = PrepareStandardReports(xdoc, xmlPeriodReport, rpParameters);
        //    List<RPEmployeePeriod> rpEmployeePeriodList = tuple.Item1;
        //    List<RPPayComponent> rpPayComponents = tuple.Item2;
        //    //I don't think the P45 report will be able to be produced from the EmployeePeriod report but I'm leaving it here for now.
        //    List<P45> p45s = tuple.Item3;
        //    List<RPPreSamplePayCode> rpPreSamplePayCodes = tuple.Item4;
        //    RPEmployer rpEmployer = tuple.Item5;
        //    rpParameters = tuple.Item6;
        //    //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
        //    decimal hmrcTotal = prWG.CalculateHMRCTotal(rpEmployeePeriodList);
        //    rpEmployer.HMRCDesc = "[" + hmrcTotal.ToString() + "]";
        //    //I now have a list of employee with their total for this period & ytd plus addition & deductions
        //    //I can print payslips from here.
        //    //Test for 2 decimal place Units
        //    //foreach(RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
        //    //{
        //    //    foreach(RPAddition rpAddition in rpEmployeePeriod.Additions)
        //    //    {
        //    //        rpAddition.Units = 12.3456m;
        //    //    }
        //    //}
        //    prWG.PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents);

        //    //Create the history csv file from the objects
        //    prWG.CreateHistoryCSV(xdoc, rpParameters, rpEmployer, rpEmployeePeriodList);

        //    //Produce bank files if necessary
        //    prWG.ProcessBankReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);

        //    //Produce the Pre Sample Excel file (.xlsx)
        //    //We have a list of the Pre Sample Pay Codes which are in use and we have list of the employees paid this period. We should be able to produce the Pre Sample Excel file from these.
        //    CreatePreSampleXLSX(xdoc,rpEmployeePeriodList,rpEmployer,rpParameters,rpPreSamplePayCodes);

        //    //Create pension reports.

        //    return new Tuple<RPEmployer, RPParameters>(rpEmployer, rpParameters);

        //}
        private void ProducePeriodReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer,
                                          List<P45> p45s, List<RPPayComponent> rpPayComponents, RPParameters rpParameters,
                                          List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);

            string textLine = null;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
            decimal hmrcTotal = prWG.CalculateHMRCTotal(rpEmployeePeriodList);
            rpEmployer.HMRCDesc = "[" + hmrcTotal.ToString() + "]";
            //I now have a list of employee with their total for this period & ytd plus addition & deductions
            //I can print payslips and standard reports from here.
            try
            {
                prWG.PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error printing standard reports.\r\n", ex);
                prWG.update_Progress(textLine, softwareHomeFolder, logOneIn);
            }
            //Produce bank files if necessary
            try
            {
                prWG.ProcessBankReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
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
            

        }
        private Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>, RPEmployer, RPParameters> PrepareStandardReports(XDocument xdoc, XmlDocument xmlReport, RPParameters rpParameters)
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

            try
            {
                bool payRunDate = false;
                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;

                    if (prWG.GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                    {
                        if (!payRunDate)
                        {
                            rpParameters.PayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PayRunDate"));
                            payRunDate = true;
                        }
                        //If the employee is a leaver before the start date then don't include.
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
                        //TotalPayTP
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
                        rpEmployeePeriod.ErPensionYTD = prWG.GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                        rpEmployeePeriod.EePensionYTD = prWG.GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                        rpEmployeePeriod.ErPensionTP = prWG.GetDecimalElementByTagFromXml(employee, "ErPensionTaxPeriod");
                        rpEmployeePeriod.EePensionTP = prWG.GetDecimalElementByTagFromXml(employee, "EePensionTaxPeriod");
                        rpEmployeePeriod.ErContributionPercent = prWG.GetDecimalElementByTagFromXml(employee, "ErContributionPercent") * 100;
                        rpEmployeePeriod.EeContributionPercent = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionPercent") * 100;
                        rpEmployeePeriod.PensionablePay = prWG.GetDecimalElementByTagFromXml(employee, "PensionablePay");
                        rpEmployeePeriod.ErPensionPayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "ErPensionPayRunDate"));
                        rpEmployeePeriod.EePensionPayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "EePensionPayRunDate"));
                        rpEmployeePeriod.DirectorshipAppointmentDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "DirectorshipAppointmentDate"));
                        rpEmployeePeriod.Director = prWG.GetBooleanElementByTagFromXml(employee, "Director");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt1 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt1");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt2 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt2");
                        rpEmployeePeriod.ErNICTP = prWG.GetDecimalElementByTagFromXml(employee, "ErContributionTaxPeriod");
                        rpEmployeePeriod.Frequency = rpParameters.PaySchedule;
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
                                rpPayComponent.Surname = rpEmployeePeriod.Surname;
                                rpPayComponent.Rate = prWG.GetDecimalElementByTagFromXml(payCode, "Rate");
                                rpPayComponent.UnitsTP = prWG.GetDecimalElementByTagFromXml(payCode, "Units");
                                rpPayComponent.AmountTP = prWG.GetDecimalElementByTagFromXml(payCode, "Amount");
                                rpPayComponent.UnitsYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                rpPayComponent.AmountYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                //if (rpPayComponent.AmountTP != 0 || rpPayComponent.AmountYTD != 0)
                                if (rpPayComponent.AmountTP != 0)
                                {
                                    //Value is not equal to zero so go through the list of Pre Sample codes and mark this one as in use
                                    rpPreSamplePayCodes = MarkPreSampleCodeAsInUse(rpPayComponent.PayCode, rpPreSamplePayCodes);
                                    if (prWG.GetElementByTagFromXml(payCode, "IsPayCode") == "true")
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


                                if (prWG.GetElementByTagFromXml(payCode, "EarningOrDeduction") == "E")
                                {
                                    RPAddition rpAddition = new RPAddition();
                                    rpAddition.EeRef = rpEmployeePeriod.Reference;
                                    rpAddition.Code = prWG.GetElementByTagFromXml(payCode, "Code");
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
                                   
                                    rpAddition.Description = prWG.GetElementByTagFromXml(payCode, "Description");
                                    rpAddition.Rate = prWG.GetDecimalElementByTagFromXml(payCode, "Rate");
                                    rpAddition.Units = prWG.GetDecimalElementByTagFromXml(payCode, "Units");
                                    rpAddition.AmountTP = prWG.GetDecimalElementByTagFromXml(payCode, "Amount");
                                    rpAddition.AmountYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance");
                                    rpAddition.AccountsYearBalance = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance");
                                    rpAddition.AccountsYearUnits = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits");
                                    rpAddition.PayeYearUnits = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits");
                                    rpAddition.PayrollAccrued = prWG.GetDecimalElementByTagFromXml(payCode, "PayrollAccrued");
                                    //if (rpAddition.AmountTP != 0 || rpAddition.AmountYTD != 0)
                                    if (rpAddition.AmountTP != 0)
                                    {
                                        rpAdditions.Add(rpAddition);
                                        
                                    }
                                    rpEmployeePeriod.TotalPayTP = rpEmployeePeriod.TotalPayTP + rpAddition.AmountTP;
                                    rpEmployeePeriod.TotalPayYTD = rpEmployeePeriod.TotalPayYTD + rpAddition.AmountYTD;
                                }
                                else
                                {
                                    RPDeduction rpDeduction = new RPDeduction();
                                    rpDeduction.EeRef = rpEmployeePeriod.Reference;
                                    rpDeduction.Code = prWG.GetElementByTagFromXml(payCode, "Code");
                                    //They want Tax then NI, then Pension to come first, then the rest in alphabetical order. This will only work if they use the following codes!
                                    switch (rpDeduction.Code)
                                    {
                                        case "TAX":
                                            rpDeduction.Code = "   TAX";
                                            break;
                                        case "NI":
                                            rpDeduction.Code = "  NI";
                                            break;
                                        case "PENSION":
                                            rpDeduction.Code = " PENSION";
                                            break;
                                        case "PENSIONRAS":
                                            rpDeduction.Code = " PENSIONRAS";
                                            break;
                                        case "PENSIONSS":
                                            rpDeduction.Code = " PENSIONSS";
                                            break;
                                        case "PENSIONTAXEX":
                                            rpDeduction.Code = " PENSIONTAXEX";
                                            break;
                                    }
                                    rpDeduction.Description = prWG.GetElementByTagFromXml(payCode, "Description");
                                    rpDeduction.AmountTP = prWG.GetDecimalElementByTagFromXml(payCode, "Amount") * -1;
                                    rpDeduction.AmountYTD = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearBalance") * -1;
                                    rpDeduction.AccountsYearBalance = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearBalance") * -1;
                                    rpDeduction.AccountsYearUnits = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsYearUnits") * -1;
                                    rpDeduction.PayeYearUnits = prWG.GetDecimalElementByTagFromXml(payCode, "PayeYearUnits") * -1;
                                    rpDeduction.PayrollAccrued = prWG.GetDecimalElementByTagFromXml(payCode, "PayrollAccrued") * -1;
                                    //if (rpDeduction.AmountTP != 0 || rpDeduction.AmountYTD != 0)
                                    if (rpDeduction.AmountTP != 0)
                                    {
                                        rpDeductions.Add(rpDeduction);
                                        
                                    }
                                    rpEmployeePeriod.TotalDedTP = rpEmployeePeriod.TotalDedTP + rpDeduction.AmountTP;
                                    rpEmployeePeriod.TotalDedYTD = rpEmployeePeriod.TotalDedYTD + rpDeduction.AmountYTD;
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
                                    if (x.Code == null && y.Code == null) return 0;
                                    else if (x.Code == null) return -1;
                                    else if (y.Code == null) return 1;
                                    else return x.Code.CompareTo(y.Code);
                                });
                                rpEmployeePeriod.Additions = rpAdditions;
                                rpEmployeePeriod.Deductions = rpDeductions;
                            }//End of for each payCode
                            preSamplePayCodes = true;
                        }//End of for each payCodes
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

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error preparing reports.\r\n{0}.\r\n", ex);
                prWG.update_Progress(textLine, configDirName, logOneIn);
            }
            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>, RPEmployer, RPParameters>(rpEmployeePeriodList, rpPayComponents, p45s, rpPreSamplePayCodes, rpEmployer, rpParameters);

        }
        //rpPreSamplePayCodes = MarkPreSampleCodeAsInUse(rpPayComponent.PayCode, rpPreSamplePayCodes);
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
                    RPEmployer rpEmployer = null;
                    
                    try
                    {
                        var tuple = PreparePeriodReport(xdoc, file);
                        rpEmployeePeriodList = tuple.Item1;
                        rpPayComponents = tuple.Item2;
                        p45s = tuple.Item3;
                        rpPreSamplePayCodes = tuple.Item4;
                        rpEmployer = tuple.Item5;
                        rpParameters = tuple.Item6;
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
                        ProducePeriodReports(xdoc, rpEmployeePeriodList, rpEmployer, p45s, rpPayComponents, rpParameters, rpPreSamplePayCodes);

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
        private Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>, RPEmployer, RPParameters> PreparePeriodReport(XDocument xdoc, FileInfo file)
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
            RPEmployer rpEmployer = tuple.Item5;
            rpParameters = tuple.Item6;
            //Test for 2 decimal place Units
            //foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            //{
            //    foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
            //    {
            //        rpAddition.Units = 12.3446m;
            //    }
            //}

            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>, RPEmployer, RPParameters>(rpEmployeePeriodList, rpPayComponents, p45s, rpPreSamplePayCodes, rpEmployer, rpParameters);

        }
        public void CreatePreSampleXLSX(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
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
        //private void CreateXLSXWorkbook()
        //{
        //    //Create a list of the required columns.
        //    List <string> reqCol = new List<string>();
        //    reqCol.Add("EeRef");
        //    reqCol.Add("Name");
        //    reqCol.Add("Dept");
        //    reqCol.Add("CostCentre");
        //    reqCol.Add("Branch");
        //    reqCol.Add("Status");
        //    reqCol.Add("TaxCode");
        //    reqCol.Add("NILetter");
        //    reqCol.Add("PreTaxAddDed");
        //    reqCol.Add("GrossedUpTaxThisRun");
        //    reqCol.Add("EeNIPdByEr");
        //    reqCol.Add("GUStudentLoan");
        //    reqCol.Add("GUNIReduction");
        //    reqCol.Add("PenPreTaxEeGU");
        //    reqCol.Add("TotalAbsencePay");
        //    reqCol.Add("HolidayPay");
        //    reqCol.Add("PenPreTaxEe");
        //    reqCol.Add("TaxablePay");
        //    reqCol.Add("Tax");
        //    reqCol.Add("NI");
        //    reqCol.Add("PostTaxAddDed");
        //    reqCol.Add("PostTaxPension");
        //    reqCol.Add("AEO");
        //    reqCol.Add("StudentLoan");
        //    reqCol.Add("NetPay");
        //    reqCol.Add("ErNI");
        //    reqCol.Add("PenEr");
        //    reqCol.Add("TotalGrossUp");

            
        //    //Need to count how many columns we are going to need
        //    string[] headings = new string[51] { "EeRef", "Name", "Dept","CostCentre", "Branch", "Status", "TaxCode", "NILetter", "PreTaxAddDed",
        //                                         "GrossedUpTaxThisRun", "EeNIPdByEr", "GUStudentLoan", "GUNIReduction", "PenPreTaxEeGU", "TotalAbsencePay",
        //                                         "HolidayPay", "PenPreTaxEe", "TaxablePay", "Tax", "NI", "PostTaxAddDed", "PostTaxPension", "AEO",
        //                                         "StudentLoan", "NetPay", "ErNI", "PenEr", "TotalGrossUp", "SSP", "SMP", "SAP", "SPPA", "SPPB", "ASPPA",
        //                                         "ASPPB", "ShPPA", "ShPPB", "TotalNICs", "TotalPens", "BIK", "BasicPay", "PerformanceRelatedPay",
        //                                         "Salary(£)", "HolidayHours", "OverPayment(£)", "Bonus", "CycleToWorkScheme(£)", "HattonGroupScheme(Er)",
        //                                         "HattonGroupScheme", "CCAEO", "DEA"};
        //    string[] columns = new string[51] { "E1234", "Jim Borland", "Automation","Software Developer", "Dromore", "Calc", "1238L", "A", "432.10",
        //                                         "32.10", "12.34", "12.35", "12.36", "12.37", "12.38",
        //                                         "12.39", "12.40", "12.41", "12.42", "12.43", "12.44", "12.45", "12.46",
        //                                         "12.47", "12.48", "12.49", "12.50", "12.51", "12.52", "12.53", "12.54", "12.55", "12.56", "12.57",
        //                                         "12.58", "12.59", "12.60", "12.61", "12.62", "12.63", "12.64", "12.65",
        //                                         "12.66", "12.67", "12.68", "12.69", "12.70", "12.71",
        //                                         "12.72", "12.73", "12.74"};
        //    //Create a workbook.
        //    Workbook workbook = new Workbook("X:\\Payescape\\PayRunIO\\PreSample.xlsx", "Pre Sample");
        //    //Write the headings.
        //    foreach(string heading in headings)
        //    {
        //        workbook.CurrentWorksheet.AddNextCell(heading);
        //    }
        //    //Move to the next row.
        //    workbook.CurrentWorksheet.GoToNextRow();
        //    //Now create a sample data line.
        //    foreach (string column in columns)
        //    {
        //        workbook.CurrentWorksheet.AddNextCell(column);
        //    }
        //    //Save the workbook.
        //    workbook.Save();
        //}
        private void CreatePreSampleXLSX(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList,
                                       RPEmployer rpEmployer, RPParameters rpParameters, List<RPPreSamplePayCode> rpPreSamplePayCodes)
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

            foreach (RPPreSamplePayCode rpPreSamplePayCode in rpPreSamplePayCodes)
            {
                if(rpPreSamplePayCode.InUse)
                {
                    reqCol.Add(rpPreSamplePayCode.Description);
                }
                
            }

            //Create a workbook.
            Workbook workbook = new Workbook("X:\\Payescape\\PayRunIO\\PreSample.xlsx", "Pre Sample");
            foreach (string col in reqCol)
            {
                workbook.CurrentWorksheet.AddNextCell(col);
            }
            
            //Now for each employee create a row and add in the values for each column
            foreach(RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                workbook.CurrentWorksheet.GoToNextRow();
                ////Check for the different pay codes and add to the appropriate total.
                //switch (rpPayComponent.PayCode)
                //{
                //    case "HOLPY":
                //    case "HOLIDAY":
                //        rpEmployeePeriod.HolidayPay = rpEmployeePeriod.HolidayPay + rpPayComponent.AmountTP;
                //        break;
                //    case "PENSION":
                foreach (string col in reqCol)
                {
                    switch (col)
                    {
                        case "EeRef":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Reference);
                            break;
                        case "Name":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Fullname);
                            break;
                        case "Dept":
                            workbook.CurrentWorksheet.AddNextCell("Department");
                            break;
                        case "CostCentre":
                            workbook.CurrentWorksheet.AddNextCell("Cost Centre");
                            break;
                        case "Branch":
                            workbook.CurrentWorksheet.AddNextCell("Branch");
                            break;
                        case "Status":
                            workbook.CurrentWorksheet.AddNextCell("Calc");
                            break;
                        case "TaxCode":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxCode);
                            break;
                        case "NILetter":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NILetter);
                            break;
                        case "PreTaxAddDed":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxAddDed);
                            break;
                        case "GrossedUpTaxThisRun":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//GrossedUpTaxThisRun
                            break;
                        case "EeNIPdByEr":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//EeNIPdByEr
                            break;
                        case "GUStudentLoan":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//GUStudentLoan
                            break;
                        case "GUNIReduction":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//GUNIReduction
                            break;
                        case "PenPreTaxEeGU":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//PenPreTaxEeGU
                            break;
                        case "TotalAbsencePay":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AbsencePay);
                            break;
                        case "HolidayPay":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.HolidayPay);
                            break;
                        case "PenPreTaxEe":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PreTaxPension);
                            break;
                        case "TaxablePay":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.TaxablePayTP);
                            break;
                        case "Tax":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.Tax);
                            break;
                        case "NI":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetNI);
                            break;
                        case "PostTaxAddDed":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxAddDed);
                            break;
                        case "PostTaxPension":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.PostTaxPension);
                            break;
                        case "AOE":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AOE);
                            break;
                        case "StudentLoan":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.StudentLoan);
                            break;
                        case "NetPay":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetPayTP);
                            break;
                        case "ErNI":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErNICTP);
                            break;
                        case "PenEr":
                            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErPensionTP);
                            break;
                        case "TotalGrossUp":
                            workbook.CurrentWorksheet.AddNextCell(0.00);//TotalGrossUP
                            break;
                        default:
                            //It's none of the set headings so now look through the pay codes (additions & deductions) to see if we can find it
                            foreach(RPAddition rpAddition in rpEmployeePeriod.Additions)
                            {
                                if(col == rpAddition.Description)
                                {
                                    workbook.CurrentWorksheet.AddNextCell(rpAddition.AmountTP);
                                    break;
                                }
                            }
                            break;
                    }
                }

                //Now go through each pay code that has a value and add it in.
                foreach(RPAddition rpAddition in rpEmployeePeriod.Additions)
                {

                }
            }
            
            //Now create a sample data line.
            //foreach (string column in columns)
            //{
            //    workbook.CurrentWorksheet.AddNextCell(column);
            //}
            //Save the workbook.
            workbook.Save();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //CreateXLSXWorkbook();
            btnProduceReports.PerformClick();
        }
    }
    
}
