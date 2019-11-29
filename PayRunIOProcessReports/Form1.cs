using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using PayRunIOClassLibrary;

namespace PayRunIOProcessReports
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Tuple<int,int> TupleTest()
        {
            return new Tuple<int,int>(0,0);
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

            FileInfo[] completedPayrollFiles = prWG.GetAllCompletedPayrollFiles(xdoc);
            foreach (FileInfo completedPayrollFile in completedPayrollFiles)
            {
                ReadProcessCompletedPayrollFile(xdoc, completedPayrollFile);
                //Put in some test for success then archive the file.
                
                prWG.ArchiveCompletedPayrollFile(xdoc, completedPayrollFile);
            }



            //
            //This is the old method with folders containing the reports.
            //
            string[] directories = prWG.GetAListOfDirectories(xdoc);
            for (int i = 0; i < directories.Count(); i++)
            {
                try
                {
                    bool success = prWG.ProduceReports(xdoc, directories[i]);
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
        private void ReadProcessCompletedPayrollFile(XDocument xdoc, FileInfo completedPayrollFile)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            XmlDocument xmlCompletedPayroll = new XmlDocument();
            xmlCompletedPayroll.Load(completedPayrollFile.FullName);

            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = new RPParameters();
            foreach (XmlElement parameter in xmlCompletedPayroll.GetElementsByTagName("Parameters"))
            {
                rpParameters.ErRef = prWG.GetElementByTagFromXml(parameter, "EmployerCode");
                rpParameters.TaxYear = prWG.GetIntElementByTagFromXml(parameter, "TaxYear");
                rpParameters.AccYearStart = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(parameter, "AccountingYearStartDate"));
                rpParameters.AccYearEnd = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(parameter, "AccountingYearEndDate"));
                rpParameters.TaxPeriod = prWG.GetIntElementByTagFromXml(parameter, "TaxPeriod");
                rpParameters.PaySchedule = prWG.GetElementByTagFromXml(parameter, "PaySchedule");
            }
            GenerateReportsFromPR(xdoc, rpParameters);

        }
        private void GenerateReportsFromPR(XDocument xdoc, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Produce and process Employee Period report. From this one report we are producing 5/6 standard pdf reports.
            //Get the history report
            string rptRef = "EEPERIOD";              //Original report name : "PayescapeEmployeePeriod"
            string parameter1 = "EmployerKey";
            string parameter2 = "TaxYear";
            string parameter3 = "AccPeriodStart";
            string parameter4 = "AccPeriodEnd";
            string parameter5 = "TaxPeriod";
            string parameter6 = "PayScheduleKey";

            //
            //Get the history report
            //

            XmlDocument xmlPeriodReport = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
                                              rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
                                              parameter6, rpParameters.PaySchedule.ToUpper());

            var tuple = ProcessPeriodReport(xdoc, xmlPeriodReport, rpParameters);
            RPEmployer rpEmployer = tuple.Item1;
            rpParameters = tuple.Item2;
            
            //
            //Produce and process Employee Ytd report.
            //

            rptRef = "EEYTD";              //Original report name : "PayescapeEmployeeYtd"
            XmlDocument xmlYTDReport = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), parameter3,
                                              rpParameters.AccYearStart.ToString("yyyy-MM-dd"), parameter4, rpParameters.AccYearEnd.ToString("yyyy-MM-dd"), parameter5, rpParameters.TaxPeriod.ToString(),
                                              parameter6, rpParameters.PaySchedule.ToUpper());

            prWG.ProcessYtdReport(xdoc, xmlYTDReport, rpParameters);

            //
            //Produce and process P45s if required. It is intended that PR will provide a list of employees who require a P45 within the completed payroll file.
            //

            //rptRef = "P45";
            //parameter2 = "EmployeeKey";
            //rpParameters.ErRef = "1176";
            //string eeRef = "14";
            //XmlDocument xmlP45Report = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, eeRef, null,
            //                                  null, null, null, null, null, null, null);

            //
            //Produce and process P32 if required. If the next pay run date gives us a different tax month than the current run date then we need to produce a P32 report.
            //
            bool p32Required = prWG.CheckIfP32Required(rpParameters);
            if(p32Required)
            {
                rptRef = "P32SUM";
                parameter2 = "TaxYear";
                XmlDocument xmlP32Report = prWG.RunReport(rptRef, parameter1, rpParameters.ErRef, parameter2, rpParameters.TaxYear.ToString(), null,
                                                  null, null, null, null, null, null, null);
            }
            
            //
            //All the reports have been produced, so now just zip them and email them.
            //
            prWG.ZipReports(xdoc, rpEmployer, rpParameters);
            prWG.EmailZippedReports(xdoc, rpEmployer, rpParameters);


        }
        public Tuple<RPEmployer, RPParameters> ProcessPeriodReport(XDocument xdoc, XmlDocument xmlPeriodReport, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            var tuple = prWG.PrepareStandardReports(xdoc, xmlPeriodReport, rpParameters);
            List<RPEmployeePeriod> rpEmployeePeriodList = tuple.Item1;
            List<RPPayComponent> rpPayComponents = tuple.Item2;
            //I don't think the P45 report will be able to be produced from the EmployeePeriod report but I'm leaving it here for now.
            List<P45> p45s = tuple.Item3;
            RPEmployer rpEmployer = tuple.Item4;
            rpParameters = tuple.Item5;
            //Get the total payable to hmrc, I'm going use it in the zipped file name(possibly!).
            decimal hmrcTotal = prWG.CalculateHMRCTotal(rpEmployeePeriodList);
            rpEmployer.HMRCDesc = "[" + hmrcTotal.ToString() + "]";
            //I now have a list of employee with their total for this period & ytd plus addition & deductions
            //I can print payslips from here.
            prWG.PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents);

            //Create the history csv file from the objects
            prWG.CreateHistoryCSV(xdoc, rpParameters, rpEmployer, rpEmployeePeriodList);

            //Produce bank files if necessary
            prWG.ProcessBankReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);

            //Produce the Pre Sample Excel file (.xlsx)
            //prWG.CreatePreSampleXLSX(xdoc,rpEmployeePeriodList,rpEmployer,rpParameters);

            //Create pension reports.

            return new Tuple<RPEmployer, RPParameters>(rpEmployer, rpParameters);

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            btnProduceReports.PerformClick();
        }
    }
    
}
