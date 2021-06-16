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
using DevExpress.XtraReports.UI;
using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
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
            string dirName = configDirName;
            ReadConfigFile configFile = new ReadConfigFile();
            XDocument xdoc = configFile.ConfigRecord(dirName);
            dirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            configDirName = dirName;
            
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            
            // Scan the folder and upload files waiting there.
            string textLine = string.Format("Starting from called program (PayRunIOProcessReports).");
            prWG.Update_Progress(textLine, configDirName);
            
            
            //Start by updating the contacts table
            prWG.UpdateContactDetails(xdoc);

            
            //Now process the reports
            ProcessReportsFromPayRunIO(xdoc);

            
            Close();

        }
        private void ProcessReportsFromPayRunIO(XDocument xdoc)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);

            string textLine;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            textLine = string.Format("Start processing the Outputs folder.");
            prWG.Update_Progress(textLine, softwareHomeFolder);
            string originalDirName = "Outputs";
            string archiveDirName = "PE-ArchivedOutputs";
            string[] directories = prWG.GetAListOfDirectories(xdoc, "Outputs");

            for (int i = 0; i < directories.Count(); i++)
            {
                try
                {
                    bool success = ProcessOutputFiles(xdoc, directories[i]);
                    if (success)
                    {
                        archiveDirName = "PE-ArchivedOutputs";
                    }
                    else
                    {
                        archiveDirName = "PE-FailedOutputs";
                    }
                    try
                    {
                        prWG.ArchiveDirectory(xdoc, directories[i], originalDirName, archiveDirName);
                    }
                    catch(Exception ex)
                    {
                        textLine = string.Format("Error archiving Outputs folder for directory {0}.\r\n{1}.\r\n", directories[i], ex);
                        prWG.Update_Progress(textLine, softwareHomeFolder);
                    }
                    

                }
                catch (Exception ex)
                {
                    textLine = string.Format("Error processing the reports for directory {0}.\r\n{1}.\r\n", directories[i], ex);
                    prWG.Update_Progress(textLine, softwareHomeFolder);
                }

            }

            textLine = string.Format("Finished processing the reports.");
            prWG.Update_Progress(textLine, softwareHomeFolder);

            //Transfer the contents of PE-Outgoing up to the Blue Marble SFTP server
            //Each company has it's own folder beneath PE-Outgoing which is just named with their company number "_" pay date.
            textLine = string.Format("Start processing the PE-Outgoing directory.");
            prWG.Update_Progress(textLine, softwareHomeFolder);

            originalDirName = "PE-Outgoing";
            archiveDirName = "PE-Outgoing\\Archive";
            directories = prWG.GetAListOfDirectories(xdoc, "PE-Outgoing");



            int nowDate = Convert.ToInt32(DateTime.Now.ToString("yyyyMMdd"));
            int payDate;
            int start, length, diff;

            for (int i = 0; i < directories.Count(); i++)
            {
                if(!directories[i].Contains("Archive"))
                {
                    bool upload = true;
                    //The directory name contains the pay date in the form yyyMMdd. It's coNo_PayDate e.g. 1880_20200528
                    //Because of a bug in Blue Marble I'm not uploading the files until it's within 1 day
                    //Emer wants to delay the upload of the pay history and ytd files until the pay date has been reached. This is from the folder PE-Outgoing
                    //I'm going to use the pay date in the file name and then I can compare for it.
                    //If it's not the live system this does not apply.
                    if(live)
                    {
                        start = directories[i].LastIndexOf("PE-Outgoing\\") + 17;
                        length = 8;
                        payDate = Convert.ToInt32(directories[i].Substring(start, length));
                        diff = payDate - nowDate;
                        if (diff > 1)
                        {
                            upload = false;
                        }
                    }
                    if(upload)
                    {
                        try
                        {
                            bool success;
                            //Emer doesn't want the PayHistory file being uploaded to Blue Marble for the test server anymore.
                            if(live)
                            {
                                success = TransferToBlueMarbleSFTPServer(xdoc, directories[i]);            // Transfer contents of the folder up to Blue Marble SFTP server.//ProduceReports(xdoc, directories[i]);
                            }
                            else
                            {
                                success = true;
                            }
                            
                            if (success)
                            {
                                //Copy the folder up to the S3 server before archiving it.
                                UploadPEOutgoingFilesToAmazonS3(xdoc, directories[i]);

                                try
                                {
                                    textLine = string.Format("Trying to archive directory {0}.", directories[i]);
                                    prWG.Update_Progress(textLine, softwareHomeFolder);

                                    prWG.ArchiveDirectory(xdoc, directories[i], originalDirName, archiveDirName);
                                }
                                catch (Exception ex)
                                {
                                    textLine = string.Format("Error archiving directory {0}.\r\n{1}.\r\n", directories[i], ex);
                                    prWG.Update_Progress(textLine, softwareHomeFolder);
                                }

                            }


                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Error processing PE-Outgoing folder for directory {0}.\r\n{1}.\r\n", directories[i], ex);
                            prWG.Update_Progress(textLine, softwareHomeFolder);
                        }
                    }
                }

            }

        }
        private bool TransferToBlueMarbleSFTPServer(XDocument xdoc, string directory)
        {
            string softwareHomeFolder = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string sftpHostName = xdoc.Root.Element("SFTPHostName").Value;
            string user = xdoc.Root.Element("User").Value;
            //bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);
            //if(!live)
            //{
            //    user = "payruntest123";//For testing purposes
            //}
            
            string passwordFile = softwareHomeFolder + "Programs\\" +xdoc.Root.Element("PasswordFile").Value;
            string textLine;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            textLine = string.Format("Start processing the reports.");
            prWG.Update_Progress(textLine, softwareHomeFolder);

            bool success = true;

            //
            // Don't think Blue Marble can cope with a directory so go through each file individually
            //
            bool isUnity = true;

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
                        prWG.Update_Progress(textLine, softwareHomeFolder);
                    }
                    else
                    {
                        //
                        // SFTP failed
                        //
                        textLine = sftpReturn[1];
                        prWG.Update_Progress(textLine, softwareHomeFolder);
                        success = false;
                    }

                }
                catch
                {
                    textLine = string.Format("Failed to upload zipped csv file {0} for client reference : {1}", csvFile, companyNo);
                    prWG.Update_Progress(textLine, softwareHomeFolder);
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
            string textLine;

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            
            //I now have a list of employee with their total for this period & ytd plus addition & deductions
            //I can print payslips and standard reports from here.
            try
            {
                PrintStandardReports(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters, p45s, rpPayComponents, rpPensionContributions);
                if (rpEmployer.P32Required)
                {
                    RPP32Report rpP32Report = CreateP32Report(xdoc, rpEmployer, rpParameters);
                    PrintP32Report(xdoc, rpP32Report, rpParameters, rpEmployer);

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
                if(rpEmployer.CalculateApprenticeshipLevy)
                {
                    PrintApprenticeshipLevyReport(xdoc, rpParameters, rpEmployer);
                }
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error printing standard reports.\r\n", ex);
                prWG.Update_Progress(textLine, softwareHomeFolder);
            }
            
            //Produce bank files if necessary
            try
            {
                ProcessBankAndPensionFiles(xdoc, rpEmployeePeriodList, rpPensionContributions, rpEmployer, rpParameters);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error processing bank reports.\r\n", ex);
                prWG.Update_Progress(textLine, softwareHomeFolder);
            }

            
            //Produce Pre Sample file (XLSX)
            CreatePreSampleXLSX(xdoc, rpEmployeePeriodList, rpParameters, rpPreSamplePayCodes);

            try
            {
                ZipReports(xdoc, rpEmployer, rpParameters);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error zipping reports.\r\n", ex);
                prWG.Update_Progress(textLine, softwareHomeFolder);
            }
            try
            {
                //Check if the reports should be zipped or not.
                if(rpEmployer.ZipReports)
                {
                    EmailZippedReports(xdoc, rpEmployer, rpParameters);
                }
                else
                {
                    EmailUnZippedReports(xdoc, rpEmployer, rpParameters);
                }
                string peReportsFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef;
                prWG.DeleteFilesThenFolder(xdoc, peReportsFolder);

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped reports.\r\n", ex);
                prWG.Update_Progress(textLine, softwareHomeFolder);
            }
            try
            {
                UploadZippedReportsToAmazonS3(xdoc);
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error uploading zipped reports to Amazon S3.\r\n", ex);
                prWG.Update_Progress(textLine, softwareHomeFolder);
            }

        }
       
        public void EmailZippedReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(reportFolder);
                FileInfo[] files = dirInfo.GetFiles(rpParameters.ErRef + "*.*");

                EmailReports(xdoc, files, rpParameters, rpEmployer);
                

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped pdf reports for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                prWG.Update_Progress(textLine, configDirName);
            }
        }
        public void EmailUnZippedReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            

            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef;
            DirectoryInfo dirInfo = new DirectoryInfo(reportFolder);
            FileInfo[] files = dirInfo.GetFiles();

            try
            {
                //EmailUnZippedReportFiles(xdoc, files, rpParameters, rpEmployer);
                EmailReports(xdoc, files, rpParameters, rpEmployer);
            }
            catch (Exception ex)
            {
                textLine = string.Format("Error emailing zipped pdf reports for report folder, {0}.\r\n{1}.\r\n", reportFolder, ex);
                prWG.Update_Progress(textLine, configDirName);
            }
            
        }
        public void UploadZippedReportsToAmazonS3(XDocument xdoc)
        {
            string reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-ReportsNoPassword";
            string amazonS3Folder = "PE-ReportsNoPassword";
            //Upload PE-ReportsNoPassword
            UploadDirectoryToAmazonS3(reportFolder, xdoc, amazonS3Folder);
            //Upload PE-Reports
            amazonS3Folder = "PE-Reports";
            reportFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            UploadDirectoryToAmazonS3(reportFolder, xdoc, amazonS3Folder);
                
            
        }
        public void UploadDirectoryToAmazonS3(string directory, XDocument xdoc, string amazonS3Folder)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            try
            {
                //Upload directory to Amazon S3
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    UploadZippedReportToAmazonS3(xdoc, file, amazonS3Folder);
                    file.MoveTo(file.FullName.Replace(directory, directory + "\\Archive"));
                }
            }
            catch(Exception ex)
            {
                textLine = string.Format("Error uploading zipped pdf reports to Amazon S3 for report folder, {0}.\r\n{1}.\r\n", directory, ex);
                prWG.Update_Progress(textLine, configDirName);
            }
            
        }
        private void UploadZippedReportToAmazonS3(XDocument xdoc, FileInfo file, string folderPath)
        {
            string awsBucketName = xdoc.Root.Element("AwsBucketName").Value;
            string awsAccessKey = xdoc.Root.Element("AwsAccessKey").Value;
            string awsAccessSecret = xdoc.Root.Element("AwsAccessSecret").Value;
            bool awsInDevelopment = Convert.ToBoolean(xdoc.Root.Element("InDevelopment").Value);
            
            bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);
            string bucketName = awsBucketName;
            IAmazonS3 s3Client;
            if (awsInDevelopment)
            {
                s3Client = new AmazonS3Client(awsAccessKey, awsAccessSecret, RegionEndpoint.EUWest2);
            }
            else
            {
                s3Client = new AmazonS3Client(RegionEndpoint.EUWest2);
            }
            if (live)
            {
                folderPath += "Live/";
            }
            else
            {
                folderPath += "Test/";
            }
            try
            {
                PutObjectRequest request = new PutObjectRequest()
                {
                    
                    InputStream = file.OpenRead(),
                    BucketName = bucketName,
                    Key = folderPath + file.ToString()
                };
                request.AutoCloseStream = true;
                
                PutObjectResponse response = s3Client.PutObject(request);
                
            }
            catch
            {

            }
            finally
            {

            }
      
            
            
        }
        public void UploadPEOutgoingFilesToAmazonS3(XDocument xdoc, string directory)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                FileInfo[] files = dirInfo.GetFiles();
                foreach (FileInfo file in files)
                {
                    UploadPEOutgoingFileToAmazonS3(xdoc, file);
                    
                }

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error uploading PEOutgoing file to Amazon S3 for report folder, {0}.\r\n{1}.\r\n", directory, ex);
                prWG.Update_Progress(textLine, configDirName);
            }
        }
        private void UploadPEOutgoingFileToAmazonS3(XDocument xdoc, FileInfo file)
        {
            string awsBucketName = xdoc.Root.Element("AwsBucketName").Value;
            string awsAccessKey = xdoc.Root.Element("AwsAccessKey").Value;
            string awsAccessSecret = xdoc.Root.Element("AwsAccessSecret").Value;
            bool awsInDevelopment = Convert.ToBoolean(xdoc.Root.Element("InDevelopment").Value);

            bool live = Convert.ToBoolean(xdoc.Root.Element("Live").Value);
            string bucketName = awsBucketName;
            //RegionEndpoint bucketRegion = RegionEndpoint.EUWest2;
            IAmazonS3 s3Client;
            if (awsInDevelopment)
            {
                s3Client = new AmazonS3Client(awsAccessKey, awsAccessSecret, RegionEndpoint.EUWest2);
            }
            else
            {
                s3Client = new AmazonS3Client(RegionEndpoint.EUWest2);
            }
            string folderPath;
            if (live)
            {
                folderPath = "PE-OutgoingLive/";
            }
            else
            {
                folderPath = "PE-OutgoingTest/";
            }

            PutObjectRequest request = new PutObjectRequest()
            {
                InputStream = file.OpenRead(),
                BucketName = bucketName,
                Key = folderPath + file.ToString()
            };
            PutObjectResponse response = s3Client.PutObject(request);

        }
        private void EmailReports(XDocument xdoc, FileInfo[] files, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            try
            {
                //
                // Send an email.
                //
                bool validEmailAddress = false;
                string hmrcDesc = null;
                if (rpEmployer.P32Required)
                {
                    //hmrcDesc = file.FullName.Substring(x, y - x);
                    hmrcDesc = rpEmployer.HMRCDesc;
                    hmrcDesc = hmrcDesc.Replace("[", "£");
                    hmrcDesc = hmrcDesc.Replace("]", "");
                }

                DateTime runDate = rpParameters.PayRunDate;
                runDate = runDate.AddDays(-5);
                runDate = runDate.AddMonths(1);
                int day = runDate.Day;
                day = 20 - day;
                runDate = runDate.AddDays(day);
                string dueDate = runDate.ToLongDateString();
                string taxYear = rpParameters.TaxYear.ToString() + "/" + (rpParameters.TaxYear + 1).ToString().Substring(2, 2);
                string mailSubject = String.Format("Payroll reports for {0}, for tax year {1}, pay period {2} ({3}).", rpEmployer.Name, taxYear, rpParameters.PeriodNo, rpParameters.PaySchedule);
                string mailBody = null;

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
                smtpEmailSettings = prWG.GetEmailSettings(xdoc, sqlConnectionString);
                //
                //Get a list of email addresses to send the reports to
                //
                List<ContactInfo> contactInfoList = new List<ContactInfo>();
                contactInfoList = prWG.GetListOfContactInfo(xdoc, sqlConnectionString, rpParameters);
                mailBody = String.Format("Hi All,\r\n\r\nPlease find attached payroll reports for {0}, for tax year {1}, pay period {2} ({3}).\r\n\r\n"
                                                 , rpEmployer.Name, taxYear, rpParameters.PeriodNo, rpParameters.PaySchedule);
                if (rpEmployer.P32Required)
                {
                    mailBody += string.Format("The amount payable to HMRC this month is {0}, this payment is due on or before {1}.\r\n\r\n"
                                                            , hmrcDesc, dueDate);
                }
                string approveBy = null;
                switch (rpParameters.PaySchedule)
                {
                    case "Weekly":
                        approveBy = "by 12 noon on Wednesday";
                        break;
                    case "Monthly":
                        approveBy = "by the 15th of this month";
                        break;
                    default:
                        approveBy = "as soon as possible";
                        break;
                }
                mailBody += string.Format("Please review and confirm if all is correct {0}.\r\n\r\nKind Regards,\r\n\r\nThe Payescape Team.", approveBy);


                //Start of MailMessage
                using (MailMessage mailMessage = new MailMessage())
                {
                    string firstEmailAddress = null;
                    foreach (ContactInfo contactInfo in contactInfoList)
                    {
                        RegexUtilities regexUtilities = new RegexUtilities();
                        bool isValidEmailAddress = regexUtilities.IsValidEmail(contactInfo.EmailAddress);
                        if (isValidEmailAddress)
                        {
                            mailMessage.To.Add(new MailAddress(contactInfo.EmailAddress));
                            if (!validEmailAddress)
                            {
                                validEmailAddress = true;
                                firstEmailAddress = contactInfo.EmailAddress;
                            }

                        }

                    }
                    if(validEmailAddress)
                    {
                        mailMessage.From = new MailAddress(smtpEmailSettings.FromAddress);
                        mailMessage.Subject = mailSubject;
                        mailMessage.Body = mailBody;

                        foreach (FileInfo file in files)
                        {
                            Attachment attachment = new Attachment(file.FullName);
                            mailMessage.Attachments.Add(attachment);


                        }

                        SmtpClient smtpClient = new SmtpClient()
                        {
                            UseDefaultCredentials = smtpEmailSettings.SMTPUserDefaultCredentials,
                            Credentials = new System.Net.NetworkCredential(smtpEmailSettings.SMTPUsername, smtpEmailSettings.SMTPPassword),
                            Port = smtpEmailSettings.SMTPPort,
                            Host = smtpEmailSettings.SMTPHost,
                            EnableSsl = smtpEmailSettings.SMTPEnableSSL
                        };
                        

                        try
                        {
                            textLine = string.Format("Attempting sending an email to, {0} from {1} with password:{2}, port:{3}, host:{4}.", firstEmailAddress,
                                                        smtpEmailSettings.SMTPUsername, smtpEmailSettings.SMTPPassword, smtpEmailSettings.SMTPPort, smtpEmailSettings.SMTPHost);
                            prWG.Update_Progress(textLine, configDirName);

                            smtpClient.Send(mailMessage);

                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Error sending an email to, {0}.\r\n{1}.\r\n", firstEmailAddress, ex);
                            prWG.Update_Progress(textLine, configDirName);
                        }
                    }
                    
                }
                //End of EmailMessage

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error sending email for file.\r\n{0}.\r\n", ex);
                prWG.Update_Progress(textLine, configDirName);
            }
            finally
            {
               
            }

        }
        
        public void ZipReports(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            string textLine;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //
            // Zip the folder.
            //
            string dateTimeStamp = DateTime.Now.ToString("yyyyMMddhhmmssfff");
            string sourceFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef;
            string zipFileName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports\\" + rpParameters.ErRef + "_PDF_Reports_" + rpEmployer.HMRCDesc + "_" + dateTimeStamp + ".zip";
            string zipFileNameNoPassword = xdoc.Root.Element("DataHomeFolder").Value + "PE-ReportsNoPassword\\" + rpParameters.ErRef + "_PDF_Reports_" + rpEmployer.HMRCDesc + "_" + dateTimeStamp + ".zip";
            //
            //The password for the zipped reports file is the first 4 characters of the employer name in lower case ( or if the employers name is less than 4 characters then all the employers name )
            //plus the employers 4 digit number. Unless the employer has specified a particluar password in which case the password is held on the Companies table.
            //
            string password;
            if(rpEmployer.ReportPassword == null || rpEmployer.ReportPassword == "" || rpEmployer.ReportPassword == " ")
            {
                password = rpEmployer.Name.ToLower().Replace(" ", "");
                if (password.Length >= 4)
                {
                    password = password.Substring(0, 4);
                }
                password += rpParameters.ErRef;

            }
            else
            {
                password = rpEmployer.ReportPassword;
            }
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

                

            }
            catch (Exception ex)
            {
                textLine = string.Format("Error zipping pdf reports for source folder, {0}.\r\n{1}.\r\n", sourceFolder, ex);
                prWG.Update_Progress(textLine, configDirName);
            }

        }
        public void ProcessBankAndPensionFiles(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, List<RPPensionContribution> rpPensionContributions, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef;
            PicoXLSX.Workbook workbook;

            //Bank file code is not equal to "001" so a bank file is required.
            switch (rpEmployer.BankFileCode)
            {
                case "001":
                    //Barclays
                    CreateBarclaysBankFile(outgoingFolder, rpEmployeePeriodList, rpEmployer);
                    break;
                case "002":
                    //Eagle
                    workbook = CreateEagleBankFile(outgoingFolder, rpEmployeePeriodList, rpEmployer, rpParameters);
                    workbook.Save();
                    break;
                case "003":
                    //Revolut
                    CreateRevolutBankFile(outgoingFolder, rpEmployeePeriodList);
                    break;
                case "004":
                    //Natwest
                    break;
                case "005":
                    //Bottomline 
                    workbook = CreateBottomlineBankFile(xdoc, outgoingFolder, rpParameters);
                    workbook.Save();
                    break;
                default:
                    //No bank file required
                    break;
            }
            string previousSchemeName = null;
            //Create a list of pension file objects for each scheme name, then use it to write the pension
            //file for that scheme then move on to the next scheme name.
            //The rpPensionContributions object is already sorted in scheme name sequence
            RPPensionFileScheme rpPensionFileScheme = new RPPensionFileScheme();
            List<RPPensionContribution> rpPensionFileSchemePensionContributions = new List<RPPensionContribution>();
            List<RPPensionFileScheme> rpPensionFileSchemes = new List<RPPensionFileScheme>();
            foreach (RPPensionContribution rpPensionContribution in rpPensionContributions)
            {
                //If the AE Assessment Override is "Opt Out" don't include in the pension report.
                if(rpPensionContribution.AEAssessmentOverride != "Opt Out")
                {
                    if (rpPensionContribution.RPPensionPeriod.SchemeName != previousSchemeName)
                    {
                        //We've moved to a new scheme.
                        if (previousSchemeName != null)
                        {
                            //The rpPensionFileScheme object we've create should now contain a scheme name plus a list for employee contributions
                            rpPensionFileScheme.RPPensionContributions = rpPensionFileSchemePensionContributions;
                            rpPensionFileSchemes.Add(rpPensionFileScheme);
                            rpPensionFileScheme = new RPPensionFileScheme();
                            rpPensionFileSchemePensionContributions = new List<RPPensionContribution>();
                        }
                        previousSchemeName = rpPensionContribution.RPPensionPeriod.SchemeName;
                        rpPensionFileScheme.Key = rpPensionContribution.RPPensionPeriod.Key;
                        rpPensionFileScheme.SchemeName = rpPensionContribution.RPPensionPeriod.SchemeName;
                        rpPensionFileScheme.ProviderName = rpPensionContribution.RPPensionPeriod.ProviderName;
                        if (rpPensionFileScheme.ProviderName == null)
                        {
                            rpPensionFileScheme.ProviderName = "UNKOWN";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("AVIVA"))
                        {
                            rpPensionFileScheme.ProviderName = "AVIVA";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("NEST"))
                        {
                            rpPensionFileScheme.ProviderName = "NEST";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("WORKERS PENSION TRUST"))
                        {
                            rpPensionFileScheme.ProviderName = "WORKERS PENSION TRUST";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("CREATIVE AUTO ENROLMENT"))
                        {
                            rpPensionFileScheme.ProviderName = "CREATIVE AUTO ENROLMENT";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("THE PEOPLES PENSION"))
                        {
                            rpPensionFileScheme.ProviderName = "THE PEOPLES PENSION";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("SMART PENSION"))
                        {
                            rpPensionFileScheme.ProviderName = "SMART PENSION";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("ROYAL LONDON PENSION"))
                        {
                            rpPensionFileScheme.ProviderName = "ROYAL LONDON PENSION";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("NOW PENSION"))
                        {
                            rpPensionFileScheme.ProviderName = "NOW PENSION";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("LEGAL & GENERAL"))
                        {
                            rpPensionFileScheme.ProviderName = "LEGAL & GENERAL";
                        }
                        else if (rpPensionFileScheme.ProviderName.ToUpper().Contains("AEGON"))
                        {
                            rpPensionFileScheme.ProviderName = "AEGON";
                        }
                        else
                        {
                            rpPensionFileScheme.ProviderName = "UNKOWN";
                        }
                        
                    }
                    rpPensionFileSchemePensionContributions.Add(rpPensionContribution);
                }
                
            }
            //After the last rpPensionContribution create the final pensionFileScheme and add it to the list.
            //The rpPensionFileScheme object we've create should now contain a scheme name plus a list for employee contributions
            rpPensionFileScheme.RPPensionContributions = rpPensionFileSchemePensionContributions;
            rpPensionFileSchemes.Add(rpPensionFileScheme);
            ProcessPensionFileSchemes(xdoc, rpPensionFileSchemes, rpEmployer, rpParameters);
        }
        private void ProcessPensionFileSchemes(XDocument xdoc, List<RPPensionFileScheme> rpPensionFileSchemes, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef;

            foreach (RPPensionFileScheme rpPensionFileScheme in rpPensionFileSchemes)
            {
                switch (rpPensionFileScheme.ProviderName)
                {
                    case "AVIVA":
                        CreateTheAvivaPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "NEST":
                        CreateTheNestPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "WORKERS PENSION TRUST":
                        CreateTheWorkersPensionTrustPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "CREATIVE AUTO ENROLMENT":
                        CreateTheCreativeAEPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "THE PEOPLES PENSION":
                        CreateThePeoplesPensionPensionFile(outgoingFolder, rpPensionFileScheme, rpEmployer);
                        break;
                    case "SMART PENSION":
                    case "ROYAL LONDON PENSION":
                    case "NOW PENSION":
                    case "LEGAL & GENERAL":
                    case "AEGON":
                        //Get the transformed from PayRun.IO
                        GetCsvPensionsReport(xdoc, rpPensionFileScheme, rpParameters);
                        break;
                    case "UNKNOWN":
                        break;
                }
            }
        }
        public static string CreateBarclaysBankFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer)
        {
            string bankFileName = outgoingFolder + "\\" + "BarclaysBankFile.txt";
            string quotes = "\"";
            string comma = ",";

            //Create the Barclays bank file which does not have a header row.
            var stringBuilder = new StringBuilder();
            foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                if (rpEmployeePeriod.PaymentMethod == "BACS")
                {
                    string fullName = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                    fullName = fullName.ToUpper();
                    //The company name has to be restricted to a maximum of 18 characters. So set maxLength to 18 or the length of the company name which ever is shorter.
                    int maxLength = Math.Min(rpEmployer.Name.Length, 18);
                    var csvLine = quotes + rpEmployeePeriod.SortCode + quotes + comma +
                                  quotes + fullName + quotes + comma +
                                  quotes + rpEmployeePeriod.BankAccNo + quotes + comma +
                                  quotes + rpEmployeePeriod.NetPayTP.ToString() + quotes + comma +
                                  quotes + rpEmployer.Name.ToUpper().Substring(0,maxLength) + quotes + comma +
                                  quotes + "99" + quotes;

                    stringBuilder.AppendLine(csvLine);
                }
            }

            if (!string.IsNullOrEmpty(outgoingFolder))
            {
                using (StreamWriter sw = new StreamWriter(bankFileName))
                {
                    sw.Write(stringBuilder.ToString());
                }
            }
            
            return stringBuilder.ToString();
        }
        public static PicoXLSX.Workbook CreateEagleBankFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            //Create a PicoXLSX workbook
            RegexUtilities regexUtilities = new RegexUtilities();
            string workBookName = outgoingFolder + "\\" + regexUtilities.RemoveNonAlphaNumericChars(rpEmployer.Name) + "_EagleBankFile_Period_" + rpParameters.PeriodNo + ".xlsx";
            Workbook workbook = new Workbook(workBookName, "BACSDetails");

            //Write the header row
            workbook.CurrentWorksheet.AddNextCell("AccName");
            workbook.CurrentWorksheet.AddNextCell("SortCode");
            workbook.CurrentWorksheet.AddNextCell("AccNumber");
            workbook.CurrentWorksheet.AddNextCell("Amount");
            workbook.CurrentWorksheet.AddNextCell("Ref");

            //Write a row for each employee
            foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                if (rpEmployeePeriod.PaymentMethod == "BACS" && rpEmployeePeriod.NetPayTP != 0)
                {
                    workbook.CurrentWorksheet.GoToNextRow();
                    
                    string fullName = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                    fullName = fullName.ToUpper();
                    
                    workbook.CurrentWorksheet.AddNextCell(fullName);
                    workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.SortCode);
                    workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.BankAccNo);
                    workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetPayTP);
                    workbook.CurrentWorksheet.AddNextCell(fullName);
                }
            }

            return workbook;
        }
        public static PicoXLSX.Workbook CreateBottomlineBankFile(XDocument xdoc, string outgoingFolder, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlReport = GetBankFileReport(xdoc, rpParameters);
            //Create a PicoXLSX workbook
            string workBookName = outgoingFolder + "\\" + rpParameters.ErRef + "_BottomlineBankReport.xlsx";
            Workbook workbook = prWG.CreateBottomlineReportWorkbook(xmlReport, workBookName);

            return workbook;
        }
        private static XmlDocument GetBankFileReport(XDocument xdoc, RPParameters rPParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlReport = new XmlDocument();
            //PE-BankReport - Bank report
            string prm1 = "EmployerKey";
            string prm2 = "PayScheduleKey";
            string prm3 = "PaymentDate";
            string rptRef = "PE-BankFileReport";
            string val1 = rPParameters.ErRef;//Employer
            string val2 = rPParameters.PaySchedule;//Pay schedule
            string val3 = rPParameters.PayRunDate.ToString("yyyy-MM-dd");//Payment date
            xmlReport = prWG.RunReport(xdoc, rptRef, prm1, val1, prm2, val2, prm3, val3, null, null, null, null, null, null);
            return xmlReport;
        }
        public static string CreateRevolutBankFile(string outgoingFolder, List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            string bankFileName = outgoingFolder + "\\" + "RevolutBankFile.csv";
            string comma = ",";
            string month, year, fullName;
            StringBuilder stringBuilder = new StringBuilder();
            string csvLine = "Name,Recipient type,Account number,Sort code,Recipient bank country,Currency,Amount,Payment reference";
            stringBuilder.AppendLine(csvLine);
            //Create the Revolut bank file which does have a header row.
            foreach (RPEmployeePeriod rpEmployeePeriod in rpEmployeePeriodList)
            {
                if (rpEmployeePeriod.PaymentMethod == "BACS")
                {
                    fullName = rpEmployeePeriod.Forename + " " + rpEmployeePeriod.Surname;
                    month = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(rpEmployeePeriod.PayRunDate.Month);
                    year = rpEmployeePeriod.PayRunDate.Year.ToString();
                    csvLine = fullName + comma + "INDIVIDUAL" + comma + rpEmployeePeriod.BankAccNo + comma + rpEmployeePeriod.SortCode + comma + "GB" + comma + "GBP" + comma + rpEmployeePeriod.NetPayTP + comma + "Salary " + month + " " + year;
                    stringBuilder.AppendLine(csvLine);
                }
            }
            if (!string.IsNullOrEmpty(outgoingFolder))
            {
                using (StreamWriter sw = new StreamWriter(bankFileName))
                {
                    sw.Write(stringBuilder.ToString());
                }
            }
            return stringBuilder.ToString();
        }
        private void CreateTheNestPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            string providerEmployerReference = rpPensionFileScheme.RPPensionContributions[0].RPPensionPeriod.ProviderEmployerReference;
            string startDate = rpPensionFileScheme.RPPensionContributions[0].StartDate.ToString("yyyy-MM-dd");
            string endDate = rpPensionFileScheme.RPPensionContributions[0].EndDate.ToString("yyyy-MM-dd");
            string frequency = rpPensionFileScheme.RPPensionContributions[0].Freq;
            string blank = "";
            string zeroContributions;
            List<RPPensionContribution> joinersThisPeriod = new List<RPPensionContribution>();
            string header = 'H' + comma + providerEmployerReference + comma +
                                            "CS" + comma + endDate + comma + rpEmployer.NESTPensionText +
                                            comma + blank + comma + frequency + comma + blank +
                                            comma + blank + comma + startDate;

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                sw.WriteLine(header);
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.IsJoiner == true)
                    {
                        joinersThisPeriod.Add(rpPensionContribution); //Joiner needs to be included in both contributions file and joiner file
                    }

                    zeroContributions = ""; //need to reset the value else it will always be 5 
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod == 0 && rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod == 0)
                    {
                        zeroContributions = "5";
                    }
                    csvLine = 'D' + comma + rpPensionContribution.Surname + comma + rpPensionContribution.NINumber +
                    comma + rpPensionContribution.EeRef + comma + rpPensionContribution.RPPensionPeriod.PensionablePayTaxPeriod + comma +
                    blank + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma + rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod +
                    comma + zeroContributions;
                    sw.WriteLine(csvLine);
                }
                string footer = 'T' + comma + rpPensionFileScheme.RPPensionContributions.Count + comma + '3';
                sw.WriteLine(footer);
            }

            //if there are any joiners we create the joiner file
            if (joinersThisPeriod.Count > 0)
            {
                pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "JoinerFile.csv";
                string joinerCSVLine;
                string joinerDateOfBirth;
                string joinerStartDate;
                char niYesNo;
                char gender;
                header = 'H' + comma + providerEmployerReference + comma + "ME";

                using (StreamWriter joinerStream = new StreamWriter(pensionFileName))
                {
                    joinerStream.WriteLine(header);
                    foreach (RPPensionContribution joiner in joinersThisPeriod)
                    {
                        joinerDateOfBirth = joiner.DOB.ToString("yyyy-MM-dd");
                        joinerStartDate = joiner.RPPensionPeriod.StartJoinDate.Value.ToString("yyyy-MM-dd");
                        niYesNo = 'N'; //need to reset value
                        if (joiner.NINumber.Length == 0)
                        {
                            niYesNo = 'Y';
                        }
                        switch (joiner.Gender) //Gender needs to be a character
                        {
                            case ("Male"):
                                gender = 'M';
                                break;
                            case ("Female"):
                                gender = 'F';
                                break;
                            default:
                                gender = ' ';
                                break;
                        }
                        joinerCSVLine = 'D' + comma + joiner.Title + comma + joiner.Forename + comma + blank + comma +
                                                    joiner.Surname + comma + joinerDateOfBirth + comma + joiner.NINumber + comma +
                                                    niYesNo + comma + joiner.EeRef + comma + blank + comma + joiner.RPAddress.Line1 + comma +
                                                    joiner.RPAddress.Line2 + comma + joiner.RPAddress.Line3 + comma + joiner.RPAddress.Line4 + comma +
                                                    joiner.RPAddress.Postcode + comma + joiner.RPAddress.Country + comma + joiner.EmailAddress + comma + blank +
                                                    comma + gender + comma + 'Y' + comma + "AE" + comma + "My group" + comma + "My source" +
                                                    comma + joinerStartDate + comma + 'N';
                        joinerStream.WriteLine(joinerCSVLine);
                    }
                    string joinerFooter = 'T' + comma + joinersThisPeriod.Count + comma + "3";
                    joinerStream.WriteLine(joinerFooter);
                }
            }
        }
        private void CreateTheWorkersPensionTrustPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            List<RPPensionContribution> joinersThisPeriod = new List<RPPensionContribution>();

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.IsJoiner == true)
                    {
                        joinersThisPeriod.Add(rpPensionContribution); //Joiner needs to be included in both contributions file and joiner file
                    }
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
                    {
                        //csvLine = rpPensionContribution.NINumber + comma + rpPensionContribution.ForenameSurname + comma +
                        //                rpPensionContribution.PayRunDate.ToString("dd/MM/yyyy") + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod +
                        //                comma + rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod;
                        csvLine = rpPensionContribution.NINumber + comma + rpPensionContribution.ForenameSurname + comma +
                                  rpPensionContribution.StartDate.ToString("dd/MM/yyyy") + comma + rpPensionContribution.EndDate.ToString("dd/MM/yyyy") + comma +
                                  rpPensionContribution.RPPensionPeriod.TotalPayTaxPeriod + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma +
                                  rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod + comma + comma + comma + comma;
                        sw.WriteLine(csvLine);
                    }
                }
            }
            if (joinersThisPeriod.Count > 0)
            {
                pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "JoinerFile.csv";
                string joinerCSVLine;
                string joinerDateOfBirth;
                string joinerStartDate;
                string frequency;
                char gender;

                using (StreamWriter joinerStream = new StreamWriter(pensionFileName))
                {
                    foreach (RPPensionContribution joiner in joinersThisPeriod)
                    {
                        joinerDateOfBirth = joiner.DOB.ToString("dd/MM/yyyy");
                        joinerStartDate = joiner.RPPensionPeriod.StartJoinDate.Value.ToString("dd/MM/yyyy");

                        switch (joiner.Gender)
                        {
                            case ("Male"):
                                gender = 'M';
                                break;
                            case ("Female"):
                                gender = 'F';
                                break;
                            default:
                                gender = ' ';
                                break;
                        }
                        switch (joiner.Freq)
                        {
                            case ("Weekly"):
                                frequency = "W";
                                break;
                            case ("Monthly"):
                                frequency = "M";
                                break;
                            case ("Fortnightly"):
                                frequency = "F";
                                break;
                            case ("Four Weekly"):
                                frequency = "FW";
                                break;
                            case ("Quarterly"):
                                frequency = "Q";
                                break;
                            case ("Annual"):
                                frequency = "A";
                                break;
                            default:
                                frequency = "";
                                break;
                        }
                        joinerCSVLine = joiner.Forename + comma + joiner.Surname + comma + joinerDateOfBirth + comma + joiner.NINumber + comma + joiner.EmailAddress + comma +
                                                    joiner.EmailAddress + comma + gender + comma + "" + comma + joiner.RPPensionPeriod.ProviderEmployerReference + comma + joinerStartDate +
                                                    comma + joiner.RPPensionPeriod.PensionablePayTaxPeriod + comma + frequency + comma + "" + comma + "" + comma + joiner.RPAddress.Line1 + comma +
                                                    joiner.RPAddress.Line2 + comma + joiner.RPAddress.Line3 + comma + joiner.RPAddress.Line4 + comma + joiner.RPAddress.Postcode + comma +
                                                    joiner.RPAddress.Country;
                        joinerStream.WriteLine(joinerCSVLine);
                    }
                }
            }
        }
        private void CreateTheAvivaPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            string header = "PayrollMonth,Name,NInumber,AlternativeuniqueID,Employerregularcontributionamount,Employeeregulardeduction,Reasonforpartialornon-payment,Employerregularcontributionamount,Employeeoneoffcontribution,NewcategoryID";

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                sw.WriteLine(header);
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
                    {
                        csvLine = rpPensionContribution.PayRunDate.ToString("MMM-yy") + comma + rpPensionContribution.Surname + comma + rpPensionContribution.NINumber +
                                    comma + rpPensionContribution.EeRef + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma +
                                    rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod +
                                    comma + comma + comma + comma;

                        sw.WriteLine(csvLine);
                    }

                }
            }
        }

        private void CreateTheCreativeAEPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                string csvLine = null;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    string dateOfBirth = null;
                    if (rpPensionContribution.DOB.Year == 1)
                    {
                        dateOfBirth = null;
                    }
                    else
                    {
                        dateOfBirth = rpPensionContribution.DOB.ToString("dd/MM/yy");
                    }
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0) //if ee has either Ee or Er contributions
                    {
                        csvLine = rpPensionContribution.EeRef + comma + rpPensionContribution.Title + comma + rpPensionContribution.Forename + comma +
                                  rpPensionContribution.Surname + comma + rpPensionContribution.NINumber + comma + dateOfBirth + comma +
                                  rpPensionContribution.Gender + comma + rpPensionContribution.RPAddress.Line1 + comma + rpPensionContribution.RPAddress.Line2 + comma +
                                  rpPensionContribution.RPAddress.Line3 + comma + rpPensionContribution.RPAddress.Line4 + comma +
                                  rpPensionContribution.RPAddress.Postcode + comma + rpPensionContribution.StartDate.ToString("dd/MM/yy") + comma +
                                  rpPensionContribution.EndDate.ToString("dd/MM/yy") + comma + rpPensionContribution.RPPensionPeriod.PensionablePayTaxPeriod + comma +
                                  rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod + comma + "0" + comma + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma +
                                  "0";

                        sw.WriteLine(csvLine);
                    }

                }
            }
        }
        private void CreateThePeoplesPensionPensionFile(string outgoingFolder, RPPensionFileScheme rpPensionFileScheme, RPEmployer rpEmployer)
        {
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionFile.csv";
            string comma = ",";
            string providerEmployerReference = rpPensionFileScheme.RPPensionContributions[0].RPPensionPeriod.ProviderEmployerReference;
            string startDate = rpPensionFileScheme.RPPensionContributions[0].StartDate.ToString("dd/MM/yyyy");
            string endDate = rpPensionFileScheme.RPPensionContributions[0].EndDate.ToString("dd/MM/yyyy");

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                //2 headr line in this file
                string csvLine = 'H' + comma + providerEmployerReference + comma +
                                 startDate + comma + endDate + comma + rpEmployer.PensionReportFileType;
                sw.WriteLine(csvLine);
                csvLine = "Record Type,Title,Gender,Forename 1,Forename 2,Surname,Date of Birth," +
                          "National Insurance Number,Unique Identifier,Address 1,Address 2," +
                          "Address 3,Address 4,Address 5,Home Phone Number,Personal Email Address," +
                          "Date Employment Started,Start/Leaver Flag,Employment Ended,AE Worker Group," +
                          "AE Status,AE Date,Scheme Join Date,Opt Out Date,Opt In Date,Total Earnings Per PRP," +
                          "Pensionable Earnings Per PRP,Employer Pension Contribution,Employee Pension Contribution," +
                          "Missing/Partial Pension Code,EAC/ELC Premium,Date AE Information Received";
                sw.WriteLine(csvLine);
                decimal totalContributions = 0;

                foreach (RPPensionContribution rpPensionContribution in rpPensionFileScheme.RPPensionContributions)
                {
                    if (rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod != 0 || rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod != 0)
                    {
                        string leavingDate = null;
                        if (rpPensionContribution.LeavingDate != null)
                        {
                            leavingDate = rpPensionContribution.LeavingDate.Value.ToString("dd/MM/yyyy");
                        }
                        csvLine = 'D' + comma +
                                  rpPensionContribution.Title + comma +
                                  rpPensionContribution.Gender + comma +
                                  rpPensionContribution.Forename + comma +
                                  "" + comma +  //2nd Forename
                                  rpPensionContribution.Surname + comma +
                                  rpPensionContribution.DOB.ToString("dd/MM/yyyy") + comma +
                                  rpPensionContribution.NINumber + comma +
                                  rpPensionContribution.EeRef + comma +
                                  rpPensionContribution.RPAddress.Line1 + comma +
                                  rpPensionContribution.RPAddress.Line2 + comma +
                                  rpPensionContribution.RPAddress.Line3 + comma +
                                  rpPensionContribution.RPAddress.Line4 + comma +
                                  rpPensionContribution.RPAddress.Postcode + comma +
                                  "" + comma + //Home phone number
                                  rpPensionContribution.EmailAddress + comma +
                                  rpPensionContribution.StartingDate.ToString("dd/MM/yyyy") + comma +
                                  "" + comma + //Starter/Leaver Flag
                                  leavingDate + comma +
                                  rpPensionContribution.RPPensionPeriod.AEWorkerGroup + comma +
                                  rpPensionContribution.RPPensionPeriod.AEStatus + comma +
                                  rpPensionContribution.RPPensionPeriod.AEAssessmentDate.Value.ToString("dd/MM/yyyy") + comma +
                                  rpPensionContribution.RPPensionPeriod.StartJoinDate.Value.ToString("dd/MM/yyyy") + comma +
                                  "" + comma + //Opt Out Date
                                  "" + comma + //Opt In Date
                                  rpPensionContribution.RPPensionPeriod.TotalPayTaxPeriod + comma +
                                  rpPensionContribution.RPPensionPeriod.PensionablePayTaxPeriod + comma +
                                  rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + comma +
                                  rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod + comma +
                                  "" + comma +  //Missing/Partial Pension Code
                                  "0" + comma + //EAC/ELC Premium
                                  "";           //Date AE Information Received
                        sw.WriteLine(csvLine);
                        totalContributions = totalContributions + rpPensionContribution.RPPensionPeriod.ErPensionTaxPeriod + rpPensionContribution.RPPensionPeriod.EePensionTaxPeriod;
                    }

                }
                csvLine = 'T' + comma + totalContributions;
                sw.WriteLine(csvLine);
            }

        }
        private void GetCsvPensionsReport(XDocument xdoc, RPPensionFileScheme rpPensionFileScheme, RPParameters rpParameters)
        {
            bool joinerRequired = false;
            if(rpPensionFileScheme.ProviderName == "LEGAL & GENERAL" || rpPensionFileScheme.ProviderName == "AEGON")
            {
                joinerRequired = true;
            }
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef;
            string pensionFileName = outgoingFolder + "\\" + rpPensionFileScheme.SchemeName + "PensionContributionsFile.csv";

            //Get the transformed report from PayRun.IO. It'll return the csv file as required and I won't need to run the CreateTheSmartPensionsPensionFile method.
            bool isJoiner = false;
            string csvReport = prWG.GetCsvPensionsReport(xdoc, rpParameters, rpPensionFileScheme, isJoiner);

            using (StreamWriter sw = new StreamWriter(pensionFileName))
            {
                sw.Write(csvReport);
            }
            if(joinerRequired)
            {
                isJoiner = true;
                csvReport = prWG.GetCsvPensionsReport(xdoc, rpParameters, rpPensionFileScheme, isJoiner);
                pensionFileName = pensionFileName.Replace("PensionContributionsFile.csv", "PensionJoinersFile.csv");
                using (StreamWriter sw = new StreamWriter(pensionFileName))
                {
                    sw.Write(csvReport);
                }
            }
        }
        public void PrintStandardReports(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters,
                                         List<P45> p45s, List<RPPayComponent> rpPayComponents, List<RPPensionContribution> rpPensionContributions)
        {
            PrintPayslips(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            if(rpEmployer.ReportsInExcelFormat)
            {
                PrintPayslipsSimple(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            }
            PrintPaymentsDueByMethodReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            PrintPensionContributionsReport(xdoc, rpPensionContributions, rpEmployer, rpParameters);
            bool showDetail = true;
            PrintComponentAnalysisReport(xdoc, rpPayComponents, rpEmployer, rpParameters,showDetail);
            showDetail = false;
            PrintComponentAnalysisReport(xdoc, rpPayComponents, rpEmployer, rpParameters, showDetail);
            PrintPayrollRunDetailsReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            if (rpEmployer.PayRunDetailsYTDRequired)
            {
                PrintPayrollRunDetailsYTDReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            }
            if (rpEmployer.PayrollTotalsSummaryRequired)
            {
                PrintPayrollTotalsSummaryReport(xdoc, rpEmployeePeriodList, rpEmployer, rpParameters);
            }
            if (rpParameters.PaidInCash)
            {
                //At least one employee was paid in cash to we may need to run the Note And Coin Requirement report
                if(rpEmployer.NoteAndCoinRequired) 
                {
                    PrintNoteAndCoinRequiredReport(xdoc, rpEmployer, rpParameters);
                }
            }
            if(rpParameters.AOERequired)
            {
                //At least one employee has an AOE amount.
                PrintCurrentAttachmentOfEarningsOrders(xdoc, rpEmployer, rpParameters);
            }

            if (p45s.Count > 0)
            {
                PrintP45s(xdoc, p45s, rpParameters, rpEmployer);
            }
        }
        private void PrintNoteAndCoinRequiredReport(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlReport = prWG.GetNoteAndCoinRequirementReport(xdoc, rpParameters);

            string reportName = "NoteAndCoinRequirementReport";
            string assemblyName = "PayRunIOClassLibrary";
            XtraReport xtraReport = prWG.CreatePDFReport(xmlReport, reportName, assemblyName);

            string dirName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef + "\\";
            Directory.CreateDirectory(dirName);
            string docName = rpParameters.ErRef + "_NoteAndCoinRequirementReportFor_TaxYear_" + rpParameters.TaxYear + "_Period_" + rpParameters.PeriodNo + ".pdf";

            xtraReport.ExportToPdf(dirName + docName);
            if (rpEmployer.ReportsInExcelFormat)
            {
                docName = docName.Replace(".pdf", ".xlsx");
                xtraReport.ExportToXlsx(dirName + docName);
            }
        }
        private void PrintCurrentAttachmentOfEarningsOrders(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {

            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlReport = prWG.GetCurrentAttachmentOfEarningsOrders(xdoc, rpParameters);

            string reportName = "CurrentAttachmentOfEarningsOrders";
            string assemblyName = "PayRunIOClassLibrary";
            XtraReport xtraReport = prWG.CreatePDFReport(xmlReport, reportName, assemblyName);

            string dirName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef + "\\";
            Directory.CreateDirectory(dirName);
            string docName = rpParameters.ErRef + "_CurrentAttachmentOfEarningsOrdersFor_TaxYear_" + rpParameters.TaxYear + "_Period_" + rpParameters.PeriodNo + ".pdf";

            xtraReport.ExportToPdf(dirName + docName);
            if (rpEmployer.ReportsInExcelFormat)
            {
                docName = docName.Replace(".pdf", ".xlsx");
                xtraReport.ExportToXlsx(dirName + docName);
            }
        }
        private void PrintPayslips(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Main payslip report
            string reportName = "Payslip";
            string assemblyName = "PayRunIOClassLibrary";
            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);

            //Don't save the full payslip as an xlsx file as it doesn't render well, which is why I've created a simple payslip.
            bool reportInExcelFormat = rpEmployer.ReportsInExcelFormat;
            rpEmployer.ReportsInExcelFormat = false;
            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
            rpEmployer.ReportsInExcelFormat = reportInExcelFormat;
        }
        
        private void PrintPayslipsSimple(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Simple payslip report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PayslipSimple";
            
            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);
            
            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintPaymentsDueByMethodReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Payments Due By Method report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PaymentsDueByMethodsReport";

            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintPensionContributionsReport(XDocument xdoc, List<RPPensionContribution> rpPensionContributions, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Pensions Copntributions report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PensionContributionsReport";
            XtraReport xtraReport = prWG.CreatePDFReport(rpPensionContributions, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintComponentAnalysisReport(XDocument xdoc, List<RPPayComponent> rpPayComponents, RPEmployer rpEmployer, RPParameters rpParameters, bool showDetail)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Component Analysis report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "ComponentAnalysisReport";
            XtraReport xtraReport = prWG.CreatePDFReport(rpPayComponents, rpEmployer, rpParameters, showDetail, reportName, assemblyName);
            if (showDetail)
            {
                reportName = "ComponentAnalysisReport";
            }
            else
            {
                reportName = "ComponentTotalsReport";
            }
            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintPayrollRunDetailsReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            //Payroll Run Details report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PayrollRunDetailsReport";

            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintPayrollRunDetailsYTDReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            //Payroll Run Details YTD report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PayrollRunDetailsYTDReport";

            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void PrintPayrollTotalsSummaryReport(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            //Payroll Totals Summary report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "PayrollTotalsSummaryReport";

            XtraReport xtraReport = prWG.CreatePDFReport(rpEmployeePeriodList, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }

        private void PrintP45s(XDocument xdoc, List<P45> p45s, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            //P45 report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "P45";

            XtraReport xtraReport = prWG.CreatePDFReport(p45s, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        public void PrintApprenticeshipLevyReport(XDocument xdoc, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlReport = prWG.GetApprenticeshipLevyReport(xdoc, rpParameters);

            string reportName = "ApprenticeshipLevyReport";
            string assemblyName = "PayRunIOClassLibrary";
            XtraReport xtraReport = prWG.CreatePDFReport(xmlReport, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        public void PrintP32Report(XDocument xdoc, RPP32Report rpP32Report, RPParameters rpParameters, RPEmployer rpEmployer)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            //Main P32 report
            string assemblyName = "PayRunIOClassLibrary";
            string reportName = "P32Report";

            XtraReport xtraReport = prWG.CreatePDFReport(rpP32Report, rpEmployer, rpParameters, reportName, assemblyName);

            SaveReport(xtraReport, xdoc, rpEmployer, rpParameters, reportName);
        }
        private void SaveReport(XtraReport xtraReport, XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters, string reportName)
        {
            string dirName = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports" + "\\" + rpParameters.ErRef + "\\";
            string docName = rpParameters.ErRef + "_" + reportName + "ReportFor_TaxYear_" + rpParameters.TaxYear + "_Period_" + rpParameters.PeriodNo + ".pdf";

            bool designMode = false;
            if (designMode)
            {
                xtraReport.ShowDesigner();
            }
            else
            {
                // Export to pdf file.
                Directory.CreateDirectory(dirName);

                //Don't save the PDF for the payslip simple.
                if(reportName != "PayslipSimple")
                {
                    xtraReport.ExportToPdf(dirName + docName);
                }
                //Check if it's required in xlsx format.
                if (rpEmployer.ReportsInExcelFormat)
                {
                    docName = docName.Replace(".pdf", ".xlsx");
                    xtraReport.ExportToXlsx(dirName + docName);
                }

            }
        }
        private static Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                           List<RPPensionContribution>, RPEmployer, RPParameters> 
                           PrepareStandardReports(XDocument xdoc, XmlDocument xmlReport, RPParameters rpParameters)
        {
            bool aoeRequired = false;
            bool paidInCash = false;
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc?.Root?.Element("LogOneIn")?.Value);
            string configDirName = xdoc?.Root?.Element("SoftwareHomeFolder")?.Value;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            
            List<RPEmployeePeriod> rpEmployeePeriodList = new List<RPEmployeePeriod>();
            List<P45> p45s = new List<P45>();
            //Create a list of Pay Code totals for the Payroll Component Analysis report
            List<RPPayComponent> rpPayComponents = new List<RPPayComponent>();
            RPEmployer rpEmployer = prWG.GetRPEmployer(xdoc, xmlReport, rpParameters);
            //Create a list of all possible Pay Codes just from the first employee
            bool preSamplePayCodes = false;
            List<RPPreSamplePayCode> rpPreSamplePayCodes = new List<RPPreSamplePayCode>();
            List<RPPensionContribution> rpPensionContributions = new List<RPPensionContribution>();

            try
            {
                //bool gotPayRunDate = false;
                foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
                {
                    bool include = false;

                    string payRunDate = prWG.GetElementByTagFromXml(employee, "PayRunDate");

                    if (payRunDate != "No Pay Run Data Found" && payRunDate != null)
                    {
                        //if (!gotPayRunDate)
                        //{
                        //    rpParameters.PayRunDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "PayRunDate"));
                        //    gotPayRunDate = true;
                            
                        //}
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
                        rpEmployeePeriod.StartingDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "StartDate"));
                        rpEmployeePeriod.Gender = prWG.GetElementByTagFromXml(employee, "Gender");
                        rpEmployeePeriod.BuildingSocRef = prWG.GetElementByTagFromXml(employee, "BuildingSocRef");
                        rpEmployeePeriod.NINumber = prWG.GetElementByTagFromXml(employee, "NiNumber");
                        rpEmployeePeriod.PaymentMethod = prWG.GetElementByTagFromXml(employee, "PayMethod");
                        if(rpEmployeePeriod.PaymentMethod=="Cash")
                        {
                            //Need to produce the NoteAndCoinRequirement report
                            paidInCash = true;
                           
                        }
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
                        rpEmployeePeriod.EeContributionsPt2 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionsPt2");
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
                            rpEmployeePeriod.TaxCode += " W1";
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
                            rpAEAssessment.WorkersGroup = prWG.GetElementByTagFromXml(aeAssessment, "WorkersGroup");
                            if(rpAEAssessment.WorkersGroup == null)
                            {
                                rpAEAssessment.WorkersGroup = rpEmployer.PensionReportAEWorkersGroup;
                            }
                            rpAEAssessment.Status = GetAEAssessmentStatus(rpAEAssessment.AssessmentCode);
                            

                        }
                        //Split these strings on capital letters by inserting a space before each capital letter.
                        rpAEAssessment.AssessmentCode = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentCode);
                        rpAEAssessment.AssessmentEvent = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentEvent);
                        rpAEAssessment.AssessmentResult = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentResult);
                        rpAEAssessment.AssessmentOverride = SplitStringOnCapitalLetters(rpAEAssessment.AssessmentOverride);

                        rpEmployeePeriod.AEAssessment = rpAEAssessment;

                        rpEmployeePeriod.EePensionTotalTP = 0;
                        rpEmployeePeriod.EePensionTotalYtd = 0;
                        rpEmployeePeriod.ErPensionTotalTP = 0;
                        rpEmployeePeriod.ErPensionTotalYtd = 0;
                        rpEmployeePeriod.Frequency = rpParameters.PaySchedule;

                        List<RPPensionPeriod> rpPensionPeriods = new List<RPPensionPeriod>();
                        foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                        {
                            RPPensionPeriod rpPensionPeriod = new RPPensionPeriod();
                            rpPensionPeriod.Key = Convert.ToInt32(pension.GetAttribute("Key"));
                            rpPensionPeriod.Code = prWG.GetElementByTagFromXml(pension, "Code");
                            rpPensionPeriod.ProviderName = prWG.GetElementByTagFromXml(pension, "ProviderName");
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
                            rpPensionPeriod.AEAssessmentDate = rpEmployeePeriod.AEAssessment.AssessmentDate;
                            rpPensionPeriod.AEWorkerGroup = rpEmployeePeriod.AEAssessment.WorkersGroup;
                            rpPensionPeriod.AEStatus = rpEmployeePeriod.AEAssessment.Status;
                            rpPensionPeriod.TotalPayTaxPeriod = rpEmployeePeriod.Gross;
                            rpPensionPeriod.StatePensionAge = rpEmployeePeriod.AEAssessment.StatePensionAge;
                            
                            

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
                            rpPensionContribution.StartingDate = rpEmployeePeriod.StartingDate;
                            rpPensionContribution.LeavingDate = rpEmployeePeriod.LeavingDate;
                            
                            //The address gets re-arranged later so that there are no blank lines shown. There address as provided by PR is in this address array.
                            RPAddress rpAddress = new RPAddress();
                            rpAddress.Line1 = address[0];
                            rpAddress.Line2 = address[1];
                            rpAddress.Line3 = address[2];
                            rpAddress.Line4 = address[3];
                            rpAddress.Postcode = address[4];
                            rpAddress.Country = address[5];

                            rpPensionContribution.RPAddress = rpAddress;
                            rpPensionContribution.EmailAddress = "";
                            rpPensionContribution.Gender = rpEmployeePeriod.Gender;
                            rpPensionContribution.NINumber = rpEmployeePeriod.NINumber;
                            rpPensionContribution.Freq = rpEmployeePeriod.Frequency;
                            rpPensionContribution.StartDate = rpEmployeePeriod.PeriodStartDate;
                            rpPensionContribution.EndDate = rpEmployeePeriod.PeriodEndDate;
                            rpPensionContribution.PayRunDate = rpEmployeePeriod.PayRunDate;
                            rpPensionContribution.SchemeFileType = "SchemeFileType";
                            rpPensionContribution.AEAssessmentOverride = rpAEAssessment.AssessmentOverride;
                            rpPensionContribution.RPPensionPeriod = rpPensionPeriod;

                            rpPensionContributions.Add(rpPensionContribution);

                            rpEmployeePeriod.EePensionTotalTP += rpPensionPeriod.EePensionTaxPeriod;
                            rpEmployeePeriod.EePensionTotalYtd += rpPensionPeriod.EePensionYtd;
                            rpEmployeePeriod.ErPensionTotalTP += rpPensionPeriod.ErPensionTaxPeriod;
                            rpEmployeePeriod.ErPensionTotalYtd += rpPensionPeriod.ErPensionYtd;
                        }
                        rpEmployeePeriod.Pensions = rpPensionPeriods;

                        rpEmployeePeriod.DirectorshipAppointmentDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "DirectorshipAppointmentDate"));
                        rpEmployeePeriod.Director = prWG.GetBooleanElementByTagFromXml(employee, "Director");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt1 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt1");
                        rpEmployeePeriod.EeContributionsTaxPeriodPt2 = prWG.GetDecimalElementByTagFromXml(employee, "EeContributionTaxPeriodPt2");
                        rpEmployeePeriod.NetPayYTD = 0;
                        rpEmployeePeriod.TotalPayTP = 0;
                        rpEmployeePeriod.TotalPayYTD = 0;
                        rpEmployeePeriod.TotalDedTP = 0;
                        rpEmployeePeriod.TotalDedYTD = 0;
                        rpEmployeePeriod.TotalOtherDedTP = 0;
                        rpEmployeePeriod.TotalOtherDedYTD = 0;
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
                        rpEmployeePeriod.AbsencePayYtd = 0;
                        rpEmployeePeriod.HolidayPay = 0;
                        rpEmployeePeriod.PreTaxPension = 0;
                        rpEmployeePeriod.Tax = 0;
                        rpEmployeePeriod.NetNI = 0;
                        rpEmployeePeriod.PostTaxAddDed = 0;
                        rpEmployeePeriod.PostTaxPension = 0;
                        rpEmployeePeriod.AEO = 0;
                        rpEmployeePeriod.AEOYtd = 0;
                        rpEmployeePeriod.StudentLoan = 0;
                        rpEmployeePeriod.StudentLoanYtd = 0;
                        rpEmployeePeriod.TotalPayComponentAdditions = 0;
                        rpEmployeePeriod.TotalPayComponentDeductions = 0;
                        rpEmployeePeriod.BenefitsInKind = 0;
                        rpEmployeePeriod.SSPSetOff = 0;
                        rpEmployeePeriod.SSPAdd = 0;
                        rpEmployeePeriod.SMPSetOff = 0;
                        rpEmployeePeriod.SMPAdd = 0;
                        rpEmployeePeriod.OSPPSetOff = 0;
                        rpEmployeePeriod.OSPPAdd = 0;
                        rpEmployeePeriod.SAPSetOff = 0;
                        rpEmployeePeriod.SAPAdd = 0;
                        rpEmployeePeriod.ShPPSetOff = 0;
                        rpEmployeePeriod.ShPPAdd = 0;
                        rpEmployeePeriod.SPBPSetOff = 0;
                        rpEmployeePeriod.SPBPAdd = 0;
                        rpEmployeePeriod.Zero = 0;

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
                                //If the pay component is one of the statutory absences set IsPayCode to false regardless of what's in the file.
                                rpPayComponent.IsPayCode = SetStatutoryAbsenceIsPayCode(rpPayComponent.IsPayCode, rpPayComponent.PayCode);
                                rpPayComponent.EarningOrDeduction = prWG.GetElementByTagFromXml(payCode, "EarningOrDeduction");
                                //Check if it's an AEO with an amount other than zero.
                                if(rpPayComponent.PayCode == "AOE" && rpPayComponent.AmountTP != 0)
                                {
                                    aoeRequired = true;
                                }
                                if (rpPayComponent.AmountTP != 0 || rpPayComponent.AmountYTD != 0)
                                {
                                    //Value is not equal to zero so go through the list of Pre Sample codes and mark this one as in use
                                    rpPreSamplePayCodes = MarkPreSampleCodeAsInUse(rpPayComponent.PayCode, rpPreSamplePayCodes);
                                    if (rpPayComponent.IsPayCode)
                                    {
                                        if(rpPayComponent.EarningOrDeduction == "E")
                                        {
                                            rpEmployeePeriod.TotalPayComponentAdditions += rpPayComponent.AmountTP;
                                        }
                                        else
                                        {
                                            rpEmployeePeriod.TotalPayComponentDeductions += rpPayComponent.AmountTP;
                                        }
                                        rpPayComponents.Add(rpPayComponent);
                                    }
                                    //Probably should bite the bullet and make use of the IsPayCode marker here rather than looking for TAX, NI, PENSION, SLOAN, AOE etc.
                                    //but I'm concerned it will cause unforseen issues.
                                    if(rpPayComponent.PayCode != "TAX" && rpPayComponent.PayCode != "NI" && !rpPayComponent.PayCode.StartsWith("PENSION") && 
                                       rpPayComponent.PayCode != "SLOAN" && rpPayComponent.PayCode != "PGL" &&
                                       rpPayComponent.PayCode != "AOE")
                                    {
                                        if (rpPayComponent.IsTaxable)
                                        {
                                            rpEmployeePeriod.PreTaxAddDed += rpPayComponent.AmountTP;
                                        }
                                        else
                                        {
                                            rpEmployeePeriod.PostTaxAddDed += rpPayComponent.AmountTP;
                                        }
                                    }

                                    //Check for the different pay codes and add to the appropriate total.
                                    switch (rpPayComponent.PayCode)
                                    {
                                        case "HOLPY":
                                        case "HOLIDAY":
                                            rpEmployeePeriod.HolidayPay += rpPayComponent.AmountTP;
                                            break;
                                        case "PENSIONRAS":
                                        case "PENSIONTAXEX":
                                            rpEmployeePeriod.PostTaxPension += rpPayComponent.AmountTP;
                                            break;
                                        case "PENSION":
                                        case "PENSIONSS":
                                            rpEmployeePeriod.PreTaxPension += rpPayComponent.AmountTP;
                                            break;
                                        case "AOE":
                                            rpEmployeePeriod.AEO += (rpPayComponent.AmountTP * -1);
                                            rpEmployeePeriod.AEOYtd += (rpPayComponent.AmountYTD * -1);
                                            break;
                                        case "SLOAN":
                                        case "PGL":
                                            rpEmployeePeriod.StudentLoan += (rpPayComponent.AmountTP * -1);
                                            rpEmployeePeriod.StudentLoanYtd += (rpPayComponent.AmountYTD * -1);
                                            break;
                                        case "TAX":
                                            rpEmployeePeriod.Tax += rpPayComponent.AmountTP;
                                            break;
                                        case "NI":
                                            rpEmployeePeriod.NetNI += rpPayComponent.AmountTP;
                                            break;
                                        case "SAP":
                                            rpEmployeePeriod.SAPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SAPOFFSET":
                                            rpEmployeePeriod.SAPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SHPP":
                                            rpEmployeePeriod.ShPPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SHPPOFFSET":
                                            rpEmployeePeriod.ShPPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SMP":
                                            rpEmployeePeriod.SMPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SMPOFFSET":
                                            rpEmployeePeriod.SMPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SPBP":
                                            rpEmployeePeriod.SPBPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SPBPOFFSET":
                                            rpEmployeePeriod.SPBPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SPP":
                                            rpEmployeePeriod.OSPPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SPPOFFSET":
                                            rpEmployeePeriod.OSPPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SSP":
                                            rpEmployeePeriod.SSPAdd += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
                                            break;
                                        case "SSPOFFSET":
                                            rpEmployeePeriod.SSPSetOff += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePay += rpPayComponent.AmountTP;
                                            rpEmployeePeriod.AbsencePayYtd += rpPayComponent.AmountYTD;
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
                                    rpAddition.IsPayCode = rpPayComponent.IsPayCode;
                                    if (rpAddition.AmountTP != 0)
                                    {
                                        rpAdditions.Add(rpAddition);
                                        
                                    }
                                    rpEmployeePeriod.TotalPayTP += rpAddition.AmountTP;
                                    rpEmployeePeriod.TotalPayYTD += rpAddition.AmountYTD;
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
                                    rpDeduction.Description = rpPayComponent.Description;
                                    rpDeduction.IsTaxable = rpPayComponent.IsTaxable;
                                    rpDeduction.AmountTP = rpPayComponent.AmountTP * -1;
                                    rpDeduction.AmountYTD = rpPayComponent.AmountYTD * -1;
                                    rpDeduction.AccountsYearBalance = rpPayComponent.AccountsYearBalance * -1;
                                    rpDeduction.AccountsYearUnits = rpPayComponent.AccountsYearUnits * -1;
                                    rpDeduction.PayeYearUnits = rpPayComponent.UnitsYTD * -1;
                                    rpDeduction.PayrollAccrued = rpPayComponent.PayrollAccrued * -1;
                                    rpDeduction.IsPayCode = rpPayComponent.IsPayCode;
                                    if (rpDeduction.AmountTP != 0 || rpDeduction.Code.Contains("PENSION"))  //Adding pension in even if they are zero because several can be added together
                                    {
                                        rpDeductions.Add(rpDeduction);

                                    }
                                    rpEmployeePeriod.TotalDedTP += rpDeduction.AmountTP;
                                    rpEmployeePeriod.TotalDedYTD += rpDeduction.AmountYTD;
                                    //Other Deduction is a column on the Pay Run Details YTD report
                                    if(!rpDeduction.Code.Contains("PENSION") && rpDeduction.Code != "TAX" && rpDeduction.Code != "NI" 
                                        && rpDeduction.Code != "AOE" && rpDeduction.Code != "SLOAN"
                                        && rpDeduction.Code != "PGL")
                                    {
                                        rpEmployeePeriod.TotalOtherDedTP += rpDeduction.AmountTP;
                                        rpEmployeePeriod.TotalOtherDedYTD += rpDeduction.AmountYTD;
                                    }
                                    

                                    //We now need a list of deductions for the PayHistory.csv file and a different one for the payslips.
                                    //Deductions used to create the PayHistory.csv file will use the PayCodes provided in the PR xml file for pensions, for the payslip use the pension list from PR.
                                    if (!rpDeduction.Code.Contains("PENSION"))
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
                        rpEmployeePeriod.Tax *= -1;
                        rpEmployeePeriod.NetNI *= -1;
                        //Multiple the Pre-Tax Pension & Post-Tax pension by -1 to make them show as positive on the Payroll Run Details report.
                        rpEmployeePeriod.PreTaxPension *= -1;
                        rpEmployeePeriod.PostTaxPension *= -1;

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
                                p45.MonthNo = rpParameters.PeriodNo;
                                p45.WeekNo = 0;
                            }
                            else
                            {
                                p45.MonthNo = 0;
                                p45.WeekNo = rpParameters.PeriodNo;
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
                            p45.ErAddress1 = rpEmployer.Address1;
                            p45.ErAddress2 = rpEmployer.Address2;
                            p45.ErAddress3 = rpEmployer.Address3;
                            p45.ErAddress4 = rpEmployer.Address4;
                            p45.ErPostcode = rpEmployer.Postcode;
                            p45.ErCountry = rpEmployer.Country;
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
                prWG.Update_Progress(textLine, configDirName);
            }
            if(paidInCash)
            {
                //At least one employee was paid in cash so we may need the NotesAndCoinRequirement report
                rpParameters.PaidInCash = true;
            }
            if(aoeRequired)
            {
                //At least one employee has an amount of AOE
                rpParameters.AOERequired = true;
            }
            return new Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                                  List<RPPensionContribution>, RPEmployer, RPParameters>
                                  (rpEmployeePeriodList, rpPayComponents, p45s, rpPreSamplePayCodes, rpPensionContributions, rpEmployer, rpParameters);

        }

        private static bool SetStatutoryAbsenceIsPayCode(bool isPayCode, string payCode)
        {
            switch (payCode)
            {
                case "SAP":
                case "SAPOFFSET":
                case "SHPP":
                case "SHPPOFFSET":
                case "SMP":
                case "SMPOFFSET":
                case "SPBP":
                case "SPBPOFFSET":
                case "SPP":
                case "SPPOFFSET":
                case "SSP":
                case "SSPOFFSET":
                    isPayCode = false;
                    break;
                default:
                    break;
            }
            return isPayCode;
        }
        private static string GetAEAssessmentStatus(string assessmentCode)
        {
            if (assessmentCode != null)
            {
                if (assessmentCode.ToUpper().Contains("NONELIGIBLE"))
                {
                    assessmentCode = "NonEligible";
                }
                else if (assessmentCode.ToUpper().Contains("ELIGIBLE"))
                {
                    assessmentCode = "Eligible";
                }
                else if (assessmentCode.ToUpper().Contains("EXCLUDED"))
                {
                    assessmentCode = "Excluded";
                }
                else if (assessmentCode.ToUpper().Contains("ENTITLED"))
                {
                    assessmentCode = "Entitled";
                }
            }
            return assessmentCode;
        }
        private static string SplitStringOnCapitalLetters(string input)
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
        
        private static List<RPPreSamplePayCode> MarkPreSampleCodeAsInUse(string payCode, List<RPPreSamplePayCode> rpPreSamplePayCodes)
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
        public bool ProcessOutputFiles(XDocument xdoc, string directory)
        {
            //Old method going through directories created by PR
            string textLine = null;
            int logOneIn = Convert.ToInt32(xdoc.Root.Element("LogOneIn").Value);
            string configDirName = xdoc.Root.Element("SoftwareHomeFolder").Value;
            bool eePeriodProcessed = false;
            bool eeYtdProcessed = false;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            DirectoryInfo dirInfo = new DirectoryInfo(directory);
            //We now have the possibility that there are multiple EmployeePeriod files and multiple EmployeeYtd files.
            //For large numbers of employees they are being split into batches of 100 employees.
            //I'm going run through an combine them into one file for EmployeePeriod and one for EmployeeYtd.
            //Combine EmployeePeriod files
            CombineXmlFiles(directory,"Period");
            //Combine EmployeeYtd files
            CombineXmlFiles(directory, "Ytd");
            FileInfo[] files = dirInfo.GetFiles("*.xml");
            //We haven't got the correct payroll run date in the "EmployeeYtd" report so I'm going use the RPParameters from the "EmployeePeriod" report.
            //I'm just a bit concerned of the order they will come in. Hopefully always alphabetical.
            RPParameters rpParameters = null;
            RPEmployer rpEmployer = null;
            foreach (FileInfo file in files.OrderBy(x => x.Name))
            {
                if (file.FullName.Contains("CompleteEmployeePeriod"))
                {
                    List<RPEmployeePeriod> rpEmployeePeriodList = null;
                    List<RPPayComponent> rpPayComponents = null;
                    List<P45> p45s = null;
                    List<RPPreSamplePayCode> rpPreSamplePayCodes = null;
                    List<RPPensionContribution> rpPensionContributions = null;
                    
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
                        prWG.Update_Progress(textLine, configDirName);
                    }
                    if(!rpEmployer.HoldPayHistory)
                    {
                        try
                        {
                            CreateHistoryCSV(xdoc, rpParameters, rpEmployer, rpEmployeePeriodList);
                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Error creating the history csv file for file {0}.\r\n{1}.\r\n", file, ex);
                            prWG.Update_Progress(textLine, configDirName);
                        }
                    }
                    

                    try
                    {
                        ProducePeriodReports(xdoc, rpEmployeePeriodList, rpEmployer, p45s, rpPayComponents, rpParameters, rpPreSamplePayCodes, rpPensionContributions);

                        eePeriodProcessed = true;
                    }   
                    catch (Exception ex)
                    {
                        textLine = string.Format("Error producing the employee period reports for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.Update_Progress(textLine, configDirName);
                    }
                   
                }
                else if (file.FullName.Contains("CompleteEmployeeYtd"))
                {
                    eeYtdProcessed = true;
                    if (!rpEmployer.HoldPayHistory)
                    {
                        try
                        {
                            var tuple = PrepareYTDReport(xdoc, file);
                            List<RPEmployeeYtd> rpEmployeeYtdList = tuple.Item1;
                            //I'm going to use the RPParameters from the "EmployeePeriod" report for now at least.
                            //RPParameters rpParameters = tuple.Item2;
                            CreateYTDCSV(xdoc, rpEmployeeYtdList, rpParameters);
                        }
                        catch (Exception ex)
                        {
                            textLine = string.Format("Error producing the employee ytd report for file {0}.\r\n{1}.\r\n", file, ex);
                            prWG.Update_Progress(textLine, configDirName);
                            eeYtdProcessed = false;
                        }
                    }
                    
                }
                else if(file.Name.StartsWith("RTI-Re"))
                {
                    try
                    {
                        prWG.ArchiveRTIOutputs(directory, file);
                        textLine = string.Format("Successfully archived RTI for file {0}.", file);
                        prWG.Update_Progress(textLine, configDirName);
                    }
                    catch(Exception ex)
                    {
                        textLine = string.Format("Error archiving RTI for file {0}.\r\n{1}.\r\n", file, ex);
                        prWG.Update_Progress(textLine, configDirName);
                    }
                }
            }
            files = dirInfo.GetFiles();
            if(files.Count() == 0)
            {
                dirInfo.Delete();
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
        private void CombineXmlFiles(string directory, string fileType)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(directory);
            
            string type = null;
            string masterFileName = null;
            if (fileType == "Period")
            {
                type = "EmployeePeriod*.xml";
                masterFileName = directory + "\\" + "CompleteEmployeePeriod.xml";
            }
            else
            {
                type = "EmployeeYtd*.xml";
                masterFileName = directory + "\\" + "CompleteEmployeeYtd.xml";
            }

            FileInfo[] files = dirInfo.GetFiles(type);
            int noOfFiles = files.Count();
            if(noOfFiles == 1)
            {
                FileInfo firstFile = files.First();
                firstFile.MoveTo(masterFileName);
            }
            else if(noOfFiles > 1)
            {
                int i = 0;
                XmlDocument masterFile = new XmlDocument();
                XmlDocument nextFile = new XmlDocument();
                foreach (FileInfo file in files.OrderBy(x => x.Name))
                {
                    if (i == 0)
                    {
                        masterFile.Load(file.FullName);
                        i++;
                    }
                    else
                    {
                        nextFile.Load(file.FullName);
                        foreach (XmlNode childNode in nextFile.DocumentElement.ChildNodes)
                        {
                            if (childNode.Name == "Employees")
                            {
                                var newNode = masterFile.ImportNode(childNode, true);
                                masterFile.DocumentElement.AppendChild(newNode);
                            }

                        }
                    }
                    file.Delete();
                }

                masterFile.Save(masterFileName);
            }
            
            
           
        }
        public Tuple<List<RPEmployeeYtd>, RPParameters> PrepareYTDReport(XDocument xdoc, FileInfo file)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            XmlDocument xmlYTDReport = new XmlDocument();
            xmlYTDReport.Load(file.FullName);

            //Now extract the necessary data and produce the required reports.

            RPParameters rpParameters = prWG.GetRPParameters(xmlYTDReport);
            List<RPEmployeeYtd> rpEmployeeYtdList = PrepareYTDCSV(xdoc, xmlYTDReport);

            return new Tuple<List<RPEmployeeYtd>, RPParameters>(rpEmployeeYtdList, rpParameters);
        }
        private List<RPEmployeeYtd> PrepareYTDCSV(XDocument xdoc, XmlDocument xmlReport)
        {
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            List<RPEmployeeYtd> rpEmployeeYtdList = new List<RPEmployeeYtd>();

            foreach (XmlElement employee in xmlReport.GetElementsByTagName("Employee"))
            {
                bool include = false;
                if (prWG.GetElementByTagFromXml(employee, "PayRunDate") != "No Pay Run Data Found")
                {
                    //If the employee is a leaver before the start date then don't include.
                    string leaver = prWG.GetElementByTagFromXml(employee, "Leaver");
                    DateTime leavingDate = new DateTime();
                    if (prWG.GetElementByTagFromXml(employee, "LeavingDate") != "")
                    {
                        leavingDate = DateTime.ParseExact(prWG.GetElementByTagFromXml(employee, "LeavingDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                    }
                    DateTime periodStartDate = DateTime.ParseExact(prWG.GetElementByTagFromXml(employee, "ThisPeriodStartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    //It seems they want to include leaver in the YTD csv file. I think this might change!
                    include = true;
                    //if (leaver.StartsWith("N"))
                    //{
                    //    include = true;
                    //}
                    //else if (leavingDate >= periodStartDate)
                    //{
                    //    include = true;
                    //}

                }

                if (include)
                {
                    RPEmployeeYtd rpEmployeeYtd = new RPEmployeeYtd();

                    rpEmployeeYtd.ThisPeriodStartDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "ThisPeriodStartDate"));
                    rpEmployeeYtd.LastPaymentDate = Convert.ToDateTime(prWG.GetDateElementByTagFromXml(employee, "LastPaymentDate"));
                    rpEmployeeYtd.EeRef = prWG.GetElementByTagFromXml(employee, "EeRef");
                    rpEmployeeYtd.Branch = prWG.GetElementByTagFromXml(employee, "Branch");
                    rpEmployeeYtd.CostCentre = prWG.GetElementByTagFromXml(employee, "CostCentre");
                    rpEmployeeYtd.Department = prWG.GetElementByTagFromXml(employee, "Department");
                    rpEmployeeYtd.LeavingDate = prWG.GetDateElementByTagFromXml(employee, "LeavingDate");
                    rpEmployeeYtd.Leaver = prWG.GetBooleanElementByTagFromXml(employee, "Leaver");
                    rpEmployeeYtd.TaxPrevEmployment = prWG.GetDecimalElementByTagFromXml(employee, "TaxPrevEmployment");
                    rpEmployeeYtd.TaxablePayPrevEmployment = prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayPrevEmployment");
                    rpEmployeeYtd.TaxThisEmployment = prWG.GetDecimalElementByTagFromXml(employee, "TaxThisEmployment");
                    rpEmployeeYtd.TaxablePayThisEmployment = prWG.GetDecimalElementByTagFromXml(employee, "TaxablePayThisEmployment");
                    rpEmployeeYtd.GrossedUp = prWG.GetDecimalElementByTagFromXml(employee, "GrossedUp");
                    rpEmployeeYtd.GrossedUpTax = prWG.GetDecimalElementByTagFromXml(employee, "GrossedUpTax");
                    rpEmployeeYtd.NetPayYTD = prWG.GetDecimalElementByTagFromXml(employee, "NetPayYTD");
                    rpEmployeeYtd.GrossPayYTD = prWG.GetDecimalElementByTagFromXml(employee, "GrossPayYTD");
                    rpEmployeeYtd.BenefitInKindYTD = prWG.GetDecimalElementByTagFromXml(employee, "BenefitInKindYTD");
                    rpEmployeeYtd.SuperannuationYTD = prWG.GetDecimalElementByTagFromXml(employee, "Superannuation");
                    rpEmployeeYtd.HolidayPayYTD = prWG.GetDecimalElementByTagFromXml(employee, "HolidayPayYTD");
                    rpEmployeeYtd.PensionablePayYtd = 0;
                    rpEmployeeYtd.EePensionYtd = 0;
                    rpEmployeeYtd.ErPensionYtd = 0;
                    List<RPPensionYtd> rpPensionsYtd = new List<RPPensionYtd>();
                    foreach (XmlElement pension in employee.GetElementsByTagName("Pension"))
                    {
                        RPPensionYtd rpPensionYtd = new RPPensionYtd();
                        rpPensionYtd.Key = Convert.ToInt32(pension.GetAttribute("Key"));
                        rpPensionYtd.Code = prWG.GetElementByTagFromXml(pension, "Code");
                        rpPensionYtd.SchemeName = prWG.GetElementByTagFromXml(pension, "SchemeName");
                        rpPensionYtd.PensionablePayYtd = prWG.GetDecimalElementByTagFromXml(pension, "PensionablePayYtd");
                        rpPensionYtd.EePensionYtd = prWG.GetDecimalElementByTagFromXml(pension, "EePensionYtd");
                        rpPensionYtd.ErPensionYtd = prWG.GetDecimalElementByTagFromXml(pension, "ErPensionYtd");

                        rpEmployeeYtd.PensionablePayYtd += rpPensionYtd.PensionablePayYtd;
                        rpEmployeeYtd.EePensionYtd += rpPensionYtd.EePensionYtd;
                        rpEmployeeYtd.ErPensionYtd += rpPensionYtd.ErPensionYtd;

                        rpPensionsYtd.Add(rpPensionYtd);
                    }
                    rpEmployeeYtd.Pensions = rpPensionsYtd;

                    rpEmployeeYtd.AeoYTD = prWG.GetDecimalElementByTagFromXml(employee, "AeoYTD");
                    rpEmployeeYtd.StudentLoanStartDate = prWG.GetDateElementByTagFromXml(employee, "StudentLoanStartDate");
                    rpEmployeeYtd.StudentLoanEndDate = prWG.GetDateElementByTagFromXml(employee, "StudentLoanEndDate");
                    rpEmployeeYtd.StudentLoanPlanType = prWG.GetElementByTagFromXml(employee, "StudentLoanPlanType");
                    rpEmployeeYtd.StudentLoanDeductionsYTD = prWG.GetDecimalElementByTagFromXml(employee, "StudentLoanDeductionsYTD");
                    rpEmployeeYtd.PostgraduateLoanStartDate = prWG.GetDateElementByTagFromXml(employee, "PostgraduateLoanStartDate");
                    rpEmployeeYtd.PostgraduateLoanEndDate = prWG.GetDateElementByTagFromXml(employee, "PostgraduateLoanEndDate");
                    rpEmployeeYtd.PostgraduateLoanDeductionsYTD = prWG.GetDecimalElementByTagFromXml(employee, "PostgraduateLoanDeductionsYTD");

                    foreach (XmlElement nicYtd in employee.GetElementsByTagName("NicYtd"))
                    {
                        RPNicYtd rpNicYtd = new RPNicYtd();
                        rpNicYtd.NILetter = nicYtd.GetAttribute("NiLetter");
                        rpNicYtd.NiableYtd = prWG.GetDecimalElementByTagFromXml(nicYtd, "NiableYtd");
                        rpNicYtd.EarningsToLEL = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToLEL");
                        rpNicYtd.EarningsToSET = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToSET");
                        rpNicYtd.EarningsToPET = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToPET");
                        rpNicYtd.EarningsToUST = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToUST");
                        rpNicYtd.EarningsToAUST = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToAUST");
                        rpNicYtd.EarningsToUEL = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsToUEL");
                        rpNicYtd.EarningsAboveUEL = prWG.GetDecimalElementByTagFromXml(nicYtd, "EarningsAboveUEL");
                        rpNicYtd.EeContributionsPt1 = prWG.GetDecimalElementByTagFromXml(nicYtd, "EeContributionsPt1");
                        rpNicYtd.EeContributionsPt2 = prWG.GetDecimalElementByTagFromXml(nicYtd, "EeContributionsPt2");
                        rpNicYtd.ErContributions = prWG.GetDecimalElementByTagFromXml(nicYtd, "ErContributions");
                        rpNicYtd.EeRebate = prWG.GetDecimalElementByTagFromXml(nicYtd, "EeRebate");
                        rpNicYtd.ErRebate = prWG.GetDecimalElementByTagFromXml(nicYtd, "ErRebate");
                        rpNicYtd.EeReduction = prWG.GetDecimalElementByTagFromXml(nicYtd, "EeReduction");
                        rpNicYtd.ErReduction = prWG.GetDecimalElementByTagFromXml(nicYtd, "ErReduction");

                        rpEmployeeYtd.NicYtd = rpNicYtd;
                    }
                    foreach (XmlElement nicAccountingPeriod in employee.GetElementsByTagName("NicAccountingPeriod"))
                    {
                        RPNicAccountingPeriod rpNicAccountingPeriod = new RPNicAccountingPeriod();
                        rpNicAccountingPeriod.NILetter = nicAccountingPeriod.GetAttribute("NiLetter");
                        rpNicAccountingPeriod.NiableYtd = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "NiableYtd");
                        rpNicAccountingPeriod.EarningsToLEL = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToLEL");
                        rpNicAccountingPeriod.EarningsToSET = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToSET");
                        rpNicAccountingPeriod.EarningsToPET = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToPET");
                        rpNicAccountingPeriod.EarningsToUST = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToUST");
                        rpNicAccountingPeriod.EarningsToAUST = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToAUST");
                        rpNicAccountingPeriod.EarningsToUEL = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsToUEL");
                        rpNicAccountingPeriod.EarningsAboveUEL = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EarningsAboveUEL");
                        rpNicAccountingPeriod.EeContributionsPt1 = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeContributionsPt1");
                        rpNicAccountingPeriod.EeContributionsPt2 = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeContributionsPt2");
                        rpNicAccountingPeriod.ErContributions = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErContributions");
                        rpNicAccountingPeriod.EeRebate = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeRebate");
                        rpNicAccountingPeriod.ErRebate = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErRebate");
                        rpNicAccountingPeriod.EeReduction = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "EeReduction");
                        rpNicAccountingPeriod.ErReduction = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErReduction");

                        rpNicAccountingPeriod.ErReduction = prWG.GetDecimalElementByTagFromXml(nicAccountingPeriod, "ErReduction");

                        rpEmployeeYtd.NicAccountingPeriod = rpNicAccountingPeriod;
                    }

                    rpEmployeeYtd.TaxCode = prWG.GetElementByTagFromXml(employee, "TaxCode");
                    rpEmployeeYtd.Week1Month1 = prWG.GetBooleanElementByTagFromXml(employee, "Week1Month1");
                    rpEmployeeYtd.WeekNumber = prWG.GetIntElementByTagFromXml(employee, "WeekNumber");
                    rpEmployeeYtd.MonthNumber = prWG.GetIntElementByTagFromXml(employee, "MonthNumber");
                    rpEmployeeYtd.PeriodNumber = prWG.GetIntElementByTagFromXml(employee, "PeriodNumber");
                    rpEmployeeYtd.EeNiPaidByErAccountsAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsAmount");
                    rpEmployeeYtd.EeNiPaidByErAccountsUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeNiPaidByErAccountsUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErAccountsUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                    rpEmployeeYtd.EeNiLERtoUERAccountsAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                    rpEmployeeYtd.EeNiLERtoUERAccountsUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                    rpEmployeeYtd.ErNiAccountsAmount = prWG.GetDecimalElementByTagFromXml(employee, "ErNiAccountAmount");
                    rpEmployeeYtd.ErNiAccountsUnits = prWG.GetDecimalElementByTagFromXml(employee, "ErNiAccountUnit");
                    rpEmployeeYtd.EeNiLERtoUERPayeAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                    rpEmployeeYtd.EeNiLERtoUERPayeUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                    rpEmployeeYtd.EeNiPaidByErPayeAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeAmount");
                    rpEmployeeYtd.EeNiPaidByErPayeUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeNiPaidByErPayeUnits");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeAmount = prWG.GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                    rpEmployeeYtd.EeGuTaxPaidByErPayeUnits = prWG.GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                    rpEmployeeYtd.ErNiPayeAmount = prWG.GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                    rpEmployeeYtd.ErNiPayeUnits = prWG.GetDecimalElementByTagFromXml(employee, "ErNiPayeUnit");

                    //Find the pension pay codes
                    rpEmployeeYtd.PensionPreTaxEeAccounts = 0;
                    rpEmployeeYtd.PensionPreTaxEePaye = 0;
                    rpEmployeeYtd.PensionPostTaxEeAccounts = 0;
                    rpEmployeeYtd.PensionPostTaxEePaye = 0;
                    foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                    {
                        foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                        {
                            string pensionCode = prWG.GetElementByTagFromXml(payCode, "Code");
                            if (pensionCode.StartsWith("PENSION"))
                            {
                                if (pensionCode == "PENSIONRAS" || pensionCode == "PENSIONTAXEX")
                                {
                                    rpEmployeeYtd.PensionPostTaxEeAccounts += prWG.GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                                    rpEmployeeYtd.PensionPostTaxEePaye += prWG.GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                                }
                                else
                                {
                                    rpEmployeeYtd.PensionPreTaxEeAccounts += prWG.GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                                    rpEmployeeYtd.PensionPreTaxEePaye += prWG.GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                                }
                            }

                        }
                    }
                    rpEmployeeYtd.PensionPreTaxEeAccounts *= -1;
                    rpEmployeeYtd.PensionPreTaxEePaye *= -1;
                    rpEmployeeYtd.PensionPostTaxEeAccounts *= -1;
                    rpEmployeeYtd.PensionPostTaxEePaye *= -1;

                    //These next few fields get treated like pay codes. Use them if they are not zero.
                    //7 pay components EeNiPaidByEr, EeGuTaxPaidByEr, EeNiLERtoUER & ErNi
                    List<RPPayCode> rpPayCodeList = new List<RPPayCode>();

                    for (int i = 0; i < 7; i++)
                    {
                        RPPayCode rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";

                        switch (i)
                        {
                            case 0:
                                rpPayCode.PayCode = "EeNIPdByEr";
                                rpPayCode.Description = "Ee NI Paid By Er";
                                rpPayCode.Type = "E";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeNiPaidByErAccountsAmount;
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeNiPaidByErPayeAmount;
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeNiPaidByErAccountsUnits;
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeNiPaidByErPayeUnits;
                                rpPayCode.IsPayCode = false;
                                break;
                            case 1:
                                rpPayCode.PayCode = "GUTax";
                                rpPayCode.Description = "Grossed up Tax";
                                rpPayCode.Type = "E";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeGuTaxPaidByErAccountsAmount;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeGuTaxPaidByErPayeAmount;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeGuTaxPaidByErAccountsUnits;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErAccountsUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeGuTaxPaidByErPayeUnits;//GetDecimalElementByTagFromXml(employee, "EeGuTaxPaidByErPayeUnit");
                                rpPayCode.IsPayCode = false;
                                break;
                            case 2:
                                rpPayCode.PayCode = "NIEeeLERtoUER";
                                rpPayCode.Description = "NIEeeLERtoUER-A";
                                rpPayCode.Type = "T";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.EeNiLERtoUERAccountsAmount;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.EeNiLERtoUERPayeAmount;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.EeNiLERtoUERAccountsUnits;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERAccountsUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.EeNiLERtoUERPayeUnits;//GetDecimalElementByTagFromXml(employee, "EeNiLERtoUERPayeUnit");
                                rpPayCode.IsPayCode = false;
                                break;
                            case 3:
                                rpPayCode.PayCode = "NIEr";
                                rpPayCode.Description = "NIEr-A";
                                rpPayCode.Type = "T";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.ErNiAccountsAmount;//GetDecimalElementByTagFromXml(employee, "ErNiAccountAmount");
                                rpPayCode.PayeAmount = rpEmployeeYtd.ErNiPayeAmount;//GetDecimalElementByTagFromXml(employee, "ErNiPayeAmount");
                                rpPayCode.AccountsUnits = rpEmployeeYtd.ErNiAccountsUnits;//GetDecimalElementByTagFromXml(employee, "ErNiAccountUnit");
                                rpPayCode.PayeUnits = rpEmployeeYtd.ErNiPayeUnits;//GetDecimalElementByTagFromXml(employee, "ErNiPayeUnit");
                                rpPayCode.IsPayCode = false;
                                break;
                            case 4:
                                rpPayCode.PayCode = "PenEr";
                                rpPayCode.Description = "PenEr";
                                rpPayCode.Type = "M";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.ErPensionYtd;//GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.ErPensionYtd;//GetDecimalElementByTagFromXml(employee, "ErPensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                rpPayCode.IsPayCode = false;
                                break;
                            case 5:
                                rpPayCode.PayCode = "PenPreTaxEe";
                                rpPayCode.Description = "PenPreTaxEe";
                                rpPayCode.Type = "D";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.PensionPreTaxEeAccounts;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.PensionPreTaxEePaye;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                rpPayCode.IsPayCode = false;
                                break;
                            default:
                                rpPayCode.PayCode = "PenPostTaxEe";
                                rpPayCode.Description = "PenPostTaxEe";
                                rpPayCode.Type = "D";
                                rpPayCode.TotalAmount = 0;
                                rpPayCode.AccountsAmount = rpEmployeeYtd.PensionPostTaxEeAccounts;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.PayeAmount = rpEmployeeYtd.PensionPostTaxEePaye;//GetDecimalElementByTagFromXml(employee, "EePensionYTD");
                                rpPayCode.AccountsUnits = 0;
                                rpPayCode.PayeUnits = 0;
                                rpPayCode.IsPayCode = false;
                                break;
                        }

                        //
                        //Check if any of the values are not zero. If so write the first employee record
                        //
                        bool allZeros = false;
                        if (rpPayCode.AccountsAmount == 0 && rpPayCode.AccountsUnits == 0 &&
                            rpPayCode.PayeUnits == 0 && rpPayCode.PayeUnits == 0)
                        {
                            allZeros = true;

                        }
                        if (!allZeros)
                        {
                            //Add employee record to the list
                            rpPayCodeList.Add(rpPayCode);
                            //rpEmployeeYtd.PayCodes.Add(rpPayCode);
                        }
                    }
                    //Add in the pension schemes
                    foreach (RPPensionYtd rpPensionYtd in rpEmployeeYtd.Pensions)
                    {
                        //Ee pension
                        RPPayCode rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Ee";
                        rpPayCode.Description = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName;
                        rpPayCode.Type = "P";
                        rpPayCode.TotalAmount = 0;
                        rpPayCode.AccountsAmount = rpPensionYtd.EePensionYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.EePensionYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;
                        rpPayCode.IsPayCode = false;

                        rpPayCodeList.Add(rpPayCode);

                        //Er pension
                        rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Er";
                        rpPayCode.Description = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName;
                        rpPayCode.Type = "P";
                        rpPayCode.TotalAmount = 0;
                        rpPayCode.AccountsAmount = rpPensionYtd.ErPensionYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.ErPensionYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;
                        rpPayCode.IsPayCode = false;

                        rpPayCodeList.Add(rpPayCode);

                        //Pensionable pay
                        rpPayCode = new RPPayCode();

                        rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                        rpPayCode.Code = "0";
                        rpPayCode.PayCode = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName + "-Pay";
                        rpPayCode.Description = rpPensionYtd.Code + "-" + rpPensionYtd.SchemeName;
                        rpPayCode.Type = "P";
                        rpPayCode.TotalAmount = 0;
                        rpPayCode.AccountsAmount = rpPensionYtd.PensionablePayYtd;
                        rpPayCode.PayeAmount = rpPensionYtd.PensionablePayYtd;
                        rpPayCode.AccountsUnits = 0;
                        rpPayCode.PayeUnits = 0;
                        rpPayCode.IsPayCode = false;

                        rpPayCodeList.Add(rpPayCode);
                    }

                    foreach (XmlElement payCodes in employee.GetElementsByTagName("PayCodes"))
                    {
                        foreach (XmlElement payCode in payCodes.GetElementsByTagName("PayCode"))
                        {
                            RPPayCode rpPayCode = new RPPayCode();

                            rpPayCode.EeRef = rpEmployeeYtd.EeRef;
                            rpPayCode.Code = prWG.GetElementByTagFromXml(payCode, "Code");
                            rpPayCode.PayCode = prWG.GetElementByTagFromXml(payCode, "Code");
                            rpPayCode.Description = prWG.GetElementByTagFromXml(payCode, "Description");
                            rpPayCode.IsPayCode = prWG.GetBooleanElementByTagFromXml(payCode, "IsPayCode");
                            //If the pay component is one of the statutory absences set IsPayCode to false regardless of what's in the file.
                            rpPayCode.IsPayCode = SetStatutoryAbsenceIsPayCode(rpPayCode.IsPayCode, rpPayCode.PayCode);
                            rpPayCode.Type = prWG.GetElementByTagFromXml(payCode, "EarningOrDeduction");
                            rpPayCode.TotalAmount = prWG.GetDecimalElementByTagFromXml(payCode, "TotalAmount");
                            rpPayCode.AccountsAmount = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsAmount");
                            rpPayCode.PayeAmount = prWG.GetDecimalElementByTagFromXml(payCode, "PayeAmount");
                            rpPayCode.AccountsUnits = prWG.GetDecimalElementByTagFromXml(payCode, "AccountsUnits");
                            rpPayCode.PayeUnits = prWG.GetDecimalElementByTagFromXml(payCode, "PayeUnits");
                            

                            //
                            //Check if any of the values are not zero. If so write the first employee record
                            //
                            bool allZeros = false;
                            if (rpPayCode.AccountsAmount == 0 && rpPayCode.AccountsUnits == 0 &&
                                rpPayCode.PayeAmount == 0 && rpPayCode.PayeUnits == 0)
                            {
                                allZeros = true;

                            }
                            if (!allZeros)
                            {
                                //I don't require TAX, NI or PENSION
                                if (rpPayCode.Code != "TAX" && rpPayCode.Code != "NI" && !rpPayCode.Code.StartsWith("PENSION"))
                                {
                                    if (rpPayCode.Type == "D")
                                    {
                                        //Deduction so multiply by -1
                                        rpPayCode.AccountsAmount *= -1;
                                        rpPayCode.PayeAmount *= -1;
                                        rpPayCode.TotalAmount *= -1;

                                    }
                                    if (rpPayCode.Code == "UNPDM")
                                    {
                                        //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                        rpPayCode.PayCode = "UNPD£";
                                    }
                                    if(rpPayCode.Code=="AOE")
                                    {
                                        RPPayCode aoePayCode = new RPPayCode();
                                        aoePayCode = GetRPPayCode(rpPayCode);
                                        //For an AOE we need to create 3 rows in the Ytd csv file.
                                        aoePayCode.PayCode = aoePayCode.PayCode + " " + aoePayCode.Description;
                                        rpPayCodeList.Add(aoePayCode);
                                        //PaidTD
                                        aoePayCode = new RPPayCode();
                                        aoePayCode = GetRPPayCode(rpPayCode);
                                        aoePayCode.Type = "A";
                                        string reference = null;
                                        string name = null;
                                        int i = aoePayCode.Description.IndexOf('-');
                                        reference = aoePayCode.Description.Substring(0, i + 1);
                                        name = aoePayCode.Description.Substring(i + 1);
                                        aoePayCode.Description = name + reference + "PaidTD";
                                        aoePayCode.PayCode = aoePayCode.Description;
                                        rpPayCodeList.Add(aoePayCode);
                                        //PayYTD
                                        rpPayCode.Type = "A";
                                        rpPayCode.Description = name + reference + "PayYTD";
                                        rpPayCode.PayCode = rpPayCode.Description;
                                    }
                                    //Add to employee record
                                    rpPayCodeList.Add(rpPayCode);
                                    //rpEmployeeYtd.PayCodes.Add(rpPayCode);
                                }



                            }

                        }
                        rpEmployeeYtd.PayCodes = rpPayCodeList;
                    }
                    rpEmployeeYtdList.Add(rpEmployeeYtd);
                }

            }
            //Sort the list of employees into EeRef sequence before returning them.
            rpEmployeeYtdList.Sort(delegate (RPEmployeeYtd x, RPEmployeeYtd y)
            {
                if (x.EeRef == null && y.EeRef == null) return 0;
                else if (x.EeRef == null) return -1;
                else if (y.EeRef == null) return 1;
                else return x.EeRef.CompareTo(y.EeRef);
            });

            return rpEmployeeYtdList;
        }
        private RPPayCode GetRPPayCode(RPPayCode rpPayCode)
        {
            RPPayCode aoePayCode = new RPPayCode();
            aoePayCode.AccountsAmount = rpPayCode.AccountsAmount;
            aoePayCode.AccountsUnits = rpPayCode.AccountsUnits;
            aoePayCode.Code = rpPayCode.Code;
            aoePayCode.Description = rpPayCode.Description;
            aoePayCode.EeRef = rpPayCode.EeRef;
            aoePayCode.IsPayCode = rpPayCode.IsPayCode;
            aoePayCode.PayCode = rpPayCode.PayCode;
            aoePayCode.PayeAmount = rpPayCode.PayeAmount;
            aoePayCode.PayeUnits = rpPayCode.PayeUnits;
            aoePayCode.TotalAmount = rpPayCode.TotalAmount;
            aoePayCode.Type = rpPayCode.Type;
            return aoePayCode;
        }
        public void CreateYTDCSV(XDocument xdoc, List<RPEmployeeYtd> rpEmployeeYtdList, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";

            string coNo = rpParameters.ErRef;
            //Create csv version and write it to the same folder.
            //string csvFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_YearToDates_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            //string csvFileName = outgoingFolder + "\\" + coNo + "_" + rpParameters.PayRunDate.ToString("yyyyMMdd") + "\\" + coNo + "_YearToDates_" +
            //                                      rpParameters.PayRunDate.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmssfff") + ".csv"; //DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            //Emer doesn't want multiple versions of the csv files being sent so if I remove the timestamp that should do it.
            string csvFileName = outgoingFolder + "\\" + coNo + "_" + rpParameters.PayRunDate.ToString("yyyyMMdd") + "\\" + coNo + "_YearToDates_" +
                                                  rpParameters.PayRunDate.ToString("yyyyMMdd") + ".csv";
            bool writeHeader = true;
            using (StreamWriter sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payYTDDetails = new string[45];


                foreach (RPEmployeeYtd rpEmployeeYtd in rpEmployeeYtdList)
                {
                    payYTDDetails[0] = rpEmployeeYtd.LastPaymentDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    //I'm using the rpParameters from the "EmployeePeriod" report.
                    payYTDDetails[0] = rpParameters.PayRunDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    payYTDDetails[1] = rpEmployeeYtd.EeRef;
                    if (rpEmployeeYtd.LeavingDate != null)
                    {
                        payYTDDetails[2] = Convert.ToDateTime(rpEmployeeYtd.LeavingDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[2] = "";
                    }
                    if (rpEmployeeYtd.Leaver)
                    {
                        payYTDDetails[3] = "Y";
                    }
                    else
                    {
                        payYTDDetails[3] = "N";
                    }
                    payYTDDetails[4] = rpEmployeeYtd.TaxPrevEmployment.ToString();
                    payYTDDetails[5] = rpEmployeeYtd.TaxablePayPrevEmployment.ToString();
                    payYTDDetails[6] = rpEmployeeYtd.TaxThisEmployment.ToString();
                    payYTDDetails[7] = rpEmployeeYtd.TaxablePayThisEmployment.ToString();
                    payYTDDetails[8] = rpEmployeeYtd.GrossedUp.ToString();
                    payYTDDetails[9] = rpEmployeeYtd.GrossedUpTax.ToString();
                    payYTDDetails[10] = rpEmployeeYtd.NetPayYTD.ToString();
                    payYTDDetails[11] = (rpEmployeeYtd.TaxablePayPrevEmployment + rpEmployeeYtd.TaxablePayThisEmployment).ToString(); //rpEmployeeYtd.GrossPayYTD.ToString();
                    payYTDDetails[12] = rpEmployeeYtd.BenefitInKindYTD.ToString();
                    payYTDDetails[13] = rpEmployeeYtd.SuperannuationYTD.ToString();
                    payYTDDetails[14] = rpEmployeeYtd.HolidayPayYTD.ToString();
                    //Add the pensions from the list of pensions
                    decimal erPensionYtd = 0;
                    decimal eePensionYtd = 0;
                    foreach (RPPensionYtd pensionYtd in rpEmployeeYtd.Pensions)
                    {
                        erPensionYtd += pensionYtd.ErPensionYtd;
                        eePensionYtd += pensionYtd.EePensionYtd;
                    }
                    payYTDDetails[15] = erPensionYtd.ToString();
                    payYTDDetails[16] = eePensionYtd.ToString();
                    payYTDDetails[17] = rpEmployeeYtd.AeoYTD.ToString();
                    if (rpEmployeeYtd.StudentLoanStartDate != null)
                    {
                        payYTDDetails[18] = Convert.ToDateTime(rpEmployeeYtd.StudentLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[18] = "";
                    }
                    if (rpEmployeeYtd.StudentLoanEndDate != null)
                    {
                        payYTDDetails[19] = Convert.ToDateTime(rpEmployeeYtd.StudentLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        payYTDDetails[19] = "";
                    }
                    payYTDDetails[20] = rpEmployeeYtd.StudentLoanDeductionsYTD.ToString();
                    if(rpEmployeeYtd.NicYtd == null)
                    {
                        payYTDDetails[21] = "";
                        payYTDDetails[22] = "0";
                        payYTDDetails[23] = "0";
                        payYTDDetails[24] = "0";
                        payYTDDetails[25] = "0";
                        payYTDDetails[26] = "0";
                        payYTDDetails[27] = "0";
                        payYTDDetails[28] = "0";
                        payYTDDetails[29] = "0";
                        payYTDDetails[30] = "0";
                        payYTDDetails[31] = "0";
                        payYTDDetails[32] = "0";
                        payYTDDetails[33] = "0";
                        payYTDDetails[34] = "0";
                        payYTDDetails[35] = "0";
                        payYTDDetails[40] = "0";
                    }
                    else
                    {
                        payYTDDetails[21] = rpEmployeeYtd.NicYtd.NILetter;
                        payYTDDetails[22] = rpEmployeeYtd.NicYtd.NiableYtd.ToString();
                        payYTDDetails[23] = rpEmployeeYtd.NicYtd.EarningsToLEL.ToString();
                        payYTDDetails[24] = rpEmployeeYtd.NicYtd.EarningsToSET.ToString();
                        payYTDDetails[25] = rpEmployeeYtd.NicYtd.EarningsToPET.ToString();
                        payYTDDetails[26] = rpEmployeeYtd.NicYtd.EarningsToUST.ToString();
                        payYTDDetails[27] = rpEmployeeYtd.NicYtd.EarningsToAUST.ToString();
                        payYTDDetails[28] = rpEmployeeYtd.NicYtd.EarningsToUEL.ToString();
                        payYTDDetails[29] = rpEmployeeYtd.NicYtd.EarningsAboveUEL.ToString();
                        payYTDDetails[30] = rpEmployeeYtd.NicYtd.EeContributionsPt1.ToString();
                        payYTDDetails[31] = rpEmployeeYtd.NicYtd.EeContributionsPt2.ToString();
                        payYTDDetails[32] = rpEmployeeYtd.NicYtd.ErContributions.ToString();
                        payYTDDetails[33] = rpEmployeeYtd.NicYtd.EeRebate.ToString();
                        payYTDDetails[34] = rpEmployeeYtd.NicYtd.ErRebate.ToString();
                        payYTDDetails[35] = rpEmployeeYtd.NicYtd.EeReduction.ToString();
                        payYTDDetails[40] = rpEmployeeYtd.NicYtd.NiableYtd.ToString();
                    }
                    payYTDDetails[36] = rpEmployeeYtd.TaxCode;
                    if (rpEmployeeYtd.Week1Month1)
                    {
                        payYTDDetails[37] = "Y";
                    }
                    else
                    {
                        payYTDDetails[37] = "N";
                    }
                    payYTDDetails[38] = rpEmployeeYtd.WeekNumber.ToString();
                    payYTDDetails[39] = rpEmployeeYtd.MonthNumber.ToString();
                    
                    switch (rpEmployeeYtd.StudentLoanPlanType)
                    {
                        case "Plan1":
                            payYTDDetails[41] = "01";
                            break;
                        case "Plan2":
                            payYTDDetails[41] = "02";
                            break;
                        case "Plan3":
                            payYTDDetails[41] = "03";
                            break;
                        case "Plan4":
                            payYTDDetails[41] = "04";
                            break;
                        default:
                            payYTDDetails[41] = "";
                            break;
                    }
                    if (rpEmployeeYtd.PostgraduateLoanStartDate != null)
                    {
                        payYTDDetails[42] = Convert.ToDateTime(rpEmployeeYtd.PostgraduateLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture); //Postgraduate Loan Start Date
                    }
                    else
                    {
                        payYTDDetails[42] = "";
                    }
                    if (rpEmployeeYtd.PostgraduateLoanEndDate != null)
                    {
                        payYTDDetails[43] = Convert.ToDateTime(rpEmployeeYtd.PostgraduateLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture); //Postgraduate Loan End Date
                    }
                    else
                    {
                        payYTDDetails[43] = "";
                    }

                    payYTDDetails[44] = rpEmployeeYtd.PostgraduateLoanDeductionsYTD.ToString(); //Postgraduate Loan Deducted

                    foreach (RPPayCode rpPayCode in rpEmployeeYtd.PayCodes)
                    {
                        //Don't use pay codes TAX, NI or any that begin with PENSION
                        if (rpPayCode.Code != "TAX" && rpPayCode.Code != "NI" && !rpPayCode.Code.StartsWith("PENSION"))
                        {
                            string[] payCodeDetails = new string[8];
                            //The BIKOFFSET IsPayCode marker should really be set to false. It is in the EEPERIOD report but not the EEYTD report.
                            if (rpPayCode.IsPayCode && rpPayCode.PayCode != "BIKOFFSET")
                            {
                                payCodeDetails[0] = "";
                            }
                            else
                            {
                                payCodeDetails[0] = "0";
                            }
                            payCodeDetails[1] = rpPayCode.Type;
                            payCodeDetails[2] = rpPayCode.PayCode;
                            payCodeDetails[3] = rpPayCode.Description;
                            payCodeDetails[4] = rpPayCode.AccountsAmount.ToString();
                            payCodeDetails[5] = rpPayCode.PayeAmount.ToString();
                            payCodeDetails[6] = rpPayCode.AccountsUnits.ToString();
                            payCodeDetails[7] = rpPayCode.PayeUnits.ToString();

                            switch (rpPayCode.Code)
                            {
                                case "UNPDM":
                                    //Change UNPDM back to UNPD£. WG uses UNPD£ PR doesn't like symbols like £ in pay codes.
                                    payCodeDetails[2] = "UNPD£";
                                    break;
                                case "SLOAN":
                                    payCodeDetails[2] = "StudentLoan";
                                    payCodeDetails[3] = "StudentLoan";
                                    break;
                                case "AOE":
                                    if(payCodeDetails[3].Contains("PaidTD"))
                                    {
                                        payCodeDetails[4] = rpPayCode.TotalAmount.ToString();
                                        payCodeDetails[5] = rpPayCode.TotalAmount.ToString();
                                    }
                                    break;
                            }

                            //Write employee record
                            WritePayYTDCSV(rpParameters, payYTDDetails, payCodeDetails, sw, writeHeader);
                            writeHeader = false;
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
                              "Taxable Pay Previous Emt,Tax This Emt,Taxable Pay This Emt,Grossed Up," +
                              "Grossed Up Tax,Net Pay,GrossYTD,Benefit in Kind,Superannuation," +
                              "Holiday Pay,ErPensionYTD,EePensionYTD,AEOYTD,StudentLoanStartDate," +
                              "StudentLoanEndDate,StudentLoanDeductions,NI Letter,Total," +
                              "Earnings To LEL,Earnings To SET,Earnings To PET,Earnings To UST," +
                              "Earnings To AUST,Earnings To UEL,Earnings Above UEL," +
                              "Ee Contributions Pt1,Ee Contributions Pt2,Er Contributions," +
                              "Ee Rebate,Er Rebate,Ee Reduction,PayCode,det,payCodeValue," +
                              "payCodeDesc,Acc Year Bal,PAYE Year Bal,Acc Year Units," +
                              "PAYE Year Units,Tax Code,Week1/Month 1,Week Number,Month Number," +
                              "NI Earnings YTD,Student Loan Plan Type,Postgraduate Loan Start Date," +
                              "Postgraduate Loan End Date,Postgraduate Loan Deducted";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "Fortnightly":
                    batch = "F";
                    break;
                case "FourWeekly":
                    batch = "FW";
                    break;
                case "Quarterly":
                    batch = "Q";
                    break;
                case "Yearly":
                    batch = "A";
                    break;
                default:
                    batch = "W";
                    break;
            }
            if (rpParameters.PaySchedule == "Monthly")
            {
                batch = "M";
            }

            string process;
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
            //From payYTDDetails[36] (TaxCode) to payYTDDetails[45] (Postgraduate Loan Deducted)
            for (int i = 36; i < 44; i++)
            {
                csvLine = csvLine + "\"" + payYTDDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }


        public void CreateHistoryCSV(XDocument xdoc, RPParameters rpParameters, RPEmployer rpEmployer, List<RPEmployeePeriod> rpEmployeePeriodList)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Outgoing";
            string coNo = rpParameters.ErRef;
            //Write the whole xml file to the folder.
            //string xmlFileName = "V:\\Payescape\\PayRunIO\\WG\\" + coNo + "_PayHistory_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml";
            string dirName = outgoingFolder + "\\" + coNo + "_" + rpParameters.PayRunDate.ToString("yyyyMMdd") + "\\";
            Directory.CreateDirectory(dirName);
            //Create csv version and write it to the same folder. 
            //Use the PayDate for the yyyyMMdd part of the name, then were going compare is to today's yyyyMMdd and only transfer it up to
            //the SFTP server if it's 1 day or less before today's date.
            //string csvFileName = outgoingFolder + "\\" + coNo + "_" + rpParameters.PayRunDate.ToString("yyyyMMdd") + "\\" + coNo + "_PayHistory_" +
            //                                      rpParameters.PayRunDate.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmssfff") + ".csv";//DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
            //Emer only wants one file sent so if I don't add a time stamp that should do it.
            string csvFileName = outgoingFolder + "\\" + coNo + "_" + rpParameters.PayRunDate.ToString("yyyyMMdd") + "\\" + coNo + "_PayHistory_" +
                                                  rpParameters.PayRunDate.ToString("yyyyMMdd") + ".csv";
            bool writeHeader = true;
            using (StreamWriter sw = new StreamWriter(csvFileName))
            {

                //Loop through each employee and write the csv file.
                string[] payHistoryDetails = new string[54];

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
                        payHistoryDetails[5] = rpEmployeePeriod.TaxablePayTP.ToString(); //rpEmployeePeriod.Gross.ToString();
                        payHistoryDetails[6] = rpEmployeePeriod.NetPayTP.ToString();
                        payHistoryDetails[7] = rpEmployeePeriod.DayHours.ToString();
                        if (rpEmployeePeriod.StudentLoanStartDate != null)
                        {
                            payHistoryDetails[8] = Convert.ToDateTime(rpEmployeePeriod.StudentLoanStartDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[8] = "";
                        }
                        if (rpEmployeePeriod.StudentLoanEndDate != null)
                        {
                            payHistoryDetails[9] = Convert.ToDateTime(rpEmployeePeriod.StudentLoanEndDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[9] = "";
                        }
                        //decimal studentLoan = rpEmployeePeriod.StudentLoan * -1;
                        //payHistoryDetails[10] = studentLoan.ToString();
                        payHistoryDetails[10] = rpEmployeePeriod.StudentLoanYtd.ToString();
                        payHistoryDetails[11] = rpEmployeePeriod.NILetter;
                        payHistoryDetails[12] = rpEmployeePeriod.CalculationBasis;
                        payHistoryDetails[13] = rpEmployeePeriod.Total.ToString();
                        payHistoryDetails[14] = rpEmployeePeriod.EarningsToLEL.ToString();
                        payHistoryDetails[15] = rpEmployeePeriod.EarningsToSET.ToString();
                        payHistoryDetails[16] = rpEmployeePeriod.EarningsToPET.ToString();
                        payHistoryDetails[17] = rpEmployeePeriod.EarningsToUST.ToString(); ;
                        payHistoryDetails[18] = rpEmployeePeriod.EarningsToAUST.ToString();
                        payHistoryDetails[19] = rpEmployeePeriod.EarningsToUEL.ToString();
                        payHistoryDetails[20] = rpEmployeePeriod.EarningsAboveUEL.ToString();
                        payHistoryDetails[21] = rpEmployeePeriod.EeContributionsPt1.ToString();
                        payHistoryDetails[22] = rpEmployeePeriod.EeContributionsPt2.ToString();
                        payHistoryDetails[23] = rpEmployeePeriod.ErNICYTD.ToString();
                        payHistoryDetails[24] = rpEmployeePeriod.EeRebate.ToString();
                        payHistoryDetails[25] = rpEmployeePeriod.ErRebate.ToString();
                        payHistoryDetails[26] = rpEmployeePeriod.EeReduction.ToString();
                        if (rpEmployeePeriod.LeavingDate != null)
                        {
                            payHistoryDetails[27] = Convert.ToDateTime(rpEmployeePeriod.LeavingDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            payHistoryDetails[27] = "";
                        }

                        if (rpEmployeePeriod.Leaver)
                        {
                            payHistoryDetails[28] = "Y";
                        }
                        else
                        {
                            payHistoryDetails[28] = "N";
                        }

                        payHistoryDetails[29] = rpEmployeePeriod.TaxCode.ToString();
                        if (rpEmployeePeriod.Week1Month1)
                        {
                            payHistoryDetails[30] = "Y";
                            //Remove the " W1" from the tax code
                            payHistoryDetails[29] = payHistoryDetails[29].Replace(" W1", "");
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
                        payHistoryDetails[36] = (rpEmployeePeriod.TaxablePayYTD - rpEmployeePeriod.TaxablePayPrevious).ToString();
                        payHistoryDetails[37] = rpEmployeePeriod.HolidayAccruedTd.ToString();

                        decimal erPensionYtd = 0;
                        decimal eePensionYtd = 0;
                        decimal erPensionTp = 0;
                        decimal eePensionTp = 0;
                        decimal erPensionPrd = 0;
                        decimal eePensionPrd = 0;
                        foreach (RPPensionPeriod pensionPeriod in rpEmployeePeriod.Pensions)
                        {
                            erPensionYtd += pensionPeriod.ErPensionYtd;
                            eePensionYtd += pensionPeriod.EePensionYtd;
                            erPensionTp += pensionPeriod.ErPensionTaxPeriod;
                            eePensionTp += pensionPeriod.EePensionTaxPeriod;
                            erPensionPrd += pensionPeriod.ErPensionPayRunDate;
                            eePensionPrd += pensionPeriod.EePensionPayRunDate;
                        }
                        payHistoryDetails[38] = erPensionYtd.ToString();
                        payHistoryDetails[39] = eePensionYtd.ToString();
                        payHistoryDetails[40] = erPensionTp.ToString();
                        payHistoryDetails[41] = eePensionTp.ToString();
                        payHistoryDetails[42] = erPensionPrd.ToString();
                        payHistoryDetails[43] = eePensionPrd.ToString();

                        payHistoryDetails[44] = rpEmployeePeriod.DirectorshipAppointmentDate.ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        if (rpEmployeePeriod.Director)
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
                        payHistoryDetails[46] = rpEmployeePeriod.EeContributionsTaxPeriodPt1.ToString();
                        payHistoryDetails[47] = rpEmployeePeriod.EeContributionsTaxPeriodPt2.ToString();
                        payHistoryDetails[48] = rpEmployeePeriod.ErNICTP.ToString();
                        if(rpEmployeePeriod.AEAssessment.AssessmentDate == null)
                        {
                            payHistoryDetails[49] = "";
                        }
                        else
                        {
                            payHistoryDetails[49] = Convert.ToDateTime(rpEmployeePeriod.AEAssessment.AssessmentDate).ToString("dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        payHistoryDetails[50] = rpEmployeePeriod.AEAssessment.AssessmentCode;
                        payHistoryDetails[51] = rpEmployeePeriod.AEAssessment.AssessmentEvent;
                        payHistoryDetails[52] = rpEmployeePeriod.AEAssessment.TaxPeriod.ToString();
                        payHistoryDetails[53] = rpEmployeePeriod.AEAssessment.TaxYear.ToString();

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
                                case 1:
                                    payCodeDetails[1] = "PenEr";
                                    payCodeDetails[2] = "PenEr";
                                    payCodeDetails[3] = "M";
                                    payCodeDetails[6] = erPensionTp.ToString();
                                    break;

                            }
                            payCodeDetails[0] = "0";
                            payCodeDetails[4] = "0";
                            payCodeDetails[5] = "0";
                            payCodeDetails[7] = "0";
                            payCodeDetails[8] = "0";
                            payCodeDetails[9] = "0";
                            payCodeDetails[10] = "0";
                            payCodeDetails[11] = "0";

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
                        foreach (RPAddition rpAddition in rpEmployeePeriod.Additions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[1] = rpAddition.Description;
                            payCodeDetails[2] = rpAddition.Code.TrimStart(' ');
                            payCodeDetails[3] = "E"; //Earnings
                            payCodeDetails[5] = rpAddition.Units.ToString();
                            payCodeDetails[6] = rpAddition.AmountTP.ToString();
                            if (rpAddition.IsPayCode)
                            {
                                payCodeDetails[0] = "";
                                if (rpAddition.Rate == 0)
                                {
                                    payCodeDetails[4] = rpAddition.AmountTP.ToString();  // Make Rate equal to amount if rate is zero.
                                }
                                else
                                {
                                    payCodeDetails[4] = rpAddition.Rate.ToString();
                                }
                                payCodeDetails[7] = rpAddition.AccountsYearBalance.ToString();
                                payCodeDetails[8] = rpAddition.AmountYTD.ToString();

                            }
                            else
                            {
                                payCodeDetails[0] = "0";
                                payCodeDetails[4] = "0";
                                payCodeDetails[7] = "0";
                                payCodeDetails[8] = "0";
                            }
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
                        decimal penPreAmount = 0, penPostAmount = 0;
                        bool wait = false;
                        foreach (RPDeduction rpDeduction in rpEmployeePeriod.Deductions)
                        {
                            string[] payCodeDetails = new string[12];
                            payCodeDetails = new string[12];
                            payCodeDetails[1] = rpDeduction.Description;
                            payCodeDetails[2] = rpDeduction.Code.TrimStart(' ');
                            payCodeDetails[3] = "D"; //Deductions


                            payCodeDetails[5] = rpDeduction.Units.ToString();
                            payCodeDetails[6] = rpDeduction.AmountTP.ToString();
                            if (rpDeduction.IsPayCode)
                            {
                                payCodeDetails[0] = "";
                                if (rpDeduction.Rate == 0)
                                {
                                    payCodeDetails[4] = rpDeduction.AmountTP.ToString();  // Make Rate equal to amount if rate is zero.
                                }
                                else
                                {
                                    payCodeDetails[4] = rpDeduction.Rate.ToString();
                                }
                                payCodeDetails[7] = rpDeduction.AccountsYearBalance.ToString();
                                payCodeDetails[8] = rpDeduction.AmountYTD.ToString();
                            }
                            else
                            {
                                payCodeDetails[0] = "0";                    // Pay code
                                payCodeDetails[4] = "0";                    // Rate
                                payCodeDetails[7] = "0";                    // Accounts Year Balance
                                payCodeDetails[8] = "0";                    // PAYE Year Balance
                            }
                            payCodeDetails[9] = rpDeduction.AccountsYearUnits.ToString();
                            payCodeDetails[10] = rpDeduction.PayeYearUnits.ToString();
                            payCodeDetails[11] = rpDeduction.PayrollAccrued.ToString();
                            switch (payCodeDetails[2]) //PayCode
                            {
                                case "TAX":
                                    //payCodeDetails[0] = "0";
                                    payCodeDetails[1] = payHistoryDetails[29];  // Tax Code
                                    payCodeDetails[2] = payHistoryDetails[29];  // Tax Code
                                    //payCodeDetails[4] = "0";                    // Rate
                                    //payCodeDetails[7] = "0";
                                    //payCodeDetails[8] = "0";
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "NI":
                                    //payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "NIEeeLERtoUER-A";      // Ee NI
                                    payCodeDetails[2] = "NIEeeLERtoUER";        // Ee NI
                                    //payCodeDetails[4] = "0";                    // Rate
                                    //payCodeDetails[7] = "0";
                                    //payCodeDetails[8] = "0";
                                    payCodeDetails[3] = "T";                    // Tax    
                                    break;
                                case "PENSION":
                                    penPreAmount = rpDeduction.AmountTP;
                                    wait = true;
                                    break;
                                case "PENSIONSS":
                                    //payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPreTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPreTaxEe";         // Ee Pension
                                    //payCodeDetails[4] = "0";                   // Rate 
                                    payCodeDetails[6] = (penPreAmount + rpDeduction.AmountTP).ToString();
                                    //payCodeDetails[7] = "0";
                                    //payCodeDetails[8] = "0";
                                    payCodeDetails[9] = "0";
                                    payCodeDetails[10] = "0";
                                    payCodeDetails[11] = "0";
                                    wait = false;
                                    break;
                                case "PENSIONRAS":
                                    penPostAmount = rpDeduction.AmountTP;
                                    wait = true;
                                    break;
                                case "PENSIONTAXEX":
                                    //payCodeDetails[0] = "0";
                                    payCodeDetails[1] = "PenPostTaxEe";         // Ee Pension
                                    payCodeDetails[2] = "PenPostTaxEe";         // Ee Pension
                                    //payCodeDetails[4] = "0";                    // Rate
                                    payCodeDetails[6] = (penPostAmount + rpDeduction.AmountTP).ToString();
                                    //payCodeDetails[7] = "0";
                                    //payCodeDetails[8] = "0";
                                    payCodeDetails[9] = "0";
                                    payCodeDetails[10] = "0";
                                    payCodeDetails[11] = "0";
                                    wait = false;
                                    break;
                                case "SLOAN":
                                    payCodeDetails[1] = "StudentLoan";
                                    payCodeDetails[2] = "StudentLoan";
                                    //payCodeDetails[4] = "0";                    // Rate
                                    break;
                                case "AOE":
                                    payCodeDetails[2] = payCodeDetails[2] +  " " + payCodeDetails[1]; //Code + Description
                                    break;
                            }
                            if (!wait)
                            {
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
                              "Acc Year Bal,PAYE Year Bal,Acc Year Units,PAYE Year Units,Payroll Accrued," +
                              "LastAutoEnrolmentAssessmentDate,AutoEnrolmentAssessment,AutoEnrolmentAssessmentEvent," +
                              "AssessmentTaxPeriod,AssessmentTaxYear";
                csvLine = csvHeader;
                sw.WriteLine(csvLine);
                csvLine = null;

            }
            string batch;
            switch (rpParameters.PaySchedule)
            {
                case "Monthly":
                    batch = "M";
                    break;
                case "Fortnightly":
                    batch = "F";
                    break;
                case "FourWeekly":
                    batch = "FW";
                    break;
                case "Quarterly":
                    batch = "Q";
                    break;
                case "Yearly":
                    batch = "Y";
                    break;
                default:
                    batch = "W";
                    break;
            }

            string process;
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
            //From payHistoryDetails[??] (LastAutoEnrolmentAssessmentDate) to payHistoryDetails[??] (Assessment Tax Year)
            for (int i = 49; i < 54; i++)
            {
                csvLine = csvLine + "\"" + payHistoryDetails[i] + "\"" + ",";
            }

            csvLine = csvLine.TrimEnd(',');

            sw.WriteLine(csvLine);

        }

        public static Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                List<RPPensionContribution>, RPEmployer, RPParameters>
            PreparePeriodReport(XDocument xdoc, XmlDocument xmlPeriodReport)
        {
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

        public static Tuple<List<RPEmployeePeriod>, List<RPPayComponent>, List<P45>, List<RPPreSamplePayCode>,
                           List<RPPensionContribution>, RPEmployer, RPParameters> 
                           PreparePeriodReport(XDocument xdoc, FileInfo file)
        {
            XmlDocument xmlPeriodReport = new XmlDocument();
            xmlPeriodReport.Load(file.FullName);

            return PreparePeriodReport(xdoc, xmlPeriodReport);
        }

        private RPP32Report CreateP32Report(XDocument xdoc, RPEmployer rpEmployer, RPParameters rpParameters)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-P32xml" + "\\";

            RPP32Report rpP32Report;
            PayRunIOWebGlobeClass prWG = new PayRunIOWebGlobeClass();

            XmlDocument p32ReportXml = new XmlDocument();

            

            bool test = false;
            if(test)
            {
                p32ReportXml.Load("C:\\Payescape\\Data\\Save\\P32.xml");
            }
            else
            {
                p32ReportXml = prWG.GetP32Report(xdoc, rpParameters);
            }
            RegexUtilities regexUtilities = new RegexUtilities();
            p32ReportXml.Save(outgoingFolder + regexUtilities.RemoveNonAlphaNumericChars(rpEmployer.Name) + "-P32.xml");
            rpP32Report = PrepareP32SummaryReport(xdoc, p32ReportXml, rpParameters, prWG);

            return rpP32Report;
        }
        
        public static RPP32Report PrepareP32SummaryReport(XDocument xdoc, XmlDocument p32ReportXml, RPParameters rpParameters, PayRunIOWebGlobeClass prWG)
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
                rpP32Report.ApprenticeshipLevyAllowance = prWG.GetDecimalElementByTagFromXml(header, "ApprenticeshipLevyAllowance");
                rpP32Report.AnnualEmploymentAllowance = prWG.GetDecimalElementByTagFromXml(header, "AnnualEmploymentAllowance");
                rpP32Report.OpeningBalancesRequired = true;
            }
            bool addToList = false;
            bool annualTotalRequired = false;
            RPP32ReportMonth obP32ReportMonth = null;   //To hold the opening balance object until I can check for month 1
            List<RPP32ReportMonth> rpP32ReportMonths = new List<RPP32ReportMonth>();
            foreach(XmlElement reportMonth in p32ReportXml.GetElementsByTagName("ReportMonth"))
            {
                RPP32ReportMonth rpP32ReportMonth = new RPP32ReportMonth();
                rpP32ReportMonth.PeriodNo = Convert.ToInt32(reportMonth.GetAttribute("Period"));
                rpP32ReportMonth.RPPeriodNo = rpP32ReportMonth.PeriodNo.ToString();
                rpP32ReportMonth.RPPeriodText = "Month " + rpP32ReportMonth.PeriodNo.ToString();
                if(rpP32ReportMonth.PeriodNo == 0)
                {
                    rpP32ReportMonth.RPPeriodNo = " ";
                    rpP32ReportMonth.RPPeriodText = "Previous Months";
                }
                rpP32ReportMonth.PeriodName = reportMonth.GetAttribute("RootNodeName");

                RPP32Breakdown rpP32Breakdown = new RPP32Breakdown();
                List<RPP32Schedule> rpP32Schedules = new List<RPP32Schedule>();

                foreach (XmlElement paySchedule in reportMonth.GetElementsByTagName("PaySchedule"))
                {
                    addToList = false;

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
                        rpP32PayRun.StudentLoan += rpP32PayRun.PostGraduateLoan;
                        rpP32PayRun.NetIncomeTax = prWG.GetDecimalElementByTagFromXml(payRun, "NetIncomeTax");
                        rpP32PayRun.GrossNICs = prWG.GetDecimalElementByTagFromXml(payRun, "GrossNICs");
                        rpP32PayRun.PayPeriod = rpP32PayRun.PayDate.Day < 6 ? rpP32PayRun.PayDate.Month - 4 : rpP32PayRun.PayDate.Month - 3;
                        if (rpP32PayRun.PayPeriod <= 0)
                        {
                            rpP32PayRun.PayPeriod += 12;
                        }
                        rpP32PayRun.TaxPeriod = prWG.GetIntElementByTagFromXml(payRun, "TaxPeriod");
                        rpP32PayRuns.Add(rpP32PayRun);
                    }
                    if(rpP32PayRuns.Count > 0)
                    {
                        rpP32Schedule.RPP32PayRuns = rpP32PayRuns;

                        addToList = true;
                    }
                    rpP32Schedules.Add(rpP32Schedule);
                    
                }
                rpP32Breakdown.RPP32Schedules = rpP32Schedules;

                rpP32ReportMonth.RPP32Breakdown = rpP32Breakdown;
                
                RPP32Summary rpP32Summary = new RPP32Summary();

                foreach(XmlElement summary in reportMonth.GetElementsByTagName("Summary"))
                {
                    rpP32Summary.Tax = prWG.GetDecimalElementByTagFromXml(summary, "Tax");
                    rpP32Summary.StudentLoan= prWG.GetDecimalElementByTagFromXml(summary, "StudentLoan");
                    rpP32Summary.PostGraduateLoan = prWG.GetDecimalElementByTagFromXml(summary, "PostGraduateLoan");
                    rpP32Summary.StudentLoan += rpP32Summary.PostGraduateLoan;
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
                    rpP32Summary.SpbpRecovered = prWG.GetDecimalElementByTagFromXml(summary, "SpbpRecovered");
                    rpP32Summary.SpbpComp = prWG.GetDecimalElementByTagFromXml(summary, "SpbpComp");
                    rpP32Summary.CisDeducted = prWG.GetDecimalElementByTagFromXml(summary, "CisDeducted");
                    rpP32Summary.CisSuffered = prWG.GetDecimalElementByTagFromXml(summary, "CisSuffered");
                    rpP32Summary.EmploymentAllowance = prWG.GetDecimalElementByTagFromXml(summary, "EmploymentAllowance");
                    rpP32Summary.NetNICs = prWG.GetDecimalElementByTagFromXml(summary, "NetNICs") - rpP32Summary.EmploymentAllowance;
                    rpP32Summary.AppLevy = prWG.GetDecimalElementByTagFromXml(summary, "ApprenticeshipLevy");
                    rpP32Summary.AmountDue = prWG.GetDecimalElementByTagFromXml(summary, "AmountDue");
                    rpP32Summary.AmountPaid = prWG.GetDecimalElementByTagFromXml(summary, "AmountPaid");
                    rpP32Summary.RemainingBalance = prWG.GetDecimalElementByTagFromXml(summary, "RemainingBalance");
                    rpP32Summary.TotalDeductions = rpP32Summary.EmploymentAllowance + rpP32Summary.SmpComp + rpP32Summary.SmpRecovered + rpP32Summary.SppComp +
                                                   rpP32Summary.SppRecovered + rpP32Summary.SapComp + rpP32Summary.SapRecovered + rpP32Summary.ShppComp +
                                                   rpP32Summary.ShppRecovered;

                }

                rpP32ReportMonth.RPP32Summary = rpP32Summary;

                if (!addToList)
                {
                    //If any of the values are not zero add the P32 period to the list
                    addToList = CheckIfNotZero(rpP32ReportMonth);

                }
                ////We only want to print the opening balances is values are not all zero or the values for period 1 are all zeros
                ////Check if PeriodNo is zero then store until we know if there's period 1
                if(addToList)
                {
                    //If it's got values always add it.
                    rpP32ReportMonths.Add(rpP32ReportMonth);
                    annualTotalRequired = true;
                }
                else
                {
                    if(rpP32ReportMonth.PeriodNo == 0)
                    {
                        //If it's period 0 (opening balances) store it we might use it even if it is all zeros
                        obP32ReportMonth = rpP32ReportMonth;
                    }
                    else if(rpP32ReportMonth.PeriodNo == 1)
                    {
                        if(obP32ReportMonth != null)
                        {
                            //If it's period 1 add period 0 was also all zeros add in period 0
                            rpP32ReportMonths.Add(obP32ReportMonth);
                            annualTotalRequired = true;
                        }
                        
                    }
                }
                //if (rpP32ReportMonth.PeriodNo == 0)
                //{
                //    if(addToList)
                //    {
                //        //Add period zero as it does have values
                //        rpP32ReportMonths.Add(rpP32ReportMonth);
                //        annualTotalRequired = true;
                //    }
                //    else
                //    {
                //        //Store opening balance line for now as it is all zeros.
                //        obP32ReportMonth = rpP32ReportMonth;
                //    }

                //}
                //else if(rpP32ReportMonth.PeriodNo == 1)
                //{
                //    if(!addToList)
                //    {
                //        if(obP32ReportMonth != null)
                //        {
                //            //Period 1 is not being added the list so add the period 0, opening balances whether it's all zeros or not.
                //            rpP32ReportMonths.Add(obP32ReportMonth);
                //            annualTotalRequired = true;
                //        }

                //    }
                //    else
                //    {
                //        rpP32ReportMonths.Add(obP32ReportMonth);
                //        annualTotalRequired = true;
                //    }
                //}
                //else
                //{
                //    if(addToList)
                //    {
                //        rpP32ReportMonths.Add(rpP32ReportMonth);
                //        annualTotalRequired = true;
                //    }
                //}
                ////Check if PeriodNo is zero.
                //if (rpP32ReportMonth.PeriodNo == 0)
                //{
                //    addToList = true;
                //}
                //if (addToList)
                //{
                //    //They want to show an opening balance line, even if it's all zeros, where a payroll started midway through the year.
                //    //However they do not want to show an opening balance line if the payroll started in month 1.
                //    //Problem is the report always returns a month 1, it'll just be all zeros if the payroll wasn't run for month 1.
                //    //So I'm going store the opening balance until I get a look at month 1.
                //    //If there's a non zero month 1 don't add the opening balances, if there's no non zero month 1 and opening balances even if they are zero.
                //    if (rpP32ReportMonth.PeriodNo == 0)
                //    {
                //        //Opening balance line store it for now.
                //        obP32ReportMonth = rpP32ReportMonth;
                //    }
                //    else if(rpP32ReportMonth.PeriodNo == 1)
                //    {
                //        //Period 1 is not all zeros so we don't want an opening balance line
                //        obP32ReportMonth = null;
                //        rpP32ReportMonths.Add(rpP32ReportMonth);
                //        annualTotalRequired = true;
                //    }
                //    else
                //    {
                //        //If we didn't get a period 1 then add in the opening balance.
                //        if(obP32ReportMonth != null)
                //        {
                //            rpP32ReportMonths.Add(obP32ReportMonth);
                //            obP32ReportMonth = null;
                //        }
                //        rpP32ReportMonths.Add(rpP32ReportMonth);
                //        annualTotalRequired = true;
                //    }

                //}

            }
            rpP32Report.RPP32ReportMonths = rpP32ReportMonths;

            if (annualTotalRequired)
            {
                RPP32ReportMonth rpP32ReportMonth = new RPP32ReportMonth();
                rpP32ReportMonth.PeriodNo = 13;
                rpP32ReportMonth.RPPeriodNo = "";
                rpP32ReportMonth.RPPeriodText = "Year " + rpP32Report.TaxYear.ToString() + "/" + (rpP32Report.TaxYear + 1).ToString();
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
                    rpP32Summary.StudentLoan += rpP32Summary.PostGraduateLoan;
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
                    rpP32Summary.EmploymentAllowance = prWG.GetDecimalElementByTagFromXml(annualTotal, "EmploymentAllowance");
                    rpP32Summary.NetNICs = prWG.GetDecimalElementByTagFromXml(annualTotal, "NetNICs") - rpP32Summary.EmploymentAllowance;
                    rpP32Summary.AppLevy = prWG.GetDecimalElementByTagFromXml(annualTotal, "ApprenticeshipLevy");
                    rpP32Summary.AmountDue = prWG.GetDecimalElementByTagFromXml(annualTotal, "AmountDue");
                    rpP32Summary.AmountPaid = prWG.GetDecimalElementByTagFromXml(annualTotal, "AmountPaid");
                    rpP32Summary.RemainingBalance = prWG.GetDecimalElementByTagFromXml(annualTotal, "RemainingBalance");
                    rpP32Summary.TotalDeductions = rpP32Summary.EmploymentAllowance + rpP32Summary.SmpComp + rpP32Summary.SmpRecovered + rpP32Summary.SppComp +
                                                   rpP32Summary.SppRecovered + rpP32Summary.SapComp + rpP32Summary.SapRecovered + rpP32Summary.ShppComp +
                                                   rpP32Summary.ShppRecovered;
                }
                rpP32ReportMonth.RPP32Summary = rpP32Summary;

                rpP32Report.RPP32ReportMonths.Add(rpP32ReportMonth);
            }

            
            return rpP32Report;
        }
        private static bool CheckIfNotZero(RPP32ReportMonth rpP32ReportMonth)
        {
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
        
        private void CreatePreSampleXLSX(XDocument xdoc, List<RPEmployeePeriod> rpEmployeePeriodList,
                                       RPParameters rpParameters, List<RPPreSamplePayCode> rpPreSamplePayCodes)
        {
            string outgoingFolder = xdoc.Root.Element("DataHomeFolder").Value + "PE-Reports";
            string coNo = rpParameters.ErRef;
            //Create a list of the required fixed columns.
            List<string> fixCol = CreateListOfFixedColumns();

            //Create a list of the required variable columns.
            List<string> varCol = CreateListOfVariableColumns(rpPreSamplePayCodes);

            //Create a workbook.
            string workBookName = outgoingFolder + "\\" + coNo + "\\Pre.xlsx";
            Workbook workbook = new Workbook(workBookName, "Pre");
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
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.AEO);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.StudentLoan);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.NetPayTP);
            workbook.CurrentWorksheet.AddNextCell(rpEmployeePeriod.ErNICTP);

            decimal erPensionTP = 0;
            foreach(RPPensionPeriod pensionPeriod in rpEmployeePeriod.Pensions)
            {
                erPensionTP += pensionPeriod.ErPensionTaxPeriod;
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
                        Protocol = WinSCP.Protocol.Sftp,
                        HostName = strHostName,    //"trans.bluemarblepayroll.com",
                        UserName = strUserName,    //"payescapetest",
                        Password = null,
                        PortNumber = 22,
                        SshHostKeyFingerprint = "ssh-rsa 2048 22:5f:d5:de:80:1d:52:69:72:55:3d:38:17:53:24:aa", //Old server  SshHostKeyFingerprint = "ssh-rsa 2048 f9:9e:38:ae:8d:55:d6:5d:f2:b3:63:67:e1:e4:d1:e1",
                        SshPrivateKeyPath = strSSHPrivateKeyPath    //"X:/jim/Documents/Payescape/Contracts/SFTP Private Key File/payescape.ppk"
                    };
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
        public void EmptyDirectory(DirectoryInfo directory)
        {
            foreach (FileInfo file in directory.GetFiles()) file.Delete();
            foreach (DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }
    }
    
}
