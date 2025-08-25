using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using log4net;
using log4net.Config;
using Microsoft.Win32;
using PdfiumViewer;
using static System.Net.WebRequestMethods;

namespace PDFPrinterService
{
    partial class PrinterService : ServiceBase
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(PrinterService));
        private FileSystemWatcher watcher;

        // P/Invoke for WTSQuerySessionInformation
        [DllImport("wtsapi32.dll", CharSet = CharSet.Auto)]
        public static extern bool WTSQuerySessionInformation(IntPtr hServer, uint sessionId, WTSInfoClass wtsInfoClass, out IntPtr ppBuffer, out uint pBytesReturned);

        [DllImport("wtsapi32.dll", CharSet = CharSet.Auto)]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        public enum WTSInfoClass
        {
            WTSUserName = 5,
            WTSDomainName = 7,
        }


        public PrinterService()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                //string folderPath = @"C:\PrintQueue";
                //if (!Directory.Exists(folderPath))
                //{
                //    Directory.CreateDirectory(folderPath);
                //    log.Info("Directory created successfully.");
                //}

                // Commented below code because it returns path as "C:\Windows\system32\config\systemprofile\Downloads"
                string folderPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");
                log.Info("FolderPath from UserProfile: " + folderPath);

                ScanAvailablePrinters();

                try
                {
                    if (folderPath.ToLower().Contains("windows"))
                    {
                        folderPath = GetLoggedInUserDownloadsFolder();
                        log.Info("Logged_In_User_Downloads_Folder_Path: " + folderPath);
                    }
                }
                catch (Exception e) 
                {
                    folderPath = GetLoggedInUserDownloadsFolder();
                    log.Info("Logged_In_User_Downloads_Folder_Path: " + folderPath);
                }                
                
                if (folderPath != null)
                {
                    if (folderPath.Trim() != "")
                    {
                        watcher = new FileSystemWatcher(folderPath)
                        {
                            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                            Filter = "PackingSlip*.pdf"
                        };
                        watcher.Created += OnFileCreated;
                        watcher.Renamed += OnFileRenamed;
                        //watcher.Changed += OnFileChanged;
                        watcher.EnableRaisingEvents = true;
                    } 
                    else
                    {
                        log.Info("folderPath is empty.");
                    }
                }
                else
                {
                    log.Info("folderPath is null.");
                }
            }
            catch (Exception e)
            {
                log.Error("Error occurred in OnStart. Exception: " + e.Message);
            }            
        }

        public static string GetLoggedInUserDownloadsFolder()
        {
            try
            {
                string username = GetLoggedInUser();
                if (!string.IsNullOrEmpty(username))
                {
                    // Get the user's profile path
                    string userProfilePath = Path.Combine(@"C:\Users", username);

                    // Construct the path to the Downloads folder
                    string downloadsFolder = Path.Combine(userProfilePath, "Downloads");

                    return downloadsFolder;
                }
            }
            catch (Exception ex)
            {
                log.Error("Error occurred in GetLoggedInUserDownloadsFolder. Exception: " + ex.Message);
            }
            return null;
        }

        // Get the logged-in user using WTS API
        public static string GetLoggedInUser()
        {
            try
            {
                IntPtr buffer = IntPtr.Zero;
                uint bytesReturned = 0;

                if (WTSQuerySessionInformation(IntPtr.Zero, 1, WTSInfoClass.WTSUserName, out buffer, out bytesReturned))
                {
                    string userName = Marshal.PtrToStringAuto(buffer);
                    WTSFreeMemory(buffer);
                    return userName;
                }
            }
            catch (Exception e) 
            {
                log.Error("Error occurred in GetLoggedInUser. Exception: " + e.Message);
            }
            return null;
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {            
            ProcessFile("OnFileCreated", e);
        }

        private void OnFileRenamed(object sender, FileSystemEventArgs e)
        {            
            ProcessFile("OnFileRenamed", e);
        }

        private void ProcessFile(string calledFrom, FileSystemEventArgs e)
        {
            try
            {
                Thread.Sleep(1000);
                //String printMethod = ConfigurationManager.AppSettings["PrintMethod"];

                //if (printMethod.ToLower() == "acrobat")
                //{
                //    PrintPDFUsingAcrobatReader(e.FullPath);
                //}
                //else if (printMethod.ToLower() == "pdfium")
                //{
                PrintPDFUsingPdfium(e.FullPath, e.Name);
                //}
            }
            catch (Exception ex)
            {
                log.Error("Error occurred in " + calledFrom + " ProcessFile. Exception: " + ex.Message);
            }
        }

        //private void OnFileChanged(object sender, FileSystemEventArgs e)
        //{
        //    log.Info("In OnFileChanged. Path: " + e.FullPath);
        //}

        //private void PrintPDFUsingAcrobatReader(string filePath)
        //{
        //    try
        //    {                
        //        log.Info("============= BEGIN ==============");
        //        log.Info($"PrintPDFUsingAcrobatReader: Printing {filePath}");
        //        String printerName = ConfigurationManager.AppSettings["PrinterName"];
        //        String acrobatReaderPath = ConfigurationManager.AppSettings["AcrobatReaderSetupPath"];
        //        if (!File.Exists(acrobatReaderPath))
        //        {
        //            log.Error("PrintPDFUsingAcrobatReader: Adobe Acrobat Reader is not installed.");
        //        }
        //        String arguments = $"/t \"{filePath}\" \"{printerName}\"";

        //        Process process = new Process();
        //        process.StartInfo.FileName = acrobatReaderPath;
        //        process.StartInfo.Arguments = arguments;
        //        process.StartInfo.UseShellExecute = true;
        //        process.StartInfo.CreateNoWindow = true;
        //        process.Start();

        //        process.WaitForExit();

        //        log.Info("PrintPDFUsingAcrobatReader: Deleting the PDF file.");
        //        if (File.Exists(filePath))
        //        {
        //            File.Delete(filePath);
        //            log.Info("PrintPDFUsingAcrobatReader: PDF file deleted.");
        //        }
        //        log.Info("============== END ===============");
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("Error occurred in PrintPDFUsingAcrobatReader. Exception: " + ex.Message);
        //    }
        //}

        private void PrintPDFUsingPdfium(string filePath, string fileName)
        {
            try
            {
                log.Info("============= BEGIN ==============");
                log.Info($"PrintPDFUsingPdfium: Printing {filePath}");
                string printerName = ConfigurationManager.AppSettings["PrinterName"];
                string numberOfCopies = ConfigurationManager.AppSettings["NumberOfCopies"];

                PrinterSettings printerSettings = new PrinterSettings()
                {
                    PrinterName = printerName,
                    Copies = short.Parse(numberOfCopies)
                };

                log.Info($"PrintPDFUsingPdfium: Paper Sizes...");
                foreach (PaperSize ps in printerSettings.PaperSizes)
                {
                    log.Info($"Name: {ps.PaperName}; Width: {ps.Width}; Height: {ps.Height}");
                }                
                log.Info($"PrintPDFUsingPdfium:=================");

                PageSettings pageSettings = new PageSettings(printerSettings)
                {
                    Margins = new Margins(50, 50, 50, 50)
                    //Margins = printerSettings.DefaultPageSettings.Margins
                };

                //pageSettings.PaperSize = new PaperSize("A4", 827, 1169);
                try
                {
                    PaperSize a4Size = printerSettings.PaperSizes.Cast<PaperSize>().FirstOrDefault(ps => ps.PaperName == "A4");
                    if (a4Size != null)
                    {
                        pageSettings.PaperSize = a4Size;
                    }
                    else
                    {
                        log.Info($"PrintPDFUsingPdfium: A4 paper size not found in printer settings.");
                    }
                }
                catch (Exception ex)
                {
                    log.Error("Error occurred while setting A4 paper size in PrintPDFUsingPdfium. Exception: " + ex.Message);
                }

                using (PdfDocument pdfDocument = PdfDocument.Load(filePath))
                {
                    using (PrintDocument printDocument = pdfDocument.CreatePrintDocument())
                    {
                        printDocument.PrinterSettings = printerSettings;
                        printDocument.DefaultPageSettings = pageSettings;
                        printDocument.PrintController = (PrintController)new StandardPrintController();
                        printDocument.Print();
                    }
                }

                //log.Info("PrintPDFUsingPdfium: Deleting the PDF file.");
                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //    log.Info("PrintPDFUsingPdfium: PDF file deleted.");
                //}
                log.Info("============== END ===============");
            }
            catch (Exception ex)
            {
                log.Error("Error occurred in PrintPDFUsingPdfium. Exception: " + ex.Message);
                SendEmail(ex.Message, fileName);
            }
        }

        private void SendEmail(string errorMessage, string fileName)
        {
            try
            {
                string body = string.Empty;

                using (StreamReader sr = new StreamReader(ConfigurationManager.AppSettings["PrintingErrorEmailTemplatePath"]))
                {
                    body = sr.ReadToEnd();
                }

                try
                {
                    body = body.Replace("websiteURLHREF", "https://www.fluidsecure.com");
                    body = body.Replace("webisteURL", "www.fluidsecure.com");
                    body = body.Replace("ImageSign", "<img src='https://www.fluidsecure.net/Content/Images/FluidSECURELogo.png' style='width:200px'/>");
                }
                catch (Exception ex)
                {
                    body = body.Replace("ImageSign", "");
                }

                body = body.Replace("fileName", "- " + fileName);
                body = body.Replace("ErrorMessage", errorMessage);

                SmtpClient mailClient = new SmtpClient(ConfigurationManager.AppSettings["smtpServer"]);
                mailClient.UseDefaultCredentials = false;
                mailClient.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["emailAccount"], ConfigurationManager.AppSettings["emailPassword"]);
                mailClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings["smtpPort"]);
                mailClient.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]);

                MailMessage messageSend = new MailMessage();
                messageSend.From = new MailAddress(ConfigurationManager.AppSettings["FromEmail"]);
                messageSend.Subject = ConfigurationManager.AppSettings["PrintingErrorEmailSubject"];
                messageSend.Body = body;
                messageSend.IsBodyHtml = true;

                string printingErrorEmailsTo = ConfigurationManager.AppSettings["PrintingErrorEmailsTo"];
                string[] emailArray = printingErrorEmailsTo.Split(';');

                foreach (var emailId in emailArray)
                {
                    if (string.IsNullOrEmpty(emailId.Trim()))
                    {
                        continue;
                    }

                    messageSend.To.Add(emailId.Trim());
                    try
                    {
                        mailClient.Send(messageSend);
                        log.Info("At SendEmail. Email sent to : " + emailId);
                        messageSend.To.Remove(new MailAddress(emailId.Trim()));

                    } catch (Exception ex)
                    {
                        log.Error("Error occurred while sending printing error email to EmailId : " + emailId + ". Exception is: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Error occurred in SendEmail. Exception: " + ex.Message);
            }
        }

        private void ScanAvailablePrinters()
        {
            List<string> availablePrinters = new List<string>();
            try
            {
                log.Info("========================================================");
                log.Info("ScanAvailablePrinters: Checking available printers");
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    log.Info("Installed Printer: " + printer);
                }
                log.Info("========================================================");
            }
            catch (Exception ex)
            {
                log.Error("Error occurred in ScanAvailablePrinters. Exception: " + ex.Message);
            }
        }

        protected override void OnStop()
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
            }
        }
    }
}
