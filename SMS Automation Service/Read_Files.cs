using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using SMS_Automation_Service_Block.Object;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SMS_Automation_Service
{
    public partial class ReadFilesServices : ServiceBase
    {
        private string ReadFrom;
        private string SaveTo;
        private string NumberOfBoot;
        private string processed;
        private int Interval;
        private string TypeOfInterval;
        private string TimeForRunningTheService;
        private bool ProcessTheService = false;
        private string ConfigFile = @"C:\Program Files\SquareOne Technologies\Ooredoo Load Balancer\Configuration.ini";

        Timer timer = new Timer();


        public ReadFilesServices()
        {
            InitializeComponent();
            eventLog1 = new System.Diagnostics.EventLog();


            try
            {
                if (!System.Diagnostics.EventLog.SourceExists("ReadFilesService"))
                {
                    System.Diagnostics.EventLog.CreateEventSource(
                        "ReadFilesService", "ReadFilesLog");
                }
                eventLog1.Source = "ReadFilesService";
                eventLog1.Log = "ReadFilesLog";
            }
            catch (System.Security.SecurityException ex)
            {
                Console.WriteLine($"SecurityException: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
   
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            eventLog1.WriteEntry("In OnTimer");

            StreamReader Casa_streamReader1;

            try
            {


                Casa_streamReader1 = File.OpenText(ConfigFile);

                StreamReader Casa_streamReader2 = Casa_streamReader1;

                this.ReadFrom = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("Read From:" + this.ReadFrom);

                this.SaveTo = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("Save To:" + this.SaveTo);

                this.processed = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("processed:" + this.processed);

                this.NumberOfBoot = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("Number Of Boot:" + this.NumberOfBoot);

                this.TypeOfInterval = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("Type of interval:" + this.TypeOfInterval);

                this.TimeForRunningTheService = Casa_streamReader2.ReadLine();
                eventLog1.WriteEntry("Value of interval : " + this.TimeForRunningTheService);

                if (!string.IsNullOrEmpty(this.TypeOfInterval))
                {
                    if (this.TypeOfInterval.Contains("At"))
                    {



                        DateTime dateTime = DateTime.Now;
                        string formattedDateTime = dateTime.ToString("h:mm tt");
                        DateTime CurrentDateTimeFormatted = DateTime.Parse(formattedDateTime);
                        DateTime TTRTSdateTime = DateTime.Parse(this.TimeForRunningTheService);
                        this.TimeForRunningTheService = TTRTSdateTime.ToString("h:mm tt");
                        DateTime TimeForRTSFormatted = DateTime.Parse(TimeForRunningTheService);




                        if (TimeForRTSFormatted.CompareTo(CurrentDateTimeFormatted) == 0)
                        {
                            this.Interval = 24 * 60 * 60 * 1000;
                            this.ProcessTheService = true;
                        }
                        else
                        {
                            this.Interval = 30 * 1000;
                            this.ProcessTheService = false;
                            timer.Interval = this.Interval;
                            timer.Enabled = true;

                        }
                    }
                    else
                    {
                        if (this.TypeOfInterval.Split(' ')[1] == "Minutes")
                        {
                            this.Interval = Convert.ToInt32(this.TimeForRunningTheService) * 60 * 1000;
                            this.ProcessTheService = true;
                        }
                        else if (this.TypeOfInterval.Split(' ')[1] == "Seconds")
                        {
                            this.Interval = Convert.ToInt32(this.TimeForRunningTheService) * 1000;
                            this.ProcessTheService = true;
                        }
                        else
                        {
                            this.Interval = 30 * 1000;

                        }
                    }
                }
                Casa_streamReader2.Close();

                this.timer.Enabled = true;
                this.timer.Interval = this.Interval;
            }
            catch (Exception ex1)
            {
                ProjectData.SetProjectError(ex1);
                Exception exception = ex1;
                try
                {
                    //streamReader1.Close();
                }
                catch (Exception ex2)
                {
                    ProjectData.SetProjectError(ex2);
                    ProjectData.ClearProjectError();
                }
                eventLog1.WriteEntry("Error occured while reading settings !\r\n" + exception.ToString());
                ProjectData.ClearProjectError();
            }
            finally
            {
                try
                {
                    //streamReader1.Close();
                }
                catch (Exception ex)
                {
                    ProjectData.SetProjectError(ex);
                    ProjectData.ClearProjectError();
                }
            }

            /*End of OnStart Method Implementation*/
            // TODO: Insert monitoring activities here.
            eventLog1.WriteEntry("Start Read Files", EventLogEntryType.Information, 0);
            if (this.ProcessTheService)
            {
                #region Code


                string FileToWrite = this.SaveTo;
                string folderPath = this.ReadFrom;
                String FileToprocessed = this.processed;
                int numberOfOutputFiles = int.Parse(this.NumberOfBoot);
                string[] files = Directory.GetFiles(folderPath);
                string filePath = "";
                List<string> documents = new List<string>();
                try
                {
                    ProcessFileContent(files);
                    //System.Threading.Thread.Sleep(1000);

                    for (int ii = 0; ii < files.Length; ii++)
                    {
                        filePath = files[ii];
                        string[] lines = File.ReadAllLines(filePath);
                        var docStartIndexes = lines.Select((line, index) => new { Line = line, Index = index })
                        .Where(item => item.Line.ToLower().StartsWith("docstart"))
                        .Select(item => item.Index)
                        .ToList();
                        for (int count = 0; count < docStartIndexes.Count; count++)
                        {
                            string documentText;
                            if (count == docStartIndexes.Count - 1)
                            {
                                documentText = string.Join("\n", lines.Skip(docStartIndexes[count]).Take(lines.Length - docStartIndexes[count]));
                                documents.Add(documentText);
                            }
                            else
                            {
                                documentText = string.Join("\n", lines.Skip(docStartIndexes[count]).Take(docStartIndexes[count + 1] - docStartIndexes[count]));
                                documents.Add(documentText);
                            }
                        }
                        try
                        {
                            MoveFile(filePath, FileToprocessed);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error: {ex.Message}");
                        }

                        //////
                        List<MyObject> myObjectsList2 = new List<MyObject>();
                        int t = 1;
                        foreach (string Doc in documents)
                        {
                            MyObject docobj = new MyObject { MyInteger = t, MyString = Doc };
                            myObjectsList2.Add(docobj);
                            if (t == numberOfOutputFiles)
                            {
                                t = 1;
                            }
                            else
                            {
                                t++;
                            }
                        }


                        for (t = 1; t <= numberOfOutputFiles; t++)
                        {
                            string outputFileName = $@"{FileToWrite}\_{DateTime.Now:yyyyMMdd}_output_{t}.txt";
                            StreamWriter writer = null;
                            try
                            {
                                writer = new StreamWriter(outputFileName, true);
                                foreach (var obj in myObjectsList2)
                                {
                                    if (obj.MyInteger == t)
                                    {
                                        writer.WriteLine(obj.MyString);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"An error occurred while writing to file: {ex.Message}");
                            }
                            finally
                            {
                                writer?.Close();
                            }
                        }


                        myObjectsList2.Clear();
                        documents.Clear();
                        //////

                        Console.WriteLine("Batches saved successfully.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }

            }
            private void ProcessFileContent(string[] filePaths)
            {
                foreach (string filePath in filePaths)
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    Encoding encoding1256 = Encoding.GetEncoding(1256);
                    string fileContent = encoding1256.GetString(fileBytes);
                    File.WriteAllText(filePath, fileContent);

                    //encoding1256 = null;
                    //fileBytes = null;
                    //fileContent = null;

                }
            }


            private void MoveFile(string sourceFilePath, string FileToprocessed)
            {
                string fileName = Path.GetFileName(sourceFilePath);
                string destinationFilePath = Path.Combine(FileToprocessed, fileName);
                string processedFileName = $"FileProcessed_{DateTime.Now:yyyyMMdd_HHmmss}_{fileName}";
                string processedFilePath = Path.Combine(FileToprocessed, processedFileName);
                if (File.Exists(sourceFilePath))
                {
                    if (File.Exists(destinationFilePath))
                    {
                        File.Delete(destinationFilePath);
                    }

                    File.Move(sourceFilePath, processedFilePath);
                    Console.WriteLine("File moved successfully.");
                }
                else
                {
                    Console.WriteLine("Source file does not exist.");
                }
            }




            #endregion
                
        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("In OnStart");
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Set up a timer that triggers every minute.
            
            timer.Interval = 3000.0; // 60 seconds
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            // Update the service state to Stop Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update the service state to Stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
            eventLog1.WriteEntry("In OnStop.");
        }
        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        private string Decrypt(string strText)
        {
            string str = "W@A@S@I@M";
            byte[] rgbIV = new byte[8]
            {
        (byte) 18,
        (byte) 52,
        (byte) 86,
        (byte) 120,
        (byte) 144,
        (byte) 171,
        (byte) 205,
        (byte) 239
            };
            byte[] numArray = new byte[checked(strText.Length + 1)];
            string message;
            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes(Strings.Mid(str, 1, 8));
                DESCryptoServiceProvider cryptoServiceProvider = new DESCryptoServiceProvider();
                byte[] buffer = Convert.FromBase64String(strText);
                MemoryStream memoryStream = new MemoryStream();
                CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, cryptoServiceProvider.CreateDecryptor(bytes, rgbIV), CryptoStreamMode.Write);
                cryptoStream.Write(buffer, 0, buffer.Length);
                cryptoStream.FlushFinalBlock();
                message = Encoding.UTF8.GetString(memoryStream.ToArray());
            }
            catch (Exception ex)
            {
                
                message = ex.Message;
                eventLog1.WriteEntry(ex.Message+" || "+ex.StackTrace + " || "+ex.InnerException);
            }
            return message;
        }

        private string Encrypt(string strText)
        {
            string str1 = "W@A@S@I@M";
            byte[] rgbIV = new byte[8]
            {
        (byte) 18,
        (byte) 52,
        (byte) 86,
        (byte) 120,
        (byte) 144,
        (byte) 171,
        (byte) 205,
        (byte) 239
            };
            string str2;
            try
            {
                byte[] bytes1 = Encoding.UTF8.GetBytes(Strings.Mid(str1, 1, 8));
                byte[] bytes2 = Encoding.UTF8.GetBytes(strText);
                DESCryptoServiceProvider cryptoServiceProvider = new DESCryptoServiceProvider();
                MemoryStream memoryStream = new MemoryStream();
                CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, cryptoServiceProvider.CreateEncryptor(bytes1, rgbIV), CryptoStreamMode.Write);
                cryptoStream.Write(bytes2, 0, bytes2.Length);
                cryptoStream.FlushFinalBlock();
                str2 = Convert.ToBase64String(memoryStream.ToArray());
            }
            catch (Exception ex)
            {
               
                str2 = ex.Message;
                eventLog1.WriteEntry(ex.Message + " || " + ex.StackTrace + " || " + ex.InnerException);
            }
            return str2;
        }

        private string GetVar(string VarText)
        {
            VarText = this.Decrypt(VarText);
            return VarText.Split('=')[1];
        }

        static void SaveBatchToFile(List<string> batch, string fileName)
        {
            // Write the batch items to a file
            using (StreamWriter writer = new StreamWriter(fileName))
            {
                foreach (string item in batch)
                {
                    writer.WriteLine(item);
                }
            }
        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}
