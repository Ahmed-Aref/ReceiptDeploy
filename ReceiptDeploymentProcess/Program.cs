using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.ServiceProcess;

namespace ReceiptDeploymentProcess
{
    class Program
    {
        //Variables
        static string version = "3.81";
        static string DPIFile = "DPPrintReceipt_ARLv" + version.Replace(".","") + ".dll";       //The Name of Domino's DLL VB6 File
        static string AREFRECLIB = "AREFRECLIB.dll";                                            //The Name of Customized DLL C#

        static string ServerIP = "192.168.80.80";
        static string ServerDB = "Shared";
        static string ServerUsername = "";
        static string ServerPassword = "";

        /*########################################################################################################################
        The Start Function*/
        static void Main(string[] args)
        {
            //Welcome Message
            Console.ForegroundColor = ConsoleColor.White; Console.WriteLine("Receipt for Pulse v" + version + "\n");

            //The Engine Function
            fncEngine();

            //Goodby Message
            Console.ForegroundColor = ConsoleColor.White; Console.WriteLine("\n* Update finish, Press Enter to Exit.");
            Console.ReadKey();
        }


        /*########################################################################################################################
        Core Function*/
        static void fncEngine()
        {
            ProcessStartInfo Proc = new ProcessStartInfo();
            Process pStart = new Process();
            string YourCommand = "";

            Proc.UseShellExecute = true;
            Proc.WorkingDirectory = @"C:\Windows\System32";
            Proc.FileName = @"C:\Windows\System32\cmd.exe";


            //-----------------------------------------
            //STEP 1 : Stop WWW Service
            try
            {
                YourCommand = @" net stop W3SVC /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Stop W3SVC Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Stop W3SVC Server [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 2 : Stop MSSQLSERVER Service
            try
            {
                YourCommand = @" net stop MSSQLSERVER /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Stop MSSQLSERVER Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Stop MSSQLSERVER Server [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 3 : Start MSSQLSERVER Service
            try
            {
                YourCommand = @" net start MSSQLSERVER /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : start MSSQLSERVER Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Start MSSQLSERVER Server [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 4 : Kill Order Entry
            try
            {
                YourCommand = @" taskkill /f /im DpOrderEntry.exe";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Kill Order Entry.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Kill Order Entry [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 5 : Kill Dispatch
            try
            {
                YourCommand = @"taskkill /f /im DpDispatch.exe";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Kill Dispatch.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Kill Dispatch [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 6 : Kill W3WP Service
            try
            {
                YourCommand = @"taskkill /f /im w3wp.exe";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Kill W3WP.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Kill W3WP [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 7 : Update Store Info
            try
            {
                updateStoreInfo();
                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Update store Information from Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Update store Information from GOLO [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 8 : Unregist Old DLL
            try
            {
                YourCommand = @" Regsvr32 -u -s " + "\"" + @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile + "\"";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Un-Regist old DLL.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Un-Regist old DLL [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 9 : Rename Old File
            string DateFormat = DateTime.Now.Year.ToString("d2") + DateTime.Now.Month.ToString("d2") + DateTime.Now.Day.ToString("d2") + DateTime.Now.Hour.ToString("d2") + DateTime.Now.Minute.ToString("d2");

            string OldName = @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile;
            string NewName = @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile + "." + DateFormat;

            if (File.Exists(OldName))
            {
                try
                {
                    System.IO.File.Move(OldName, NewName);
                    Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Rename old DLL.");
                }
                catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Rename old DLL [" + EX.Message + "]"); }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine("* Warning : Receipt file not exist.");
            }


            //-----------------------------------------
            // STEP 11 : Copy Receipt DPI DLL
            try
            {
                File.Copy(System.Environment.CurrentDirectory + @"\" + DPIFile, @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile, true);

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Copy DPI DLL.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Copy DPI DLL [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 12 : Copy Receipt AREFRECLIB DLL
            try
            {
                File.Copy(System.Environment.CurrentDirectory + @"\" + AREFRECLIB, @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + AREFRECLIB, true);

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Copy AREFRECLIB DLL.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Copy AREFRECLIB DLL [" + EX.Message + "]"); }



            //-----------------------------------------
            // STEP 13 : Regist the DLL
            try
            {
                YourCommand = @" Regsvr32 -u -s " + "\"" + @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile + "\"";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                YourCommand = @" Regsvr32 -s " + "\"" + @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + DPIFile + "\"";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Regist new PULSE DLL.");

            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Regist new DLLs [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 14 : Regist AREFRECLIB
            try
            {
                YourCommand = @"C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe /codebase /tlb " + "\"" + @"C:\Program Files\Dominos\Pulse\ReceiptPrintDLL\" + AREFRECLIB + "\"";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Regist new Custom DLL.");

            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Regist new Custom DLL. [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 15 : Start Dominos MapMarker Server
            try
            {
                YourCommand = @" net start MMS /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Start Dominos MapMarker Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Start Dominos MapMarker Server [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 16 : Start Pulse Live Data Refresh Service
            try
            {
                YourCommand = @" net start PulseDeviceManager /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : start Dominos Pulse Device Manager Service.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : start Dominos Pulse Device Manager Service [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 17 : Start Pulse Server
            try
            {
                YourCommand = @" net start PulseServer /yes";
                Proc.Arguments = "/c " + YourCommand;
                pStart = Process.Start(Proc);
                pStart.WaitForExit();

                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Start Dominos Pulse Server.");
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Start Dominos Pulse Server [" + EX.Message + "]"); }


            //-----------------------------------------
            // STEP 18 : Start IIS Server
            try
            {
                using (ServiceController controller = new ServiceController("IISADMIN"))
                {
                    ServiceRestart(controller);
                    Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("* Succeed : Restart IIS Service.");
                }
            }
            catch (Exception EX) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("* Failed : Restart IIS Service [" + EX.Message + "]"); }
        }


        /*########################################################################################################################
        General Functions*/
        public static void updateStoreInfo()
        {
            //Get Store Number from Local
            string StoreNumber = SQLSelectLoccal("SELECT Location_Code From POS.dbo.Location_Codes").Rows[0][0].ToString();

            //Get Information from Remote Server
            DataTable DT = SQLRemoteQuery(@"SELECT Store_Name,Store_Lat,Store_Long FROM Shared.dbo.Store WHERE Store_Number = " + StoreNumber, ServerIP, ServerDB, ServerUsername, ServerPassword);

            //Update Store Information from Remote Server
            string Query = @"   UPDATE Receipt.dbo.rcpVariable SET Value = '' WHERE Name = 'Storename_AR'
                                UPDATE Receipt.dbo.rcpVariable SET Value = '" + DT.Rows[0]["Store_Name"].ToString() + @"' WHERE Name = 'Storename_EN'
                                UPDATE Receipt.dbo.rcpVariable SET Value = '" + DT.Rows[0]["Store_Lat"].ToString() + @"' WHERE Name = 'Store_Lat'
                                UPDATE Receipt.dbo.rcpVariable SET Value = '" + DT.Rows[0]["Store_Long"].ToString() + @"' WHERE Name = 'Store_Long'
                                EXEC Receipt.dbo.spUpdateStoreLocation";

            SQLUpdateLoccal(Query);
        }

        //Update SQL Query : [Windows Credential]
        public static void SQLUpdateLoccal(string Query)
        {
            using (SqlConnection connection = new SqlConnection("Server= localhost; Database= master;Integrated Security=SSPI;"))
            {
                SqlCommand command = new SqlCommand(Query, connection);
                connection.Open();
                DataTable DT = new DataTable();
                DT.Load(command.ExecuteReader());
                connection.Close();
            }
        }

        //Select SQL Query : [Windows Credential]
        public static DataTable SQLSelectLoccal(string Query)
        {
            DataTable DT;

            using (SqlConnection connection = new SqlConnection("Server= localhost; Database= master;Integrated Security=SSPI;"))
            {
                SqlCommand command = new SqlCommand(Query, connection);
                connection.Open();
                DT = new DataTable();
                DT.Load(command.ExecuteReader());
                connection.Close();
            }
            return DT;
        }

        //SQL Query : [SQL Credential]
        public static DataTable SQLRemoteQuery(string queryString, string IP, string DB, string Username, string Password)
        {
            //The Connection String to make a connection
            string connectionString = "Server=" + IP + ";Database=" + DB + ";User Id=" + Username + ";Password=" + Password + ";TrustServerCertificate=True";
            DataTable dt = new DataTable();

            // Send Query
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();

                dt.Load(command.ExecuteReader());
                connection.Close();
            }
            return dt;
        }

        //Restart a service
        public static void ServiceRestart(ServiceController controller)
        {
            List<ServiceController> dependencies = new List<ServiceController>();
            FillDependencyTreeLeaves(controller, dependencies);
            controller.Stop();
            controller.WaitForStatus(ServiceControllerStatus.Stopped);
            foreach (ServiceController dependency in dependencies)
            {
                try
                {
                    dependency.Start();
                    dependency.WaitForStatus(ServiceControllerStatus.Running);
                }
                catch { }
            }
        }
        public static void FillDependencyTreeLeaves(ServiceController controller, List<ServiceController> controllers)
        {
            bool dependencyAdded = false;
            foreach (ServiceController dependency in controller.DependentServices)
            {
                ServiceControllerStatus status = dependency.Status;
                // add only those that are actually running
                if (status != ServiceControllerStatus.Stopped && status != ServiceControllerStatus.StopPending)
                {
                    dependencyAdded = true;
                    FillDependencyTreeLeaves(dependency, controllers);
                }
            }
            // if no dependency has been added, the service is dependency tree's leaf
            if (!dependencyAdded && !controllers.Contains(controller))
            {
                controllers.Add(controller);
            }
        }
    }
}