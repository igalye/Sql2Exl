using ConsoleLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sql2ExlNoGui
{
    class Program
    {
        static internal CmdLineMain runMe;
        static List<string> AvailableParams = new List<string> {"Sys","Connection", "ConnectionEncrypted", "ScriptName","OutputFile","UseTimeStamp","TimeOut",
            "OutputFolder","UseTitleFromScript","Extension","Debug","SQLConnectionForXl","SuppressType","Encrypt","Decrypt","Paraphrase",
            "env","var_name","TryMatchingParamDefinition" };//,"AppendToFile"};

        private static void CmdParamsInitialize()
        {
            StringBuilder sb = new StringBuilder();            
            
            sb.AppendLine("Extension = xls/xlsx/csv");
            sb.AppendLine("SuppressType = Always/IfEmpty/Never");
            sb.AppendLine("Sys = ms/oracle");

            runMe.HelpAvailableParams = sb.ToString();
            Helper.AvailableParams = AvailableParams;
            runMe.sParaphrase = "ExcelNoGui";
        }

        static int Main(string[] args)
        {            
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(threadExc);            
            
            string sVersion;            

            try
            {
                string sReleasePath = Properties.Settings.Default.ReleasePath;
                bool bCheckUpdate = Properties.Settings.Default.CheckForUpdate;
                IgalControls.AutoUpdate autoUpdate = new IgalControls.AutoUpdate(sReleasePath);

                Console.WriteLine(autoUpdate.ShowProgramFilesVersions());

                if (bCheckUpdate && !autoUpdate.VersionCheck(out sVersion))
                {
                    Console.WriteLine("updating program");
                    autoUpdate.UpdateProgram();
                    Environment.ExitCode = (int)ExitCode.UpdatingProg;
                    return (int)ExitCode.UpdatingProg;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in checking/updating program version:\n{ex.ToString()}");
            }            

            int iExported = 0;
            Sql2XlParams sql2Xl;

            if (args.Length == 0)
            {
                Environment.ExitCode = (int)ExitCode.Fail;
                return (int)ExitCode.Fail;
            }

            runMe = new CmdLineMain();
            CmdParamsInitialize();

            if (args.Length == 1)
            {
                if ((args[0].ToLower() == "help" || args[0].ToLower() == "/help"))
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("encrypt=password - encrypts to 64-base based on internal paraphrase");
                    sb.AppendLine("decrypt=64-EncryptedPassword paraphrase=givenparaphrase - decrypts a given 64-base string upon a paraphrase");
                    sb.AppendLine();
                    sb.AppendLine("For oracle paramters available option TryMatchingParamDefinition=true (or 1) - tries to set date value as OracleDate, and number as OracleNumber");
                    sb.AppendLine("Others - varchar2. by default (recommended) - false, i.e. every parameter will stay varchar2");                    
                    Console.WriteLine(runMe.Help(sb.ToString()));
                    return (int)ExitCode.Success;
                }

                if (args[0].ToLower().Contains("encrypt="))
                {
                    runMe.CreateParamPairs(args, new List<string>{"encrypt"});

                    string sPwd = runMe.ParamsForProgramPairs[0].SqlParams.FieldValue;
                    Console.WriteLine(runMe.Encrypt(sPwd));
                    return (int)ExitCode.Success;
                }
            }

            if (args.Any(i=>i.ToLower().Contains("decrypt=")))
            {
                if (!args.Any(i => i.ToLower().Contains("paraphrase=")))
                {
                    Console.WriteLine("Paraphrase needed to decrypt this string. You have to know it");
                    return (int)ExitCode.Fail;
                }

                runMe.CreateParamPairs(args, new List<string> { "decrypt","paraphrase"});

                try
                {
                    string sPwd = runMe.ParamsForProgramPairs.First(col => col.SqlParams.FieldName.ToLower() == "decrypt").SqlParams.FieldValue;
                    string sParaphrase = runMe.ParamsForProgramPairs.First(col => col.SqlParams.FieldName.ToLower() == "paraphrase").SqlParams.FieldValue;
                    Console.WriteLine(runMe.Decrypt(sPwd, sParaphrase));
                    return (int)ExitCode.Success;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    Environment.ExitCode = (int)ExitCode.Fail;
                    return (int)ExitCode.Fail;
                }                
            }

            try
            {
                CmdLineMain.StartHere();                
                
                runMe.CreateParamPairs(args, AvailableParams);

                sql2Xl = new Sql2XlParams(runMe.ParamsForProgramPairs);
                sql2Xl.CallingProgramm = "@sql2Xl";

                //ParamsShow
                Console.WriteLine(string.Format("Run params:\n{0}",runMe.ToString()));

                //igal 22/7/19 
                if (sql2Xl.suppress != ConsoleLib.Sql2XlParams.SuppressType.Always)
                {                    
                    switch (sql2Xl.FileExportType)
                    {                    
                        case Export2ExlEngine.ExportType.Excel:
                            iExported = sql2Xl.CreateExcel();
                        break;
                        case Export2ExlEngine.ExportType.CSV:
                            iExported = sql2Xl.CreateCsv();
                        break;
                        default:                        
                            throw new Exception("No File Export type is chosen. Should be xls/xlsx/csv");
                    }
                } 
            }
            catch (Exception ex)
            {
                Environment.ExitCode = (int)ExitCode.Fail;
                StringBuilder sEx = new StringBuilder();
                sEx.AppendLine(ex.ToString());
                sEx.AppendLine();
                sEx.AppendLine("Passed arguments in command line:");
                sEx.AppendLine(runMe.ToString());
                Console.WriteLine(sEx.ToString());                
                return (int)ExitCode.Fail;                               
            }

            if(sql2Xl.OutputFromScript != null && sql2Xl.OutputFromScript.Trim().Length>0)
                Console.WriteLine("@sql2Xl print - " + sql2Xl.OutputFromScript);
            Console.WriteLine(sql2Xl.ExportFinalMessage);
            
            Environment.ExitCode = (int)ExitCode.Success;
            return (int)ExitCode.Success;  
        }

        static public void threadExc(object sender, UnhandledExceptionEventArgs ta)
        {
            string MyAppTitle = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Console.WriteLine(string.Format("{0}: An unhandled exception occured...", MyAppTitle));
            Console.WriteLine(ta.ExceptionObject.ToString());
            Console.WriteLine(string.Format("{0}: End of the unhandled exception. Abort with exit code 1.", MyAppTitle));
            Environment.ExitCode = (int)ExitCode.Fail;            
        }
    }
}
