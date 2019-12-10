using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sql2ExlNoGui
{
    class Program
    {
        static List<CmdLineParams> MyParams = new List<CmdLineParams>();
        enum ExitCode : int
        {
            Success = 0,
            Fail = 1
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        
        static int Main(string[] args)
        {
            string sCurrDir = System.IO.Path.GetDirectoryName((System.Reflection.Assembly.GetExecutingAssembly().Location));
            int iExported = 0;

            System.IO.Directory.SetCurrentDirectory(sCurrDir);

            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(threadExc);            

            List<string> MyParamsTmp = new List<string>();            
            CmdLineParams param;
            string sShowInConsoleParams = "";
            Sql2XlParams sql2Xl;

            if (args.Length == 1)
            {
                if((args[0].ToLower() == "help"))
                {                    
                    Console.WriteLine(CmdLineParams.Help());
                    return (int)ExitCode.Success; 
                }
                if(args[0].ToLower().Contains("encrypt"))
                {
                    string sPwd = args[0].Substring("encrypt=".Length, args[0].Length - "encrypt=".Length);
                    Console.WriteLine(CmdLineParams.Encrypt(sPwd));
                    return (int)ExitCode.Success; 
                }                
            }

            try
            {
                if (args.Length == 0)
                {
                    Environment.ExitCode = (int)ExitCode.Fail;
                    return (int)ExitCode.Fail;                         
                }
                
                for (int i = 0; i < args.Length; i++)
                {
                    bool bIsSqlParam = !args[i].Contains(":=");
                    string [] sSeparator;
                    if (bIsSqlParam)	                
		                sSeparator = new string [] {"="};
                    else
                        sSeparator = new string[] {":="};
                     
                    MyParamsTmp = args[i].Split(sSeparator,StringSplitOptions.None).ToList();
                    param = new CmdLineParams(MyParamsTmp, !bIsSqlParam);
                    if (param.ProgramParams.FieldName.ToLower().Contains("connection") && param.ProgramParams.FieldValue.ToLower().Contains("password"))
                    {
                        try
                        {
                            sShowInConsoleParams = IgalControls.IgalTools.HidePasswordInConnectionString(param.ProgramParams.FieldValue);
                            args[i] = sShowInConsoleParams;
                        }
                        catch (Exception)
                        {
                         //do nothing
                        }
                    }
                    MyParams.Add(param);
                }

                sql2Xl = new Sql2XlParams(MyParams);

                string sParamsShow = string.Join(Environment.NewLine, args);
                
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sCurrDir);                
                StringBuilder versions4Screen = new StringBuilder();

                string sMyAppName = System.Reflection.Assembly.GetCallingAssembly().ManifestModule.FullyQualifiedName;
                versions4Screen.AppendLine(GetFileVersion(sMyAppName));

                foreach (System.IO.FileSystemInfo fi in di.GetFileSystemInfos("*.dll",System.IO.SearchOption.TopDirectoryOnly))
                {
                    versions4Screen.AppendLine(GetFileVersion(fi.FullName));
                }
                
                Console.WriteLine(versions4Screen.ToString());
                Console.WriteLine(string.Format("Run params:\n{0}",sParamsShow));

                //igal 22/7/19 
                if (sql2Xl.suppress != Sql2XlParams.SuppressType.Always)
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
                sEx.AppendLine(ex.Message);
                sEx.AppendLine(ex.StackTrace);
                sEx.AppendLine();
                sEx.AppendLine("Passed arguments in command line:");
                sEx.AppendLine(string.Join(" ", args));
                Console.WriteLine(sEx.ToString());                
                return (int)ExitCode.Fail;                               
            }

            if(sql2Xl.OutputFromScript != null && sql2Xl.OutputFromScript.Trim().Length>0)
                Console.WriteLine("@sql2Xl print - " + sql2Xl.OutputFromScript);
            Console.WriteLine(sql2Xl.ExportFinalMessage);
            
            Environment.ExitCode = (int)ExitCode.Success;
            return (int)ExitCode.Success;  
        }

        private static string GetFileVersion(string file)
        {            
            return string.Format("File {0} version {1}", System.IO.Path.GetFileName(file), System.Reflection.Assembly.LoadFrom(file).GetName().Version.ToString());
        }

        static public void threadExc(object sender, UnhandledExceptionEventArgs ta)
        {
            Console.WriteLine("Sql2ExlNoGui: An unhandled exception occured...");
            Console.WriteLine(ta.ExceptionObject.ToString());
            Console.WriteLine("Sql2ExlNoGui: End of the unhandled exception. Abort with exit code 1.");
            Environment.ExitCode = (int)ExitCode.Fail;
            //return (int)ExitCode.Fail;
        }

        //static void exitProgram(ExitCode runResult)
        //{

        //}
    }
}
