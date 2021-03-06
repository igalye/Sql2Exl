using System;

using System.Collections;

using System.Collections.Generic;

using System.Linq;

using System.Text;

using IgalDAL;

using System.Data;

using System.Data.SqlClient;

using Oracle.DataAccess.Client;

using System.IO;

using Export2ExlEngine;

using IgalControls;

using System.Text.RegularExpressions;



namespace Sql2ExlNoGui

{    

    internal class CmdLineParams : IEnumerable

    {        

        internal static string  sParaphrase = "ExcelNoGui";

        internal clsSimpleField ProgramParams { get; set; }

        internal clsSimpleField SqlParams { get; set; }



        private static List<string> ParamNames = new List<string> {"Sys","Connection", "ConnectionEncrypted", "ScriptName","OutputFile","UseTimeStamp","env","var_name","TimeOut",

            "OutputFolder","UseTitleFromScript","Extension","Debug","SQLConnectionForXl","SuppressType" };//,"AppendToFile"};



        internal CmdLineParams()

        {}



        internal CmdLineParams(List<string> ParamPair, bool bDefinedParamsOnly = true)

        {

            if (ParamPair.Count != 2)

            {

                StringBuilder sError = new StringBuilder();

                sError.AppendLine("The paramers that are passed to a constructor should be a pair of name and value");

                sError.AppendLine("The fault value is:");

                sError.AppendLine(ParamPair.Aggregate((a,b) => a + " " + b));

                throw new Exception(sError.ToString());

            }

            else

            {

                if (bDefinedParamsOnly)

                {

                    IsParamNamesCorrect(ParamPair[0]);                        

                }

                else

                {

                    SqlParams = new clsSimpleField(ParamPair[0], ParamPair[1]);

                }

                ProgramParams = new clsSimpleField(ParamPair[0], ParamPair[1]);

            }

        }



        internal static bool IsBasicParamName(string ParamName)

        {

            return ParamNames.Any(p => ParamName.ToLower() == p.ToLower());

        }



        private bool IsParamNamesCorrect(string ParamName)

        {

            if (!IsBasicParamName(ParamName))

            {

                throw new Exception("The unknown parameter name - " + ParamName);

            }

            

            return true;

        }



        public IEnumerator GetEnumerator()

        {   return this.GetEnumerator();    }



        public override string ToString()

        {

            string s = "";

            if(ProgramParams != null)

                s = ProgramParams.FieldName + " = " + ProgramParams.FieldValue;

            return s;

        }



        internal static string Help()

        {

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Available parameters:");

            sb.AppendLine();

            sb.AppendLine(string.Join("\n", ParamNames));

            sb.AppendLine("Help");

            sb.AppendLine();

            sb.AppendLine("Assigning value to program's parameter  :=  (semicolon equal)");

            //sb.AppendLine("parameters env and var_name are related to a function call configuration_pck on ALIS_PCLAL, and requires both");

            sb.AppendLine("After the last parameter - all parameters belong to sql. Assigning parameter to sql = (equal)");

            sb.AppendLine("E.g.: Sql2ExlNoGui.exe sys:=ms scriptname:=" + (char)32 + "my script.sql" + (char)32 + "OutputFile:=path\\filename.extension Connection:=connectionstring param1=value1 param2=value2");

            sb.AppendLine("extension = xls/xlsx/csv");

            sb.AppendLine("SuppressType = Always/IfEmpty/Never");

            sb.AppendLine("or,");            

            sb.AppendLine();

            sb.AppendLine("Sql2ExlNoGui.exe help");

            sb.AppendLine("or,");

            sb.AppendLine("Sql2ExlNoGui.exe encrypt=passwordToEncrypt");



            sb.AppendLine(IgalDAL.PublicModule.Help());

            return sb.ToString();

        }



        internal static string Encrypt(string Password)

        {

            try

            {

                return Password.Encrypt(sParaphrase);

            }

            catch (Exception ex)

            {   throw ex;  }                

        }

    }



    internal class Sql2XlParams : IEnumerable

    {

        internal enum SqlSystem

	    {

            NoSystem,

	        MSSql,

            Oracle

	    };



        //igal 22/7/19 

        internal enum SuppressType

        {

            Never,

            IfEmpty,

            Always

        }



        //internal CmdLineParams RunProgram;

        internal SqlSystem sys = SqlSystem.NoSystem;

        internal SuppressType suppress = SuppressType.Never; //igal 22/7/19 

        internal string SqlScriptName { get; set; }

        internal string OutputFile { get; set; }

        string sOutputFolder = "";

        internal string OutputFolder

        {

            get { return sOutputFolder; }

            set

            {

                DirectoryInfo di = new DirectoryInfo(value);

                if (!di.Exists)

                    throw new Exception("Folder " + value + " doesn't exist");

                sOutputFolder = value;

            }

        }

        internal Export2ExlEngine.ExportType FileExportType { get; set; }

        private SqlParameter[] sqlParam = null;

        private OracleParameter[] oraParam = null;

        private string _ConnectionString = "", _SqlCon4Xl = "";

        private bool bUseTimeStamp = false, bUseTitleFromScript = false, bDebug = false;

        private bool bAppendToFile = false; //igal 19-6-19

        private List<string> xlSheets = new List<string>();

        private string sTimeStamp = "", sExtension = "", sFileName = "";

        //private string sVarName = "", sEnv = "";

        private int iTimeOut = -1;

        string sql = "";



        internal string OutputFromScript { get; set; }        



        public string ExportFinalMessage

        {

            get

            {

                StringBuilder sb = new StringBuilder();

                for (int i = 0; i < ds.Tables.Count; i++)

                {

                    sb.AppendLine(string.Format("@sql2Xl out lines table {0} - {1} records", ds.Tables[i].TableName,ds.Tables[i].Rows.Count.ToString()));

                }

                

                return sb.ToString();

            }

        }



        internal string ConnectionString 

        {

            get { return _ConnectionString; }

            set

            {                

                _ConnectionString = value;

            }

        }



        internal string ConnectionStringEncrypted

        {

            set

            {

                string tmp = value, sFinalConString = "";

                string pwdEncrypted = "", pwdDecrypted = "", pwdEncryptedExtract = "";



                string[] sCon = tmp.Split(';');

                if (sCon.Any(p => p.StartsWith("Password")))

                {

                    try

                    {

                        //encrypted password - decrypt it.

                        pwdEncrypted = sCon.First(p => p.StartsWith("Password"));

                        pwdEncryptedExtract = pwdEncrypted.Substring(pwdEncrypted.IndexOf("=") + 1);

                        pwdDecrypted = pwdEncryptedExtract.Decrypt(CmdLineParams.sParaphrase);

                        int index = Array.FindIndex<string>(sCon, p => p.Contains("Password"));

                        sCon[index] = "Password=" + pwdDecrypted;

                        sFinalConString = string.Join(";", sCon);

                    }

                    catch (Exception ex)

                    {

                        throw ex;

                    }



                }

                else

                    sFinalConString = value;

                _ConnectionString = sFinalConString;

            }

        }



        private DataSet ds = new DataSet();



        public IEnumerator GetEnumerator()

        {   return GetEnumerator();    }



        internal Sql2XlParams(List<CmdLineParams> _params)

        {

            int iParamCount = 0;



            foreach (CmdLineParams item in _params)

            {

                switch (item.ProgramParams.FieldName.ToLower())

                {

                    case "debug":

                        try

                        {

                            bDebug = IgalTools.StringToBool(item.ProgramParams.FieldValue);

                        }

                        catch (Exception ex)

                        { throw ex; }



                        break;

                    case "sys":

                        switch (item.ProgramParams.FieldValue.ToLower())

	                    {

                            case "oracle":

                                sys = SqlSystem.Oracle;

                                break;

                            case "ms":

                                sys = SqlSystem.MSSql;

                                break;

		                    default:

                                throw new Exception("Incorrect or no sql system value: " + item.ProgramParams.FieldValue);                            

	                    }

                        break;

                    case "connection":

                        ConnectionString = item.ProgramParams.FieldValue;

                        break;

                    case "connectionencrypted":

                        ConnectionStringEncrypted = item.ProgramParams.FieldValue;

                        break;

                    case "scriptname":

                        SqlScriptName = item.ProgramParams.FieldValue.Trim();

                        break;

                    case "outputfile":

                        OutputFile = item.ProgramParams.FieldValue.Trim();

                        FileInfo fi = new FileInfo(OutputFile);

                        FileExportType = GetExportTypeByExtension(fi.Extension.Replace(".",""));

                        if(FileExportType == ExportType.NoType)

                            throw new Exception("Incorrect or no output file extension. Possible options: xls/xlsx/csv");

                        

                        break;

                    case "env":

                        //sEnv = item.ProgramParams.FieldValue.Trim();

                        throw new Exception("parameter env is now not supported");

                        //break;

                    case "var_name":

                        //sVarName = item.ProgramParams.FieldValue.Trim();

                        throw new Exception("parameter var_name is now not supported");

                        //break;

                    case "timeout":

                        int _iTimeOut = -1;

                        int.TryParse(item.ProgramParams.FieldValue, out _iTimeOut);

                        if (_iTimeOut < 0)

                            throw new Exception("TimeOut parameter is incorrect. Should be number greater or equal to 0!");

                        else

                        {

                            iTimeOut = _iTimeOut;

                            break;

                        }

                    case "appendtofile":

                        try

                        {

                            bAppendToFile = IgalTools.StringToBool(item.ProgramParams.FieldValue);

                        }

                        catch (Exception)

                        {

                            throw new Exception("Incorrect AppendToFile value. Possible options: yes/no/true/false/0/1"); ;

                        }

                        

                        break;

                    case "usetimestamp":

                        try

                        {

                            bUseTimeStamp = IgalTools.StringToBool(item.ProgramParams.FieldValue);

                        }

                        catch (Exception)

                        {

                            //check datetime stamp is correct - only ".", "-", "_" or numbers are allowed                                                                

                            if (!Regex.IsMatch(item.ProgramParams.FieldValue, @"^[0-9._-]*$"))

                                throw new Exception("Incorrect UseTimeStamp value. Possible options: yes/no/true/false/0/1");

                            else

                            {

                                sTimeStamp = item.ProgramParams.FieldValue.Trim();                                

                            }                            

                        }



                        break;

                    case "outputfolder":

                        OutputFolder = item.ProgramParams.FieldValue;

                        break;

                    case "extension":

                        sExtension = item.ProgramParams.FieldValue.Trim();

                        FileExportType = GetExportTypeByExtension(sExtension);

                        if (FileExportType == ExportType.NoType)

                            throw new Exception("Incorrect or no output file extension. Possible options: xls/xlsx/csv");



                        break;

                    case "usetitlefromscript":

                        try

                        {

                            bUseTitleFromScript = IgalTools.StringToBool(item.ProgramParams.FieldValue);

                        }

                        catch (Exception ex)

                        {   throw ex;   }

                        

                        break;

                    case "sqlconnectionforxl":

                        _SqlCon4Xl = item.ProgramParams.FieldValue.Trim();

                        break;

                    //igal 22/7/19 

                    case "suppresstype":

                        switch (item.ProgramParams.FieldValue.ToLower())

                        {

                            case "never":

                                suppress = SuppressType.Never;

                                break;

                            case "ifempty":

                                suppress = SuppressType.IfEmpty;

                                break;

                            case "always":

                                suppress = SuppressType.Always;

                                break;

                            default:

                                throw new Exception("Incorrect or no suppress type. Possible options: Always/IfEmpty/Never");                                

                        }

                        break;

                    default:

                        iParamCount++;

                        switch (sys)

	                    {

                            case SqlSystem.MSSql:                                

                                Array.Resize(ref sqlParam, iParamCount);

                                sqlParam[iParamCount-1] = new SqlParameter(item.ProgramParams.FieldName,item.ProgramParams.FieldValue);

                                break;

                            case SqlSystem.Oracle:

                                Array.Resize(ref oraParam, iParamCount);

                                oraParam[iParamCount-1] = new OracleParameter(item.ProgramParams.FieldName,item.ProgramParams.FieldValue);

                                break;

                            default:

                                break;

	                    }

                        break;

                }

            }



            try

            {

                /*

                if ((sEnv != "" & sVarName == "") | (sVarName != "" & sEnv == ""))

                    throw new Exception("A use of parameters 'env' and 'var_name' requires both");

                else

                {

                    if (sEnv != "" & sVarName != "")

                        ConnectionString = GetConfigurationVar();

                }

                */



                sql = GetSqlFromScript();



                if (bUseTitleFromScript)

                {

                    sFileName = GetFileNameFromScript();

                    sFileName = sFileName.Replace("{timestamp}", sTimeStamp);

                    

                    int index = sFileName.Trim().IndexOfAny(Path.GetInvalidFileNameChars());

                    if (sFileName.Length > 0 && index != -1)

                    {                        

                        throw new Exception(string.Format ("Invalid character(s) in file name [{0}]. index wrong={1}", sFileName ,index.ToString()));                        

                    }



                }



                if (bUseTitleFromScript)

                {

                    if (OutputFolder != "" & sExtension != "" & sFileName.Length > 0)

                        OutputFile = Path.Combine(new DirectoryInfo(OutputFolder).FullName, sFileName + "." + sExtension);

                    else

                        throw new Exception("Error extracting filename from parameters - used UseTitleFromScript switch");

                }

                else

                    if (OutputFile == "")

                        throw new Exception("No output file name specified");



                if (bUseTimeStamp)

                {

                    string sDate = DateTime.Now.ToString("yyyyMMdd_hhmmss");

                    FileInfo fi = new FileInfo(OutputFile);

                    OutputFile = fi.Directory + ((fi.Directory.ToString().Length > 0) ? "\\" : "") + Path.GetFileNameWithoutExtension(fi.Name) + "_" + sDate + fi.Extension;

                }



                if(sys==SqlSystem.Oracle && FileExportType==ExportType.Excel && _SqlCon4Xl.Length==0)

                {

                    throw new Exception("A parameter [SQLConnectionForXl] is obligatory when specified oracle query together with excel export type");

                }



            }

            catch (Exception ex)

            {

                throw ex;

            }

            

        }



        private Export2ExlEngine.ExportType GetExportTypeByExtension(string extension)

        {

            Export2ExlEngine.ExportType thisType = ExportType.NoType;

            switch (extension)

            {

                case "xls":

                case "xlsx":

                    thisType = ExportType.Excel;

                    break;

                case "csv":

                    thisType = ExportType.CSV;

                    break;

            }



            return thisType;

        }



        private string GetFileNameFromScript()

        {

            string sTitle = "";

            if (!sql.ToLower().Contains("sql2xlTitle".ToLower()))

                throw new Exception("Script doesn't contain a 'sql2xlTitle' defenition as expected by using UseTitleFromScript switch");

            if (sTimeStamp.Length > 0 && !sql.ToLower().Contains("{timestamp}"))

                throw new Exception("Script doesn't contain a '{timestamp}' definition as expected by using UseTimeStamp switch");



            string[] lines = sql.Split('\n');            

            

            string sTitleLine = lines.First(l => l.ToLower().Contains("sql2xlTitle:".ToLower()));

            sTitleLine = sTitleLine.Replace("\n", "").Replace("\r","");



            int iStartPos = sTitleLine.ToLower().IndexOf("sql2xlTitle:".ToLower());

            sTitle = sTitleLine.ToLower().Substring(iStartPos + "sql2xlTitle:".ToLower().Length).Trim();



            LoadSheetNames();



            return sTitle;

        }



        private void LoadSheetNames()

        {

            string sSheetName = "";

            int iStartPos = -1, iEndPos = -1;

            do 

            {

                iStartPos = sql.ToLower().IndexOf("sql2xlSubTitle:".ToLower(), iStartPos+1);

                iEndPos = sql.IndexOf("\r\n", iStartPos + 1);

                if (iStartPos >= 0)

                {

                    sSheetName = sql.Substring(iStartPos + "sql2xlSubTitle:".Length, iEndPos - iStartPos - "sql2xlSubTitle:".Length);

                    if(sSheetName.Trim().Length>0)

                        xlSheets.Add(sSheetName.Trim());

                    iStartPos = iEndPos ;

                }

            } while (iStartPos >= 0);

        }



        /*

        private string GetConfigurationVar()

        {            

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("select * from openquery(ALIS_PCLAL,");

            //sb.Append("'select configuration_pck.fget_value_limited@to_common_lts_dep_app(''@env'',''@var'') from dual')");

            sb.Append("'select configuration_pck.fget_value_limited@to_common_lts_dep_app(''" + sEnv + "'',''" + sVarName + "'') from dual')");

            //SqlParameter [] param = new SqlParameter [2];

            //param[0] = new SqlParameter("@env", sEnv);

            //param[1] = new SqlParameter("@var", sVarName);



            //כיון ששליפה מחזירה ערכים שונים באם נעשית משרתים שונים, בודקים קודם אם מריצים בייצור או לא

            //פתרון כנראה זמני 

            string sCon4ConfigurationRequest = "";

            if (sEnv.Contains("prod"))

                sCon4ConfigurationRequest = Properties.Settings.Default.DbCopyConn;

            else

                sCon4ConfigurationRequest = Properties.Settings.Default.DbTest01Conn;



            try

            {

                var result = IgalDAL.SqlDAC.ExecuteScalar(sCon4ConfigurationRequest, CommandType.Text, sb.ToString(), null);

                if (result != null && result.ToString() != "")

                    return result.ToString();

                else

                    throw new Exception("ConfigurationVar didn't return any value");

            }

            catch (Exception ex)

            {

                throw ex; 

            }

        }

        */



        private DataSet RunScript()

        {

            ds = new DataSet();



            if (!IsParametersCorrect())

                return null;



            switch (sys)

            {

                case SqlSystem.MSSql:

                    if (iTimeOut >= 0)

                        SqlDAC.TimeOut = iTimeOut;

                    ds = SqlDAC.ExecuteDataset(ConnectionString, CommandType.Text, sql, sqlParam);

                    OutputFromScript = SqlDAC.OutputFromScript;

                    break;

                case SqlSystem.Oracle:

                    OracleWork ora = new OracleWork();



                    ora.ConnectionString = ConnectionString;



                    ora.Connect();

                    ds = ora.OpenQry(sql, oraParam);

                    break;

            }



            return ds;

        }



        internal int CreateExcel()

        {

            int m_exported = 0;

            Export2Exl XlUtil = null;



            try

            {                

                string sCon = (sys == SqlSystem.Oracle) ? _SqlCon4Xl : ConnectionString;

                XlUtil = new Export2Exl(sCon);

                XlUtil.XlFileName = OutputFile;

                XlUtil.AppendToFile = bAppendToFile;

                XlUtil.CheckFileAndFolderPermissions(true);

                XlUtil.SuppressFileIfEmpty = (suppress == SuppressType.IfEmpty); //igal 22/7/19



                ds = RunScript();

                

                for (int i = 0; i < xlSheets.Count; i++)

                {

                    if (i< ds.Tables.Count )                        

                        ds.Tables[i].TableName = xlSheets[i];

                }

                

                XlUtil.SilentOpen = !bDebug;                

                m_exported = XlUtil.ExportToExcel(ds);

                if(!(suppress == SuppressType.IfEmpty & m_exported==0))

                    XlUtil.SaveFile(OutputFile);                

            }            

            finally

            {

                if (XlUtil != null)

                    XlUtil.Dispose();

            }



            return m_exported;

        }



        internal int CreateCsv()

        {

            int m_exported = 0;

            Export2Csv CsvUtil = new Export2Csv(); ;

            

            try

            {                

                ds = RunScript();                

                CsvUtil.FileName = OutputFile;

                CsvUtil.SuppressFileIfEmpty = (suppress == SuppressType.IfEmpty);

                m_exported = CsvUtil.ExportToCSV(ds);                

            }

            catch (Exception ex)

            {

                throw ex;

            }

            finally

            { if(CsvUtil != null)

                    CsvUtil.Dispose();

            }



            return m_exported;

        }



        private string GetSqlFromScript()

        {

            string sql = "", sErr = "";

            try

            {

                sql = File.ReadAllText(SqlScriptName);

                if (sql.Length == 0)

                    throw new Exception("Sql script file is empty!");

                //if (!SqlDAC.ParseSql(ConnectionString, sql, out sErr))

                //    throw new Exception("Error pasing SQL script \r\n" + sErr);

            }

            catch (FileNotFoundException exF)

            {

                throw exF;

            }

            catch (Exception ex)

            {                

                throw ex;

            }

            return sql;

        }



        private bool IsParametersCorrect()

        {

            if(sys == SqlSystem.NoSystem)

                throw new Exception("No sql system chosen");

            if(SqlScriptName == null || SqlScriptName.Trim().Length == 0)

                throw new Exception("No sql script chosen");

            if (OutputFile == null || OutputFile.Trim().Length == 0) 

            {

                //choose default file name by script

                string sName;

                sName = Path.GetFileNameWithoutExtension(SqlScriptName);

                //sExtension = Path.GetExtension(SqlScriptName);

                OutputFile = sName + "_" + DateTime.Now.ToString("ddMMyy_hhmmss") + "." + sExtension;

            }

            if(ConnectionString.Trim() == "")

            {

                throw new Exception("Connection string empty");

            }



            return true;

        }



        public override string ToString()

        {

            return base.ToString();

        }

    }

}

