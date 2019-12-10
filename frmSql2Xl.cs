﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Export2ExlEngine;
using IgalDAL;
using System.Configuration;

namespace Sql2Xl
{
    public partial class frmSqlExport : Form
    {
        bool _IsInitializing = false;
        string sCon = "", sServer, sDB;
        DataSet dsAllQrys;
        Export2Exl XlUtil;
        Export2Csv CsvUtil;
        Export2ExlEngine.ExportType xportType = ExportType.Excel;

        #region Form Basic Actions
        public frmSqlExport()
        {
            _IsInitializing = true;
            InitializeComponent();
            this.Resize += new System.EventHandler(ResizeMe);
            string[] myServers = ConfigurationManager.AppSettings["DBServers"].Split(',');
            cmbServers.Items.AddRange(myServers);
            this.WindowState = FormWindowState.Maximized;
        }
        private void frmSqlExport_Load(object sender, EventArgs e)
        {
            optXl.Checked = true;
            LoadSettings(); 
        }

        private void frmSqlExport_Activated(object sender, EventArgs e)
        {
            if (!_IsInitializing) return;

            _IsInitializing = false;
        }

        private void frmSqlExport_FormClosing(object sender, FormClosingEventArgs e)
        {            SaveSettings();        }

        #endregion

        #region Visuality
        private void p_EnableButtons(bool p)
        {            
            btnExlExport.Enabled = p;
            btnShow.Enabled = p;
        }

        private void ResizeMe(object sender, EventArgs e)
        {
            btnExlExport.Top = this.Height - btnExlExport.Height - statusStrip1.Height - 60;
            btnExportAll.Top = btnExlExport.Top;

            grdSummary.Width = this.Width - 40;
            grdSummary.Height = btnExportAll.Top - 10 - grdSummary.Top;
            
            txtSql.Width = this.Width - 40;

            grpExportType.Top = this.Height - grpExportType.Height - 60;
        }

        #endregion

        #region Additional Loading
        private void SaveSettings()
        {
            Properties.Settings.Default.MyServer = cmbServers.Text;
            Properties.Settings.Default.MyDB = cmbDB.Text;          
            Properties.Settings.Default.Save();            
        }

        private void LoadSettings()
        {
            cmbServers.Text = Properties.Settings.Default.MyServer;
            cmbDB.SelectedIndex = -1;
            cmbDB.Text = Properties.Settings.Default.MyDB;
            if (cmbDB.SelectedIndex == -1)
                cmbDB.Text = "";            
        }

        #endregion

        #region Sql section
        private void GetDataBaseList()
        {
            sCon = "server=" + cmbServers.Text + ";initial catalog=master;Integrated Security=true;";
            try
            {
                cmbDB.DataSource = DataBaseList(chkExcludeSystem.Checked);
                cmbDB.DisplayMember = "name";
                cmbDB.ValueMember = "name";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string BuildConnString()
        {
            sCon = "server=" + sServer + ";initial catalog=" + sDB + ";Integrated Security=true;";
            return sCon;
        }

        private DataTable DataBaseList(bool bExcludeSystemDB=true)
                {
                    DataSet dbList = new DataSet();
                    string sql = "select [name]  from sys.databases ";
                    sql += ((bExcludeSystemDB) ? "where [name] not in ('master','model','msdb','tempdb')" : "");

                    if (sCon == "")
                        return null;

                    dbList = SqlDAC.ExecuteDataset(sCon, CommandType.Text, sql, null);

                    return (dbList.Tables.Count > 0) ? dbList.Tables[0] : null;
                }

        #endregion

        #region Controls Behaviour
        private void cmbServers_SelectedIndexChanged(object sender, EventArgs e)
        { GetDataBaseList(); }       

        private void cmbDB_SelectedIndexChanged(object sender, EventArgs e)
        {            sCon = BuildConnString();        }

        private void cmbServers_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbDB.Items.Count > 0)
                { 
                    cmbDB.DataSource=null;
                    cmbDB.Text = "";
                }
                sServer = cmbServers.Text.Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); 
            }

        }

        private void cmbDB_TextChanged(object sender, EventArgs e)
        {            sDB = cmbDB.Text.Trim();        }

        private void cmbServers_Validated(object sender, EventArgs e)
        {            GetDataBaseList();        }

        private void chkExcludeSystem_CheckedChanged(object sender, EventArgs e)
        {            GetDataBaseList(); }

        private void cmbTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_IsInitializing == true)
                return;

            int iIndex = (sender as ComboBox).SelectedIndex;
            ShowData(iIndex);
        } 

        private void btnShow_Click(object sender, EventArgs e)
        {
            string ErrMsg = "";

            if (txtSql.Text.Trim()=="")
            {
                MessageBox.Show("No SQL query");
                return;
            }

            sCon = BuildConnString();            

            if (!SqlDAC.ParseSql(sCon, txtSql.Text, out ErrMsg))
            {
                MessageBox.Show(ErrMsg, "ScriptCreate error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            cmbTables.Items.Clear();            
            try
            {
                dsAllQrys = SqlDAC.ExecuteDataset(sCon, CommandType.Text, txtSql.Text, null);
                for (int i = 0; i < dsAllQrys.Tables.Count; i++)
                {
                    cmbTables.Items.Add(dsAllQrys.Tables[i].TableName);
                }
                if (cmbTables.Items.Count > 0)
                    cmbTables.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtServer_Validated(object sender, EventArgs e)
        {
            sServer = cmbServers.Text.Trim();
        }

        private void txtDB_Validated(object sender, EventArgs e)
        {
            sDB = cmbDB.Text.Trim();
        }

        private void btnExlExport_Click(object sender, EventArgs e)
        {
            if (grdSummary.Rows.Count <= 0)
            {
                MessageBox.Show("אין נתונים לייצא", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }         
            
            int m_exported = 0;
            IDisposable dispose = null;

            try
            {                
                this.Cursor = Cursors.WaitCursor;
                p_EnableButtons(false);

                statusStrip1.Items[0].Text = "מבצע ייצוא נתונים לאקסל ";                
                DataSet ds = new DataSet();
                ds.Tables.Add((grdSummary.DataSource as DataTable).Copy());

                switch (xportType)
                {
                    case ExportType.Excel:
                        XlUtil = new Export2Exl(sCon);
                        dispose = XlUtil;
                        m_exported = XlUtil.ExportToExcel(ds);
                        break;
                    case ExportType.CSV:
                        CsvUtil = new Export2Csv();
                        dispose = CsvUtil;
                        m_exported = CsvUtil.ExportToCSV(ds);
                        break;
                    default:
                        break;
                }   

                
                statusStrip1.Items[0].Text = m_exported.ToString() + " שורות יו~ּצאו בהצלחה";
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message);  }
            finally
            {
                m_exported = 0;
                p_EnableButtons(true);
                this.Cursor = Cursors.Arrow;
                dispose.Dispose(); 
            }
        }

        private void btnExportAll_Click(object sender, EventArgs e)
        {

            int m_exported = 0;

            try
            {
                this.Cursor = Cursors.WaitCursor;
                p_EnableButtons(false);

                statusStrip1.Items[0].Text = "מבצע ייצוא נתונים לאקסל ";

                XlUtil = new Export2Exl(sCon);
                m_exported = XlUtil.ExportToExcel(dsAllQrys);
                statusStrip1.Items[0].Text = m_exported.ToString() + " שורות יו~ּצאו בהצלחה";
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message);  }
            finally
            {
                m_exported = 0;
                p_EnableButtons(true);
                this.Cursor = Cursors.Arrow;
                XlUtil.Dispose();
            }
        }

        private void ShowData(int Index)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                grdSummary.DataSource = dsAllQrys.Tables[Index];

                grdSummary.ColumnHeadersVisible = true;
            }
            catch (System.Data.DataException ex)
            { MessageBox.Show(ex.Message); }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
            finally
            { this.Cursor = Cursors.Default; }
        }

        #endregion

        private void optXl_CheckedChanged(object sender, EventArgs e)
        {
            xportType = ExportType.Excel;
        }

        private void optCSV_CheckedChanged(object sender, EventArgs e)
        {
            xportType = ExportType.CSV;
        }
    }
}
