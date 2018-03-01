using System.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.CSharp;
using Excel = Microsoft.Office.Interop.Excel;
using MyMarshal = System.Runtime.InteropServices.Marshal;


namespace SummaryGenerator
{
    public partial class Form1 : Form
    {

        // keys stored in app.config file 
        const String MS4_SRC_FILE_KEY = "ms4SrcFile", MS4_DEST_FILE_KEY = "ms4DestFile", INI_SRC_FILE_KEY = "initiateSrcFile", INI_DEST_FILE_KEY = "initiateDestFile";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAdd_Incidents_Click(object sender, EventArgs e)
        {
            DataTable sourceDataTable = loadDataFromSourceExcel();

            if (sourceDataTable.Rows.Count <= 0) {
                MessageBox.Show("No records found.. Aborting..","Something went wrong", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

            if (rbMs4.Checked)
            {   
                config.AppSettings.Settings[MS4_SRC_FILE_KEY].Value = txtSrcFilePath.Text;
                config.AppSettings.Settings[MS4_DEST_FILE_KEY].Value = txtDestFilePath.Text;
            }

            if (rbInitiate.Checked)
            {
                config.AppSettings.Settings[INI_SRC_FILE_KEY].Value = txtSrcFilePath.Text;
                config.AppSettings.Settings[INI_DEST_FILE_KEY].Value = txtDestFilePath.Text;
            }
            config.Save(ConfigurationSaveMode.Modified);

            updateIncidentTracker(sourceDataTable, txtDestFilePath.Text); 
        }

        private DataTable loadDataFromSourceExcel() {

            String _null = "";
            String srcSheetName = "Page 1$";
            //String srcConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\..\Desktop\ServiceMattersIncidentTool\incident.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=Yes;""";
            String srcConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtSrcFilePath.Text + @";Extended Properties=""Excel 12.0 Xml;HDR=Yes;""";
            String selCommandText = "SELECT * FROM ['" + srcSheetName + "'] WHERE [Number] <> '" + _null + "' ORDER BY [Number]";

            Cursor.Current = Cursors.WaitCursor;
            DataTable dataTable = new DataTable();
            OleDbConnection srcFileConnection = new OleDbConnection(srcConnStr);
            OleDbCommand oleDbCmnd = new OleDbCommand();
            OleDbDataAdapter oleDataAdapter = new OleDbDataAdapter();

            try
            {
                srcFileConnection.Open();
                oleDbCmnd.Connection = srcFileConnection;
                oleDbCmnd.CommandText = selCommandText;

                oleDataAdapter.SelectCommand = oleDbCmnd;
                oleDataAdapter.Fill(dataTable);
                //dataSet.Tables.Add(dataTable);
                //dataTable.Columns["Close Notes"].ColumnName = "CloseNotes";
                //dataTable.Columns.Add("NoteLength", typeof(int), "len(CloseNotes)");
                //dataTable.DefaultView.Sort = "NoteLength desc";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                Cursor.Current = Cursors.Default;
            }
            finally {
                oleDbCmnd = null;
                srcFileConnection.Close();
            }
            return dataTable;
        }

        private void updateIncidentTracker(DataTable srcDataset, String destFilePath) {

            String destSheetName = "Incident Tracker$"; // Destination worksheet name in excel
            DataTable destDataSet =  getAllIncidentsFromDestFile(destSheetName, destFilePath);

            Excel.Workbook MyBook = null;
            Excel.Workbooks xlWorkBooks = null;
            Excel.Application MyApp = null;
            Excel.Worksheet MySheet = null;

            MyApp = new Excel.Application();
            MyApp.Visible = false;
            xlWorkBooks = MyApp.Workbooks;
            MyBook = xlWorkBooks.Open(destFilePath);
            MySheet = MyBook.Sheets[7];

            // variable declaration
            String incidentNumber, incidentDesc, severity, status, resolutionSummary, assigneeIndividual, 
                customerName, incidentOpenDt, incidentActualStartDt, incidentResolvedDt, assigneeGroup;

            String existingState;

            int noOfIncidentsUpdated = 0, noOfIncidentsAdded = 0, noOfIncidentsSkipped = 0;
            int countOfIncidentFound;

            DataRow[] foundRows;
            int totalCount = destDataSet.Rows.Count + 2; // one for blank row and one for getting new row for intertion

            try
            {
                foreach (DataRow dataRow in srcDataset.Rows)
                {
                    countOfIncidentFound = -1;

                    incidentNumber = dataRow["Number"].ToString();
                    resolutionSummary = dataRow["Close notes"].ToString().Replace("'", "''");
                    customerName = dataRow["Customer"].ToString();
                    incidentDesc = dataRow["Short description"].ToString().Replace("'", "''");
                    severity = dataRow["Priority"].ToString();
                    status = dataRow["State"].ToString();
                    assigneeGroup = dataRow["Assignment group"].ToString();
                    assigneeIndividual = dataRow["Assigned to"].ToString();
                    incidentOpenDt = dataRow["Created"].ToString();
                    incidentActualStartDt = dataRow["Actual start"].ToString();
                    incidentResolvedDt = dataRow["Resolved"].ToString();

                    if (incidentResolvedDt == String.Empty)
                        incidentResolvedDt = DateTime.MinValue.ToString();

                    foundRows = destDataSet.Select("F1 = '" + incidentNumber + "'");


                    if (foundRows.Count() > 0)
                        countOfIncidentFound = destDataSet.Rows.IndexOf(foundRows[0]);

                    if (countOfIncidentFound <= 0)//countOfIncidentFound < 1) // No match found hence insertion will be performed
                    {

                        MySheet.Cells[totalCount, 1] = incidentNumber;
                        MySheet.Cells[totalCount, 3] = incidentDesc;
                        MySheet.Cells[totalCount, 4] = severity;
                        MySheet.Cells[totalCount, 8] = status;
                        MySheet.Cells[totalCount, 9] = resolutionSummary;
                        MySheet.Cells[totalCount, 10] = assigneeIndividual;
                        MySheet.Cells[totalCount, 11] = assigneeIndividual;
                        MySheet.Cells[totalCount, 12] = assigneeIndividual;
                        MySheet.Cells[totalCount, 13] = customerName;
                        MySheet.Cells[totalCount, 14] = incidentOpenDt;
                        MySheet.Cells[totalCount, 16] = incidentResolvedDt;

                        totalCount++;
                        noOfIncidentsAdded++;
                    }
                    else
                    { // match found hence update will be performed

                        //MySheet.Cells[countOfIncidentFound + 2, 1] = incidentNumber;
                        //MySheet.Cells[countOfIncidentFound + 2, 9] = resolutionSummary;

                        countOfIncidentFound += 2;
                        existingState = (MySheet.Cells[countOfIncidentFound, 8]).Value.ToString();

                        if (existingState != "Closed")
                        {
                            MySheet.Cells[countOfIncidentFound, 1] = incidentNumber;
                            MySheet.Cells[countOfIncidentFound, 3] = incidentDesc;
                            MySheet.Cells[countOfIncidentFound, 4] = severity;
                            MySheet.Cells[countOfIncidentFound, 8] = status;
                            MySheet.Cells[countOfIncidentFound, 9] = resolutionSummary;
                            MySheet.Cells[countOfIncidentFound, 10] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 11] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 12] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 13] = customerName;
                            MySheet.Cells[countOfIncidentFound, 14] = incidentOpenDt;
                            MySheet.Cells[countOfIncidentFound, 16] = incidentResolvedDt;

                            noOfIncidentsUpdated++;
                        }
                        else {
                            noOfIncidentsSkipped++;
                        }
                        
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message , "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally {
                MyBook.Save();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                MyBook.Close(false, destFilePath, null);
                MyApp.Quit();

                MyMarshal.ReleaseComObject(MySheet);
                MyMarshal.ReleaseComObject(xlWorkBooks);
                MyMarshal.ReleaseComObject(MyBook);
                MyMarshal.ReleaseComObject(MyApp);
            }

            int totalTkts = noOfIncidentsUpdated + noOfIncidentsAdded + noOfIncidentsSkipped;
            Cursor.Current = Cursors.Default;
            MessageBox.Show("Tracker has been updated successfully with: " +
                "\n Rows Added: "+ noOfIncidentsAdded + " " +
                "\n Rows Updated: "+ noOfIncidentsUpdated +
                "\n Rows Skipped: "+ noOfIncidentsSkipped +"" +
                "\n Total Rows: "+ totalTkts + "","Status", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private DataTable getAllIncidentsFromDestFile(String destSheetName, String destFilePath)
        {
            String _null = "";
            String selectQuery = "SELECT F1 FROM ['" + destSheetName + "'] WHERE [F1] <> '" + _null + "'";
            //String destConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;
            //                      Data Source= C:\Users\128953\Desktop\ServiceMattersIncidentTool\MS4_IncidentTracker_QR202_2018.xlsm;
            //                      Extended Properties='Excel 12.0 XML;HDR=NO'";
            String destConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;
                                  Data Source= '" + destFilePath + "';" +
                                  @"Extended Properties='Excel 12.0 XML;HDR=NO'";

            OleDbConnection con = new OleDbConnection(destConnStr);
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter sda = new OleDbDataAdapter();
            DataTable dt = new DataTable();

            try
            {
                con.Open();
                cmd.CommandText = selectQuery;
                cmd.Connection = con;
                
                sda.SelectCommand = cmd;
                sda.Fill(dt);
            }
            catch (Exception e) {
            }
            finally {
                cmd = null;
                con.Close();
            }

            return dt;
        }

        private void btnSrcFileDialog_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;";
            openFileDialog.InitialDirectory = txtSrcFilePath.Text;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtSrcFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnDestFileDialog_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;";
            openFileDialog.InitialDirectory = txtDestFilePath.Text;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtDestFilePath.Text = openFileDialog.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            rbMs4.Checked = true;
        }

        private void rbMs4_CheckedChanged(object sender, EventArgs e)
        {
            if (rbMs4.Checked)
            {
                txtSrcFilePath.Text = ConfigurationManager.AppSettings.Get(MS4_SRC_FILE_KEY);
                txtDestFilePath.Text = ConfigurationManager.AppSettings.Get(MS4_DEST_FILE_KEY);
            }
        }

        private void rbInitiate_CheckedChanged(object sender, EventArgs e)
        {
            if (rbInitiate.Checked)
            {
                txtSrcFilePath.Text = ConfigurationManager.AppSettings.Get(INI_SRC_FILE_KEY);
                txtDestFilePath.Text = ConfigurationManager.AppSettings.Get(INI_DEST_FILE_KEY);
            }
        }
    }
}
