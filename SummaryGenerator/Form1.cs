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
using System.Threading;

namespace SummaryGenerator
{
    public partial class Form1 : Form
    {

        // keys stored in app.config file used to show recent file paths
        const String MS4_SRC_FILE_KEY = "ms4SrcFile", MS4_DEST_FILE_KEY = "ms4DestFile", INI_SRC_FILE_KEY = "initiateSrcFile", INI_DEST_FILE_KEY = "initiateDestFile";
        String readAssigneeGroup;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAdd_Incidents_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Initializing.."; 

            // Saving file path in app.config
            Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

            if (rbMs4.Checked)
            {
                readAssigneeGroup = "DHE-PtRevCycle-MS4-AppDev";
                config.AppSettings.Settings[MS4_SRC_FILE_KEY].Value = txtSrcFilePath.Text;
                config.AppSettings.Settings[MS4_DEST_FILE_KEY].Value = txtDestFilePath.Text;
            }

            if (rbInitiate.Checked)
            {
                readAssigneeGroup = "DHE-EMPI-Initiate";
                config.AppSettings.Settings[INI_SRC_FILE_KEY].Value = txtSrcFilePath.Text;
                config.AppSettings.Settings[INI_DEST_FILE_KEY].Value = txtDestFilePath.Text;
            }
            config.Save(ConfigurationSaveMode.Modified);

            // loading data from source sheet 
            DataTable sourceDataTable = loadDataFromSourceExcel();

            if (sourceDataTable.Rows.Count <= 0) {
                MessageBox.Show("No records found.. Aborting..","Something went wrong", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // uploading data in destination sheet
            updateIncidentTracker(sourceDataTable, txtDestFilePath.Text); 
        }

        private DataTable loadDataFromSourceExcel() {
            String whereCondition = null;
            whereCondition = "AND ([Assignment group] = '" + readAssigneeGroup + "' OR [Original Assignment Group] = '" + readAssigneeGroup + "')";

            toolStripStatusLabel1.Text = "Reading records from source file..";
            Thread.Sleep(2000);

            String _null = "";
            String srcSheetName = "Page 1$";
            //String srcConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\..\Desktop\ServiceMattersIncidentTool\incident.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=Yes;""";
            String srcConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtSrcFilePath.Text + @";Extended Properties=""Excel 12.0 Xml;HDR=Yes;""";
            String selCommandText = "SELECT * FROM ['" + srcSheetName + "'] WHERE [Number] <> '" + _null + "' "+ whereCondition  + " ORDER BY [Number]";

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
                MessageBox.Show(e.Message, "Error while reading from source", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor.Current = Cursors.Default;
                toolStripStatusLabel1.Text = "Something went wrong..";
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
            MySheet = MyBook.Sheets["Incident Tracker"];

            // variable declaration
            String incidentNumber, incidentShortDesc, incidentLongDesc, severity, status, resolutionSummary, assigneeIndividual, 
                customerName, incidentOpenDt, incidentAssignDt, incidentResolvedDt, assigneeGroup, originalAssigneeGroup;

            // fetching relevant data from resolution summary
            String keywordCategory = "CATEGORY: ", keywordRootCause = "ROOT CAUSE: ",
                keywordAffItems = "AFFECTED ITEMS/SYSTEMS: ", keywordFacility = "AFFECTED FACILITIES/USERS: ", keywordOutage = "OUTAGE DURATION: ";              ;
            String ticketType = null, category = null,  subCategory = null, affctedItem = null, affectedFacility = null;
            String existingState;

            int noOfIncidentsUpdated = 0, noOfIncidentsAdded = 0, noOfIncidentsSkipped = 0;
            int countOfIncidentFound, i = 1;

            DataRow[] foundRows;
            int totalCount = destDataSet.Rows.Count + 2; // one for blank row and one for getting new row for intertion
            

            try
            {
                foreach (DataRow dataRow in srcDataset.Rows)
                {
                    ticketType = category = subCategory = affctedItem = affectedFacility = null;

                    toolStripStatusLabel1.Text = "Prcessing " + i + " of " + srcDataset.Rows.Count + " incidents..";
                    countOfIncidentFound = -1;

                    incidentNumber = dataRow["Number"].ToString();
                    resolutionSummary = dataRow["Close notes"].ToString().Replace("'", "''");
                    customerName = dataRow["Customer"].ToString();
                    incidentShortDesc = dataRow["Short description"].ToString().Replace("'", "''");
                    incidentLongDesc = dataRow["Description"].ToString().Replace("'", "''");
                    severity = dataRow["Priority"].ToString();
                    status = dataRow["State"].ToString();
                    assigneeGroup = dataRow["Assignment group"].ToString();
                    originalAssigneeGroup = dataRow["Original Assignment Group "].ToString();
                    assigneeIndividual = dataRow["Assigned to"].ToString();
                    incidentOpenDt = dataRow["Opened"].ToString(); // previously it was created
                    incidentAssignDt = dataRow["Actual start"].ToString();
                    incidentResolvedDt = dataRow["Resolved"].ToString();


                    if (customerName == "OMI Integration Service")
                        incidentShortDesc = incidentLongDesc;

                    try
                    {
                        if (resolutionSummary.Length > 10)
                        {

                            if (resolutionSummary.Contains(keywordRootCause))
                                ticketType = MySubString(keywordCategory.Length, resolutionSummary.IndexOf(keywordRootCause) - 2, resolutionSummary);

                            if (resolutionSummary.Contains(keywordCategory))
                            {
                                category = MySubString(resolutionSummary.IndexOf(keywordRootCause) + keywordRootCause.Length, resolutionSummary.IndexOf(keywordAffItems) - 2, resolutionSummary);
                                subCategory = category.Substring(category.IndexOf("-") + 1);
                                category = category.Substring(0, category.IndexOf("-"));
                            }

                            if (resolutionSummary.Contains(keywordAffItems))
                                affctedItem = MySubString(resolutionSummary.IndexOf(keywordAffItems) + keywordAffItems.Length, resolutionSummary.IndexOf(keywordFacility) - 2, resolutionSummary);

                            if (resolutionSummary.Contains(keywordFacility))
                            {
                                affectedFacility = MySubString(resolutionSummary.IndexOf(keywordFacility) + keywordFacility.Length, resolutionSummary.IndexOf(keywordOutage) - 2, resolutionSummary);

                                if (affectedFacility.Contains("/")) // this checks if affected facility is having customer name appeneded
                                    affectedFacility = affectedFacility.Substring(0, affectedFacility.IndexOf("/"));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        
                    }

                    // populating cat. and sub-category if ticket was reassigned to another group
                    if (category == null && subCategory == null && assigneeGroup != originalAssigneeGroup && (status == "Closed" || status == "Resolved")) {
                        category = "Others";
                        subCategory = "Re-assigned";
                    }

                    // setting ticket-type value as not-applicable in case initiate application is selected
                    if (rbInitiate.Checked) {
                        ticketType = "Not Applicable";
                    }

                    // searching for the incident from incidents picked from tracker to know its row-index
                    foundRows = destDataSet.Select("F1 = '" + incidentNumber + "'");

                    if (foundRows.Count() > 0)
                        countOfIncidentFound = destDataSet.Rows.IndexOf(foundRows[0]);

                    if (countOfIncidentFound <= 0)//countOfIncidentFound < 1) // No match found hence insertion will be performed
                    {

                        MySheet.Cells[totalCount, 1] = incidentNumber;

                        if (ticketType != null)
                            MySheet.Cells[totalCount, 2] = ticketType;

                        MySheet.Cells[totalCount, 3] = incidentShortDesc;
                        MySheet.Cells[totalCount, 4] = severity;

                        if (category != null)
                            MySheet.Cells[totalCount, 5] = category;

                        if (subCategory != null)
                            MySheet.Cells[totalCount, 6] = subCategory;

                        if (affectedFacility != null)
                            MySheet.Cells[totalCount, 7] = affctedItem;

                        MySheet.Cells[totalCount, 8] = status;
                        MySheet.Cells[totalCount, 9] = resolutionSummary;
                        MySheet.Cells[totalCount, 10] = assigneeIndividual;
                        MySheet.Cells[totalCount, 11] = assigneeIndividual;
                        MySheet.Cells[totalCount, 12] = assigneeIndividual;
                        MySheet.Cells[totalCount, 13] = customerName;
                        MySheet.Cells[totalCount, 14] = incidentOpenDt;
                        MySheet.Cells[totalCount, 15] = incidentAssignDt;
                        MySheet.Cells[totalCount, 17] = incidentResolvedDt;

                        if (affectedFacility != null)
                            MySheet.Cells[totalCount, 31] = affectedFacility;

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

                            if (ticketType!= null)
                                MySheet.Cells[countOfIncidentFound, 2] = ticketType;

                            MySheet.Cells[countOfIncidentFound, 3] = incidentShortDesc;
                            MySheet.Cells[countOfIncidentFound, 4] = severity;

                            if (category != null)
                                MySheet.Cells[countOfIncidentFound, 5] = category;

                            if (subCategory != null)
                                MySheet.Cells[countOfIncidentFound, 6] = subCategory;

                            if (affectedFacility != null)
                                MySheet.Cells[countOfIncidentFound, 7] = affctedItem;

                            MySheet.Cells[countOfIncidentFound, 8] = status;
                            MySheet.Cells[countOfIncidentFound, 9] = resolutionSummary;
                            MySheet.Cells[countOfIncidentFound, 10] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 11] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 12] = assigneeIndividual;
                            MySheet.Cells[countOfIncidentFound, 13] = customerName;
                            MySheet.Cells[countOfIncidentFound, 14] = incidentOpenDt;
                            MySheet.Cells[countOfIncidentFound, 15] = incidentAssignDt;
                            MySheet.Cells[countOfIncidentFound, 17] = incidentResolvedDt;
                            
                            if (affectedFacility != null)
                                MySheet.Cells[countOfIncidentFound, 31] = affectedFacility;

                            noOfIncidentsUpdated++;
                        }
                        else {
                            noOfIncidentsSkipped++;
                        }
                        
                    }
                    i++;
                }
                toolStripStatusLabel1.Text = "Processing completed !!";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message , "Error while updating..", MessageBoxButtons.OK, MessageBoxIcon.Error);
                toolStripStatusLabel1.Text = "Something went wrong..";
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

            toolStripStatusLabel1.Text = "Reading records from tracker..";
            Thread.Sleep(2000);

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
                MessageBox.Show(e.Message, "Error while reading from tracker");
            }
            finally {
                cmd = null;
                con.Close();
            }

            return dt;
        }

        private String MySubString(int startIndex, int endIndex, String text) {

            char[] textArray = text.ToCharArray();
            String subText = null;

            for (int i = startIndex; i <= endIndex; i++) {
                subText += textArray[i];
            }

            return subText;
        }

        private void btnSrcFileDialog_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;";
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(txtSrcFilePath.Text);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtSrcFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnDestFileDialog_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;";
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(txtDestFilePath.Text);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtDestFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure to Exit !!", "Confirmation..", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
                Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Make sure to close the excel files before running the tool.";
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
