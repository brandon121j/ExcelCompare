using ClosedXML.Excel;
using IronXL;
using IronXL.Formatting;
using IronXL.Formatting.Enums;
using IronXL.Styles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelComparer
{
    public partial class Form2 : Form
    {
        public string ExcelSheet1 { get; set; }
        public string ExcelSheet2 { get; set; }
        public string FormTitle { get; set; }
        public List<string> Columns { get; set; }
        public string Filter { get; set; }
        public (int, int, int, int) Counts { get; set; }
        private static bool ColorBlindMode { get; set; }
        private Color Red { get; set; } = Color.FromArgb(255, 0, 0);
        private Color Green { get; set; } = Color.FromArgb(19, 174, 75);
        private Color Yellow { get; set; } = Color.FromArgb(247, 222, 58);
        private Color Orange { get; set; } = Color.FromArgb(255, 140, 0);
        public (DataTable, DataTable, DataTable, DataTable) ResultsDT { get; set; }
        public string AdditionsCMD { get { return $"SELECT * into Additions FROM Sheet1 WHERE NOT EXISTS (select 1 from Sheet2 where Sheet1.{Filter} = Sheet2.{Filter})"; } }
        public string MatchesCMD { get { return $"SELECT * into Matches FROM Sheet1 where exists (select 1 from Sheet2 where {Filter} = Sheet1.{Filter})"; } }
        public string RemovedCMD { get { return $"SELECT * into Removed FROM Sheet2 where NOT EXISTS (select 1 from Sheet1 where Sheet1.{Filter} = Sheet2.{Filter})"; } }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            this.Text = FormTitle;
            changedItemLabel.BackColor = Orange;
            newItemLabel.BackColor = Green;
            sameItemLabel.BackColor = Yellow;
            removedItemLabel.BackColor = Red;
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            // Converts Excel File to DataTable so it can be pushed to SQL later on
            var (excelData1, excelData2) = ExcelFileReader(ExcelSheet1, ExcelSheet2);

            // Sets the Filter and Columns Object Property so we can access them globally 
            Filter = FilterPrompt(Columns = ColumnFetcher());

            // Filter will only be empty if user closes out of FilterPrompt 
            if (!string.IsNullOrEmpty(Filter))
                Comparer(excelData1, excelData2);
        }

        private List<string> ColumnFetcher()
        {
            List<string> columnHeaders = new List<string>();

            WorkBook workbook = WorkBook.Load(ExcelSheet1);

            WorkSheet worksheet = workbook.WorkSheets.First();

            for (int i = 0; i < worksheet.Columns.Count(); i++)
            {
                columnHeaders.Add(worksheet.Rows[0].Columns[i].Value.ToString());
            }

            return columnHeaders;
        }

        // Popup that prompts user to select a unique column as filter
        private string FilterPrompt(List<string> columnHeaders)
        {
            Form3 filterPrompt = new Form3(columnHeaders);

            filterPrompt.StartPosition = FormStartPosition.CenterParent;

            string filter = "";

            if (filterPrompt.ShowDialog(this) == DialogResult.OK)
            {
                filter = filterPrompt.SelectedFilter;
            }
            else { ReturnToCompareForm(true); }


            return filter;
        }

        //Converts Excel Sheet to DataTable
        private (DataTable, DataTable) ExcelFileReader(string path1, string path2)
        {
            DataTable excel1 = new DataTable();
            DataTable excel2 = new DataTable();

            for (var i = 0; i < 2; i++)
            {
                var path = i == 1 ? path1 : path2;

                WorkBook workbook = WorkBook.Load(path);

                WorkSheet worksheet = workbook.DefaultWorkSheet;

                worksheet.Name = "Sheet1";

                if (i == 1)
                {
                    excel1 = worksheet.ToDataTable(true);
                } else { excel2 = worksheet.ToDataTable(true); }

            }

            return (excel1, excel2);
        }

        private void Comparer(DataTable excel1, DataTable excel2)
        {
            try
            {
                DatabaseManager dbm = new DatabaseManager();

                using (SqlConnection sqlcon = dbm.CreateConnection())
                {
                    excel2.TableName = "Sheet2";
                    DataTableExporter(sqlcon, excel1);
                    DataTableExporter(sqlcon, excel2);
                    SQLDataFetcher(sqlcon, nameof(AdditionsCMD), AdditionsCMD);
                    SQLDataFetcher(sqlcon, nameof(MatchesCMD), MatchesCMD);
                    SQLDataFetcher(sqlcon, nameof(RemovedCMD), RemovedCMD);
                    ChangesFetcher(sqlcon);

                    var (changes, matches, additions, removed) = TableFetcher(sqlcon);
                    ResultsDT = (changes, matches, additions, removed);
                    Counts = (RowAdder(changes), RowAdder(matches), RowAdder(additions), RowAdder(removed));
                    CountUpdater(changes.Rows.Count, matches.Rows.Count, additions.Rows.Count, removed.Rows.Count);
                    RowFormatter(Counts);
                    sqlcon.Close();
                }
            }
            finally
            {
                
            }

        }

        // Exports DataTable to SQL
        private void DataTableExporter(SqlConnection sqlcon, DataTable dataTable)
        {

            if (sqlcon.State != ConnectionState.Open)
                sqlcon.Open();

            TableExists(sqlcon, dataTable.TableName);

            StringBuilder sqlBuilder = new StringBuilder();

            sqlBuilder.Append($"CREATE TABLE {dataTable.TableName} (");

            Appender(sqlBuilder, " VARCHAR(8000),", true);

            new SqlCommand(sqlBuilder.ToString(), sqlcon).ExecuteNonQuery();

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlcon))
            {
                foreach (DataColumn col in dataTable.Columns)
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);

                bulkCopy.DestinationTableName = dataTable.TableName;
                try
                {
                    bulkCopy.WriteToServer(dataTable);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }

        // Tests if table exists if it does it deletes it
        private void TableExists(SqlConnection sqlcon, string tableName)
        {
            var existsCmd = @"DROP TABLE IF EXISTS " + tableName;
            SqlCommand tableStatus = new SqlCommand(existsCmd, sqlcon);
            tableStatus.ExecuteNonQuery();
        }

        // Creates SQL Table w/ columns from DataTable
        private (DataTable, DataTable, DataTable, DataTable) TableFetcher(SqlConnection sqlcon)
        {
            ColumnAdder(excelSheet1DG);

            DataTable changes = TableConverter(sqlcon, "Changes");
            DataTable matches = TableConverter(sqlcon, "Matches");
            DataTable additions = TableConverter(sqlcon, "Additions");
            DataTable removed = TableConverter(sqlcon, "Removed");

            return (changes, matches, additions, removed);
        }

        // Converts SQL table to DataTable 
        private DataTable TableConverter(SqlConnection sqlcon, string tableName)
        {
            SqlCommand cmd = new SqlCommand($"SELECT * FROM {tableName}", sqlcon);

            DataTable dt = new DataTable();

            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                adapter.Fill(dt);
            }

            dt.TableName = tableName;

            return dt;
        }

        private void SQLDataFetcher(SqlConnection sqlcon, string tableName, string cmd)
        {
            tableName = tableName.Substring(0, tableName.Length - 3);

            TableExists(sqlcon, tableName);

            new SqlCommand(cmd, sqlcon).ExecuteNonQuery();
        }

        // Creates Changes table and removes those items from Matches table
        private void ChangesFetcher(SqlConnection sqlcon)
        {
            TableExists(sqlcon, "Changes");
            TableExists(sqlcon, "Differences");

            new SqlCommand($"SELECT * INTO Differences FROM Sheet2 WHERE EXISTS (SELECT 1 FROM Matches WHERE {Filter} = Sheet2.{Filter})", sqlcon).ExecuteNonQuery();
            new SqlCommand($"SELECT * into Changes FROM Matches EXCEPT SELECT * FROM Differences", sqlcon).ExecuteNonQuery();

            new SqlCommand($"DELETE FROM Matches WHERE EXISTS (SELECT 1 FROM Changes WHERE Changes.{Filter} = Matches.{Filter})", sqlcon).ExecuteNonQuery();
        }

        // Adds column values to the search criteria for the StringBuilder
        private StringBuilder Appender(StringBuilder sb, string text = "", bool parenthesis = false)
        {
            foreach (var item in Columns)
            {
                sb.Append($"{item}{text} ");
            }

            if (sb.ToString().EndsWith(", "))
                sb = sb.Remove(sb.Length - 2, 1);

            if (parenthesis)
                sb.Append(")");


            return sb;
        }

        private void CountUpdater(int changes, int matches, int additions, int removed)
        {
            excelChangedCount.Text = $"Changed Items: {changes}";
            excelNewCount.Text = $"New Items: {additions}";
            excelSameCount.Text = $"Same Items: {matches}";
            excelRemovedCount.Text = $"Removed Items: {removed}";
            excelTotalCount.Text = $"Total Items: {changes + matches + additions + removed}";
        }

        // Save File Dialog Prompt
        private string SaveFile(string filter, string ext)
        {

            SaveFileDialog saveFile = new SaveFileDialog()
            {
                Title = "Save File",
                DefaultExt = ext,
                Filter = filter,
                InitialDirectory = @"C:\\documents",
            };

            saveFile.ShowDialog();

            return saveFile.FileName;
        }

        private void toolStripCopyButton_Click(object sender, EventArgs e)
        {
            var newline = System.Environment.NewLine;
            var tab = "\t";
            var clipboard_string = new StringBuilder();

            foreach (DataGridViewRow row in excelSheet1DG.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (i == (row.Cells.Count - 1))
                        clipboard_string.Append(row.Cells[i].Value + newline);
                    else
                        clipboard_string.Append(row.Cells[i].Value + tab);
                }
            }

            Clipboard.SetText(clipboard_string.ToString());

            MessageBox.Show("Data copied to clipboard!", "Success!");
        }

        // Popup displayed when closing application
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Application.OpenForms.Count == 2)
            {
                var formToShow = Application.OpenForms.Cast<Form>()
                    .FirstOrDefault(c => c is Form1);


                if (formToShow != null)
                {
                    formToShow.WindowState = FormWindowState.Normal;
                }
                else
                {
                    Application.Exit();
                }

            }

            this.Dispose();

        }

        // Exports to .xlsx
        private void ExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                DatabaseManager dbm = new DatabaseManager();

                using (SqlConnection connection = dbm.CreateConnection())
                {
                    var (changes, matches, additions, removed) = ResultsDT;

                    XLWorkbook wb = new XLWorkbook();

                    wb.Worksheets.Add(additions, "New Items");

                    wb.Worksheets.Add(removed, "Removed Items");

                    wb.Worksheets.Add(changes, "Changed Items");

                    wb.Worksheets.Add(matches, "Same Items");

                    wb.SaveAs(SaveFile("Excel Files| *.xlsx; *.xlsm", ".xlsx"));

                    MessageBox.Show("Excel file saved!", "Success!");

                    connection.Close();
                }            
            }
            catch
            {
                MessageBox.Show("An error has occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Handles clicking the "New" button
        private void toolStripNewButton_Click_1(object sender, EventArgs e)
        {
            ReturnToCompareForm(false);
        }

        private void colorblindModeOffToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorBlindMode = !ColorBlindMode;
            Green = ColorBlindMode ? Color.FromArgb(155, 191, 133) : Color.FromArgb(19, 174, 75);
            Red = ColorBlindMode ? Color.FromArgb(202, 0, 32) : Color.FromArgb(255, 0, 0);
            Yellow = ColorBlindMode ? Color.FromArgb(244, 165, 130) : Color.FromArgb(247, 222, 58);
            Orange = ColorBlindMode ? Color.FromArgb(255, 133, 59) : Color.FromArgb(255, 140, 0);
            colorblindModeOffToolStripMenuItem.Text = ColorBlindMode ? "Colorblind Mode: On" : "Colorblind Mode: Off";
            changedItemLabel.BackColor = Orange;
            newItemLabel.BackColor = Green;
            sameItemLabel.BackColor = Yellow;
            removedItemLabel.BackColor = Red;
            RowFormatter(Counts);
        }

        // Returns to the first form that pops up
        private void ReturnToCompareForm(bool closeForm)
        {
            var formToShow = Application.OpenForms.Cast<Form>()
                .FirstOrDefault(c => c is Form1);

            if (formToShow != null)
            {
                formToShow.WindowState = FormWindowState.Normal;
                if (closeForm) { Application.OpenForms["Form2"].Dispose(); }
            }
        }

        // Adds columns to DataGridView
        private DataGridView ColumnAdder(DataGridView dgv)
        {
            foreach(var item in Columns)
            {
                dgv.Columns.Add(item, item);
            }

            return dgv;
        }

        // Adds the rows from the SQL Tables
        private int RowAdder(DataTable dt)
        {
            var dg = excelSheet1DG;

            foreach(DataRow item in dt.Rows)
            {
                excelSheet1DG.Rows.Add(item.ItemArray);
            }

            return excelSheet1DG.Rows.Count;
        }
        
        // Adds color to cells 
        private void RowFormatter((int, int, int, int) counts)
        {
            var (changes, matches, additions, removed) = counts;

            for (var i = 0; i < excelSheet1DG.Rows.Count; i++)
            {
                if (i < changes)
                {
                    excelSheet1DG.Rows[i].DefaultCellStyle.BackColor = Orange;
                }
                else if (i < matches)
                {
                    excelSheet1DG.Rows[i].DefaultCellStyle.BackColor = Yellow;
                }
                else if (i < additions)
                { 
                    excelSheet1DG.Rows[i].DefaultCellStyle.BackColor = Green; 
                } else if (i >= additions)
                {
                    excelSheet1DG.Rows[i].DefaultCellStyle.BackColor = Red;
                }
            }
        }

        private void excelSheet1DG_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            Clipboard.SetDataObject(this.excelSheet1DG.GetClipboardContent());
            Clipboard.SetText(Clipboard.GetText());

        }
    }

}
