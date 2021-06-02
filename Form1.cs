using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Title = "Select file";
            openFileDialog.InitialDirectory = @"c:\";           
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
           
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_path.Text = openFileDialog.FileName;
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            string filePath = txt_path.Text;
            string extension = Path.GetExtension(filePath);
            string conStr="";

            switch (extension)
            {
                case ".xls": //Excel 97-03
                    conStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'", filePath, "YES");
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'", filePath, "YES");
                    break;
            }

            OleDbConnection conn = new OleDbConnection(conStr);

            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("Select * from ["+txt_sheet.Text+"$]", conn);
            DataTable dataTable = new DataTable();

            oleDbDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;


            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;
        }

        private void txt_format_Click(object sender, EventArgs e)
        {
            Int32 selectedColumnCount = dataGridView1.Columns
        .GetColumnCount(DataGridViewElementStates.Selected);

            if (selectedColumnCount > 0)
            {
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    if (cell.Value == null)
                    {
                        break;
                    }
                  cell.Value= cell.Value.ToString().ToLower();
                } 
            }
        }

        private void btn_to_upper_Click(object sender, EventArgs e)
        {
            Int32 selectedColumnCount = dataGridView1.Columns
                  .GetColumnCount(DataGridViewElementStates.Selected);

            if (selectedColumnCount > 0)
            {

                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    if (cell.Value == null)
                    {
                        break;
                    }
                    cell.Value = cell.Value.ToString().ToUpper();
                }

            }
        }

        private void btn_email_Click(object sender, EventArgs e)
        {
            Int32 selectedColumnCount = dataGridView1.Columns
                 .GetColumnCount(DataGridViewElementStates.Selected);
            List<char> chars;
            List<char> charsM = new List<char>();
            string value="";

            if (selectedColumnCount > 0)
            {

                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    if (cell.Value == null)
                    {
                        break;
                    }

                    chars = cell.Value.ToString().ToList();
                    charsM = cell.Value.ToString().ToList();

                    foreach (var item in chars)
                    {
                        if (item == ' ')
                        {
                            charsM.Remove(item);
                        }
                    }

                    value = "";

                    foreach (var item in charsM)
                    {
                        value += item;
                    }

                    cell.Value = value;

                }

            }
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void txt_export_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

        }
    }
}
