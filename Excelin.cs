using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace Excelin
{
    public partial class Excelin : Form
    {
        public Excelin()
        {
            InitializeComponent();
        }
        public static string connectionString;
        public static DataTable dtExcelRecords;// = new DataTable();
        public static OleDbConnection con;// = new OleDbConnection(connectionString);
        public static OleDbCommand cmd;// = new OleDbCommand();
        public static OleDbDataAdapter dAdapter;//

        private void FillDropDownList(DataTable dtOPType)
        {
            comboBox1.Items.Clear();
            foreach (DataRow drState in dtOPType.Rows)
            {
                comboBox1.Items.Add(drState["Table_Name"].ToString());
            }
 
            comboBox1.SelectedIndex = 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "xls| *.xls|xlsx|*.xlsx";  //"Text files | *.txt"; // file types, that will be allowed to upload
            dialog.Multiselect = false; // allow/deny user to upload more than one file at a time
            if (dialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                String path = dialog.FileName; // get name of file
                //using (StreamReader reader = new StreamReader(new FileStream(path, FileMode.Open), new UTF8Encoding())) // do anything you want, e.g. read it
                //{
                //    // ...
                //}


                if (dialog.CheckFileExists)
                {
                    string fileName = Path.GetFileName(dialog.FileName);
                    string fileExtension = Path.GetExtension(dialog.FileName);

 
                    if (fileExtension == ".xls")
                    {
                        connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                          path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (fileExtension == ".xlsx")
                    {
                        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                          path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }

                    con = new OleDbConnection(connectionString);
                    cmd = new OleDbCommand();
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.Connection = con;
                    dAdapter = new OleDbDataAdapter(cmd);
                    dtExcelRecords = new DataTable();
                    con.Open();
                    DataTable dtExcelSheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    FillDropDownList(dtExcelSheetName);

                    string getExcelSheetName = dtExcelSheetName.Rows[0]["Table_Name"].ToString();
                    adapt(getExcelSheetName);
                }
            }
        }
        public void adapt(string getExcelSheetName)
        {
            dtExcelRecords = new DataTable();
            dataGridView1.DataSource = null;
            while (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            dataGridView1.Rows.Clear();
            cmd.CommandText = "SELECT * FROM [" + getExcelSheetName + "]";
            dAdapter.SelectCommand = cmd;
            dAdapter.Fill(dtExcelRecords);
            dataGridView1.DataSource = dtExcelRecords;
            // dataGridView1.DataBind(); }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            ComboBox comboBox = (ComboBox)sender;
            //while (dataGridView1.Rows.Count > 0)
            //{
            //    dataGridView1.Rows.RemoveAt(0);
            //}
            string getExcelSheet  = (string)comboBox1.SelectedItem;
         

            //MessageBox.Show(getExcelSheet);

            adapt(getExcelSheet);  
        }
                private void button2_Click(object sender, EventArgs e)
        {
            string sqlConnectionString = @"Data Source=.;Initial Catalog=CCOps; uid=User; Password = 123456;Connection Timeout=120;Integrated Security=SSPI;";
            // Create DbDataReader to Data Worksheet
            using (DbDataReader dr = cmd.ExecuteReader())
            {

                // SQL Server Connection String

                // Bulk Copy to SQL Server
                using (SqlBulkCopy bulkCopy =
                           new SqlBulkCopy(sqlConnectionString))
                {
                    bulkCopy.DestinationTableName = "PTS_T_Tmp_PREGRINE";
                    bulkCopy.WriteToServer(dr);
                    MessageBox.Show("The data has been exported succefuly from Excel to SQL");

                }
                ExecuteSQLStr("exec CCOPS.dbo.PTS_SP_PREGRINE_FILE_FIN", sqlConnectionString);
            }
        }
        private static void ExecuteSQLStr(string queryString,
    string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(
                       connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
        }
    }
}
