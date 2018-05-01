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
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.OracleClient;

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
            string getExcelSheet = (string)comboBox1.SelectedItem;


            //MessageBox.Show(getExcelSheet);

            adapt(getExcelSheet);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sqlConnectionStringyeni = @"Data Source=;Initial Catalog=; uid=; Password = ;Connection Timeout=120;";
            string sqlConnectionString = @"Data Source=;Initial Catalog=; uid=; Password = ;Connection Timeout=120;";
            // Create DbDataReader to Data Worksheet

            try
            {
                using (DbDataReader dr = cmd.ExecuteReader())
                {
                    using (SqlBulkCopy bulkCopy =
                               new SqlBulkCopy(sqlConnectionString))
                    {
                        bulkCopy.DestinationTableName = "dbo.PTS_T_Tmp_PREGRINE";
                        bulkCopy.WriteToServer(dr);


                    }
                    MessageBox.Show("Data aktarildi, Prosedur bekleniyor..");
                    ExecuteSQLStr("exec CCOPS.dbo.PTS_SP_PREGRINE_FILE_FIN", sqlConnectionStringyeni);
                    MessageBox.Show("Prosedur tamamlandı. Mail bekleniyor..");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }
        private static void ExecuteSQLStr(string queryString, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(
                       connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }

            Application.Exit();
        }
        //public DataTable OraDt(string cmdstr)
        //{
        // string constr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=ed02-scan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=edw2)));User Id=" + USR + ";Password=" + PASS + ";Integrated Security=no;";



        //OracleDataAdapter adapter = new OracleDataAdapter(cmdstr, constr);

        //OracleCommandBuilder builder = new OracleCommandBuilder(adapter);


        //DataSet dataset = new DataSet();
        //adapter.Fill(dataset, "EMP");

        //DataTable dt = dataset.Tables["EMP"];
        //    return dt;
        //}


        public void createsqltable(string strconnection, DataTable dt, string tablename)
        {
            try
            { //  = "";
                string table = "";
                //table += "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + tablename + "]') AND type in (N'U'))";
                //table += "BEGIN ";
                table += "create table " + tablename.ToUpper().ToString() + "";
                table += "(";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i != dt.Columns.Count - 1)
                        table += dt.Columns[i].ColumnName + " " + "    VARCHAR2(100 BYTE)" + ",";
                    else
                        table += dt.Columns[i].ColumnName + " " + "  VARCHAR2(100 BYTE)";
                }
                table += ") GO COMMIT  ";
                //table += "END";
              //  table += "; commit;";
                // InsertQuery(table, strconnection);
                //   CopyData(strconnection, dt, tablename);

                OleDbConnection myConnection = new OleDbConnection(strconnection);
                OleDbCommand myCommand = new OleDbCommand(table, myConnection);
                myConnection.Open();
                myCommand.ExecuteNonQuery();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            string sunucu = txtSunucu.Text.ToString();
            string kullanici = txtUser.Text.ToString();
            string sifre = txtPass.Text.ToString();
            string tablo = TxtTable.Text.ToString();

            string sConnectionString = "Provider=MSDAORA.1;Data Source=" + sunucu + ";User Id=" + kullanici + ";Password=" + sifre + ";Persist Security Info=False";
            //Integrated Security=no
           // OleDbConnection myConnection = new OleDbConnection(sConnectionString);
           // OleDbCommand myCommand = new OleDbCommand();//new OleDbCommand(mySelectQuery, myConnection);
           // MessageBox.Show("bağlantı ok ");
            try
            {
                createsqltable(sConnectionString, dtExcelRecords, tablo);
                MessageBox.Show("TABLO ok ");
                //using (var connection = new OracleConnection(connectionString))
                //{
                //    connection.Open();
                //    using (var bulkCopy = new OracleBulkCopy(connection, OracleBulkCopyOptions.UseInternalTransaction))
                //    {
                //        bulkCopy.DestinationTableName = dt.TableName;
                //        bulkCopy.WriteToServer(dt);
                //    }
                //} 

                //using (DbDataReader dr = myCommand.ExecuteReader())
                //{
                //    using (SqlBulkCopy bulkCopy =
                //               new SqlBulkCopy(sConnectionString))
                //    {
                //        bulkCopy.DestinationTableName = tablo;
                //        bulkCopy.WriteToServer(dr);


                //    }
                //    MessageBox.Show("Data aktarildi");
                //    //ExecuteSQLStr("exec CCOPS.dbo.PTS_SP_PREGRINE_FILE_FIN", sqlConnectionStringyeni);
                //    //MessageBox.Show("Prosedur tamamlandı. Mail bekleniyor..");
                //}



                //String mySelectQuery =    "SELECT * FROM TestTable where c1 LIKE ?";




                //myCommand.Parameters.Add("@p1", OleDbType.Char, 5).Value = "Test%";
                //myConnection.Open();
                //OleDbDataReader myReader = myCommand.ExecuteReader();
                //int RecordCount=0;
                //try
                //{
                //    while (myReader.Read())
                //    {
                //        RecordCount = RecordCount + 1;
                //    MessageBox.Show(myReader.GetString(0).ToString());
                //    }
                //    if (RecordCount == 0)
                //    {
                //    MessageBox.Show("No data returned");
                //    }
                //    else
                //    {
                //    MessageBox.Show("Number of records returned: " + RecordCount);
                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.ToString());
                //}
                //finally
                //{
                //    myReader.Close();
                //    myConnection.Close();
                //}
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }
    }




}
