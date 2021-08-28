using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Common;
namespace ExcelToSql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //SqlConnection sqlConnectionString = new SqlConnection(@"data source.\SQLEXPRESS;database=excelToTraining;integrated security=true");
        SqlConnection sqlConnectionString = new SqlConnection(@"data source=.\SQLEXPRESS;database=excelToTraining;integrated security=true");
        string m;
        private void button1_Click(object sender, EventArgs e)
        {
            
            DialogResult g= openFileDialog1.ShowDialog();
            if (DialogResult.OK==g)
            {
                 m = openFileDialog1.FileName;
                MessageBox.Show(m);
            }
            #region tries
            //SqlBulkCopy b = new SqlBulkCopy(con);
            //con.Open();
            //b.DestinationTableName = "excelToTraining.dbo.Query3";
            //try
            //{
            //    // Write from the source to the destination.
            //    b.WriteToServer(reader);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
            //finally
            //{
            //    // Close the SqlDataReader. The SqlBulkCopy
            //    // object is automatically closed at the end
            //    // of the using block.
            //    reader.Close();
            //}
            #endregion
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+m+" ;Extended Properties=Excel 8.0";
            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand command = new OleDbCommand("Select * FROM [Sheet1$]", connection);
                connection.Open();

                using (DbDataReader reader = command.ExecuteReader())
                {
                    sqlConnectionString.Open(); 
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnectionString))
                    {
                        bulkCopy.DestinationTableName = "Book";
                        bulkCopy.WriteToServer(reader);
                        MessageBox.Show("Transfer data is Succeeded");
                    }

                }
                //DataSet ds = new DataSet();
                //OleDbDataAdapter adap = new OleDbDataAdapter(command);
                //adap.Fill(ds);
                //dataGridView1.DataSource = ds;
                ////dataGridView1.
            }
        }
    }
}
