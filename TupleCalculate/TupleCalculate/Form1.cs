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
using System.Data.Odbc;

namespace TupleCalculate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        
          
         
        private void button1_Click(object sender, EventArgs e)
        {
            //OleDbConnection cnn = new OleDbConnection();
            //string ss = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\我的文件\Desktop\TupleCalculate\新增資料夾\尺寸資料.xlsx;";
            //ss += "Extended Properties=\"Excel 12.0;HDR=YES;\"";
            //cnn.ConnectionString = ss;
            //try
            //{
            //    cnn.Open();
            //    MessageBox.Show("連接 Excel 365 資料成功！");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //    MessageBox.Show("連接Excel資料失敗！");
            //}
            //finally
            //{
            //    cnn.Close();
            //}
            /*
           // DSN 的連接法
            OdbcConnection cnn = new OdbcConnection();
            cnn.ConnectionString = @"DSN=MyDSN;UID=;PWD=;";
            try
            {
                cnn.Open();
                MessageBox.Show("連接 Excel 365 資料成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show("連接Excel資料失敗！");
            }
            finally
            {
                cnn.Close();
            }
            */
            OleDbConnection cnn = new OleDbConnection();
            string ss = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Ryoma\NET\Web\TubeToExcel\尺寸資料.xlsx;";
            ss += "Extended Properties=\"Excel 12.0;HDR=YES;\"";
            cnn.ConnectionString = ss;
            using (OleDbCommand cmd = new OleDbCommand())
            {
                cmd.CommandText = "CREATE TABLE 管料(尺寸 FLOAT,數量 INT)";
                cmd.Connection = cnn;
                try
                {
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "INSERT INTO 管料(尺寸,數量) VALUES( textBox1.Text ,textBox2.Text)";
                    cmd.ExecuteReader();
                    MessageBox.Show("連接 Excel 管料sheet表 資料成功！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    
                }
                finally
                {
                    cmd.Dispose();
                    cnn.Close();
                }
            }
               

        }

        private void button3_Click(object sender, EventArgs e)
        {
        }
    }
}
