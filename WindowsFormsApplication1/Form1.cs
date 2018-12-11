using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private MySqlConnection con;
        public Form1()
        {
            InitializeComponent();
            string conn = "Data Source=127.0.0.1;User ID=root;Password=root;DataBase=busmanager";
            con = new MySqlConnection(conn);
            con.Open();
            Form1_Resize(null, null);
            //使用List<>泛型集合填充DataGridView  
            List<Student> students = new List<Student>();
            Student hat = new Student("Hathaway", "12", "Male");
            Student peter = new Student("Peter", "14", "Male");
            Student dell = new Student("Dell", "16", "Male");
            Student anne = new Student("Anne", "19", "Female");
            students.Add(hat);
            students.Add(peter);
            students.Add(dell);
            students.Add(anne);
            this.dataGridView1.DataSource = students;  
           
        }
        public DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }


        private string search(string name,string uid) 
        {

            MySqlCommand cmd = new MySqlCommand("select count(*) from member_info where groupid=2 and  name='" + name + "' and uid='" + uid + "' and (email='' or email is null)", con);


            object res = cmd.ExecuteScalar();
            string result = res.ToString();
            //MessageBox.Show(result);
            return result;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[4].Value.ToString().Equals("1"))
                    {
                        MySqlCommand cmd = new MySqlCommand("Update member_info set email='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "@stu.edu.cn' where name='" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'", con);
                        cmd.ExecuteNonQuery();
                    }

                }
            }
            catch (Exception eee) { }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Height = this.ClientRectangle.Height - 140;
        }

    }
}
