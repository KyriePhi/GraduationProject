using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DetayliSinavAnalizi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            exam.Filter = "Excel Files|*.xlsx;";
            if (exam.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = exam.FileName;
                string connectionadress = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties='Excel 12.0;IMEX=1;'";
                OleDbConnection connectionn = new OleDbConnection(connectionadress);
                OleDbCommand cmd = new OleDbCommand("Select * From [" + "Sayfa1" + "$]", connectionn);
                connectionn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable data = new DataTable();
                da.Fill(data);
                dataGridView1.DataSource = data;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            student.Filter = "Excel Files|*.xlsx;";
            if (student.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = student.FileName;
            }
        }
    }
}
