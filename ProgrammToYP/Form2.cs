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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.Common;
using System.Xml.Linq;

namespace ProgrammToYP
{
    public partial class Form2 : Form
    {
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataReader dr;
        public Form2()
        {
            InitializeComponent();
        }
      
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            frm1.Show();
            this.Hide();

        }

        private void button_update1_Click(object sender, EventArgs e)
        {
            
                OleDbConnection con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Авторизация.mdb");
                OleDbDataAdapter ada = new OleDbDataAdapter("Select count(*) From Login where Логин = '" + txtUser.Text + "'and Пароль = " + txtPass.Text + "", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                    {
                        Hide();
                        Form3 fGl = new Form3();
                        fGl.Show();
                    }
                else
                    {
                        MessageBox.Show("Неправильный логин или пароль!!!");
                    }
            
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
        }
    }
}
