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

namespace ProgrammToYP
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку", "Внимание");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;


            try
            {
                if (dataGridView1.SelectedRows[index].Cells[0].Value == null ||
                dataGridView1.SelectedRows[index].Cells[1].Value == null ||
                dataGridView1.SelectedRows[index].Cells[2].Value == null)


                {
                    MessageBox.Show("Не все данные введены", "Внимание");
                    return;
                }
            }
            catch
            {

            }

            string Номер_полиса = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string ФИО_пациента = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string Врач = dataGridView1.Rows[index].Cells[2].Value.ToString();

            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pac.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "INSERT INTO pac VALUES (" + Номер_полиса + ",'" + ФИО_пациента + "','" + Врач + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            try
            {
                if (dbCommand.ExecuteNonQuery() != 1)
                    MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
                else
                    MessageBox.Show("Данные добавлены", "Внимание");
                MessageBox.Show("Вы записаны.");
            }
            catch
            {
                MessageBox.Show("Данные совпадают,такой пациент существует.");
            }
            
            dbConnection.Close();

            Form1 frm1 = new Form1();
            frm1.Show();
            this.Hide();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
        }
    }
}
