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
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
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

            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
            else
                MessageBox.Show("Данные добавлены", "Внимание");


            dbConnection.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку", "Внимание");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;


            string Номер_полиса = dataGridView1.Rows[index].Cells[0].Value.ToString();

            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pac.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "DELETE FROM pac WHERE Номер_полиса =" + Номер_полиса;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
            else
            {
                MessageBox.Show("Данные удалены", "Внимание");
                dataGridView1.Rows.RemoveAt(index);
            }


            dbConnection.Close();
        }

        private void button_update_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку", "Внимание");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            if (dataGridView1.SelectedRows[index].Cells[0].Value == null ||
                dataGridView1.SelectedRows[index].Cells[1].Value == null ||
                dataGridView1.SelectedRows[index].Cells[2].Value == null )
               

            {
                MessageBox.Show("Не все данные введены", "Внимание");
                return;
            }
            string Номер_полиса = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string ФИО_пациента = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string Врач = dataGridView1.Rows[index].Cells[2].Value.ToString();
          
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pac.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "UPDATE pac SET Номер_полиса = " + Номер_полиса + ",ФИО_пациента = '" + ФИО_пациента + "',Врач= '" + Врач + "' WHERE Номер_полиса =" + Номер_полиса;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
            else
            {
                MessageBox.Show("Данные изменены", "Внимание");
            }




            dbConnection.Close();
        }

        private void button_download_Click(object sender, EventArgs e)
        {
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pac.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "SELECT * FROM pac";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены", "Ошибка");
            }
            else
            {
                while (dbReader.Read())
                {
                    dataGridView1.Rows.Add(dbReader["Номер_полиса"], dbReader["ФИО_пациента"], dbReader["Врач"]);
                }
            }
            dbReader.Close();
            dbConnection.Close();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Hide();
        }
    }
}
