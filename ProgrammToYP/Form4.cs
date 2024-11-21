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

namespace ProgrammToYP
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            frm1.Show();
            this.Hide();
        }
        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.LightGray;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromKnownColor(KnownColor.Control);
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
                dataGridView1.SelectedRows[index].Cells[2].Value == null ||
                dataGridView1.SelectedRows[index].Cells[3].Value == null ||
                dataGridView1.SelectedRows[index].Cells[4].Value == null ||
                dataGridView1.SelectedRows[index].Cells[5].Value == null)

            {
                MessageBox.Show("Не все данные введены", "Внимание");
                return;
            }
            string Номер_полиса = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string ФИО_пациента = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string Спец_Заболевание = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string Степень_тяжести = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string Врач = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string Кабинет = dataGridView1.Rows[index].Cells[5].Value.ToString();


            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pol2.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "UPDATE pol2 SET Номер_полиса = " + Номер_полиса + ",ФИО_пациента = '" + ФИО_пациента + "',Заболевание = '" + Спец_Заболевание + "',Степень_тяжести = '" + Степень_тяжести + "',Врач= '" + Врач + "' ,Кабинет=" + Кабинет + " WHERE Номер_полиса =" + Номер_полиса;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
            else
            {
                MessageBox.Show("Данные изменены", "Внимание");
            }




            dbConnection.Close();
        }

        private void button_add_Click(object sender, EventArgs e)
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
                dataGridView1.SelectedRows[index].Cells[2].Value == null ||
                dataGridView1.SelectedRows[index].Cells[3].Value == null ||
                dataGridView1.SelectedRows[index].Cells[4].Value == null ||
                dataGridView1.SelectedRows[index].Cells[5].Value == null)


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
            string Спец_Заболевание = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string Степень_тяжести = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string Врач = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string Кабинет = dataGridView1.Rows[index].Cells[5].Value.ToString();

            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pol2.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "INSERT INTO pol2 VALUES (" + Номер_полиса + ",'" + ФИО_пациента + "','" + Спец_Заболевание + "','" + Степень_тяжести + "','" + Врач + "'," + Кабинет + ")";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка");
            else
                MessageBox.Show("Данные добавлены", "Внимание");


            dbConnection.Close();
        }

        private void button_delete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку", "Внимание");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

           
            string Номер_полиса = dataGridView1.Rows[index].Cells[0].Value.ToString();

            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pol2.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "DELETE FROM pol2 WHERE Номер_полиса =" + Номер_полиса;
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

        private void button_download_Click(object sender, EventArgs e)
        {
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source = pol2.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);
            dbConnection.Open();
            string query = "SELECT * FROM pol2";
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
                    dataGridView1.Rows.Add(dbReader["Номер_полиса"], dbReader["ФИО_пациента"], dbReader["Спец_Заболевание"], dbReader["Степень_тяжести"], dbReader["Врач"], dbReader["Кабинет"]);
                }
            }
            dbReader.Close();
            dbConnection.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form6 frm6 = new Form6();
            frm6.Show();
            this.Hide();
        }
    }
}
