using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Globalization;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) //Обработчик события загрузки приложение
        {
            this.Text = "Выпускники КГЭУ";

            //Подключение БД
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Student"].ConnectionString);

            sqlConnection.Open();

            //Проверка подключения БД
            if (sqlConnection.State == ConnectionState.Open)
            {
                MessageBox.Show("Подключение установлено");
            }
            else { MessageBox.Show("Ошибка в подключении БД"); }

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Students", sqlConnection);

            DataSet db = new DataSet();

            dataAdapter.Fill(db);

            dataGridView2.DataSource = db.Tables[0];
        }

        private void button1_Click(object sender, EventArgs e) //Кнопка "Добавить"
        {
            //Подключение БД
            SqlCommand command = new SqlCommand(
                $"INSERT INTO [Students] (Name, Birthday, Graduation) VALUES (@Name, @Birthday, @Graduation)", sqlConnection);

            //Заполнение БД
           
            string validformat = "dd/MM/yyyy";

            CultureInfo provider = new CultureInfo("ru_RU");
            DateTime date = DateTime.ParseExact(textBox2.Text, validformat, provider);

            command.Parameters.AddWithValue("Name", textBox1.Text);
            command.Parameters.AddWithValue("Birthday", $"{date.Day}/{date.Month}/{date.Year}");
            command.Parameters.AddWithValue("Graduation", textBox3.Text);

            //Уведомление о количестве заполненных строк
            MessageBox.Show(command.ExecuteNonQuery().ToString());
        }

        private void button2_Click(object sender, EventArgs e) //Кнопка "Поиск"
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                textBox4.Text,
                sqlConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            dataGridView1.DataSource = dataSet.Tables[0];
        }

        private void button3_Click(object sender, EventArgs e) //Кнопка "Поиск" (2)
        {
            listView1.Items.Clear();

            SqlDataReader dataReader = null;

            try 
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Students",
                    sqlConnection);

                dataReader = sqlCommand.ExecuteReader();

                ListViewItem item = null;

                while (dataReader.Read())
                {
                    item = new ListViewItem(new string[] { Convert.ToString(dataReader["ID"]), 
                        Convert.ToString(dataReader["Name"]), 
                        Convert.ToString(dataReader["Birthday"]), 
                        Convert.ToString(dataReader["Graduation"]) });

                    listView1.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (dataReader != null && !dataReader.IsClosed)
                {
                    dataReader.Close();
                }
            }
        }

        //Фильтрация
        //Оброботчики события изменения текста в полях.(При изменении текста в любом из полей запускается фильтрация по заполненным полям)
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Name LIKE '%{textBox6.Text}%' AND Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{textBox7.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{textBox8.Text}%'");
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Name LIKE '%{textBox6.Text}%' AND Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{textBox7.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{textBox8.Text}%'");
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Name LIKE '%{textBox6.Text}%' AND Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{textBox7.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{textBox8.Text}%'");
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Name LIKE '%{textBox6.Text}%' AND Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{textBox7.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{textBox8.Text}%'");
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Name LIKE '%{textBox6.Text}%' AND Convert(id, 'System.String') LIKE '%{textBox5.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{textBox7.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{textBox8.Text}%'");
        }

    }
}
