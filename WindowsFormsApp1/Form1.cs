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
using ExcelDataReader;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


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
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["studName"].ConnectionString);

            sqlConnection.Open();

            //Проверка подключения БД
            if (sqlConnection.State == ConnectionState.Open)
            {
                MessageBox.Show("Подключение установлено");
            }
            else { MessageBox.Show("Ошибка в подключении БД"); }

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM KGEU_Diploma", sqlConnection);

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

            //string validformat = "dd-MM-yyyy";

            command.Parameters.AddWithValue("Name", stud_name_textBoxAdd.Text);
            command.Parameters.AddWithValue("Birthday", $"{diploma_issue_dateTimePickerAdd.Value.Day}/{diploma_issue_dateTimePickerAdd.Value.Month}/{diploma_issue_dateTimePickerAdd.Value.Year}");
            command.Parameters.AddWithValue("Graduation", diplomaForm_SN_textBoxAdd.Text);

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

            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM KGEU_Diploma",
                sqlConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            dataGridView4.DataSource = dataSet.Tables[0];
        }

        //Фильтрация
        //Оброботчики события изменения текста в полях.(При изменении текста в любом из полей запускается фильтрация по заполненным полям)
        //private void textBox5_TextChanged(object sender, EventArgs e)
        //{
          //  (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"Convert(id, 'System.String') LIKE '%{traningDN_textBoxS.Text}%' AND Name LIKE '%{passport_textBoxS.Text}%' AND Convert(id, 'System.String') LIKE '%{traningDN_textBoxS.Text}%' AND Convert(Birthday, 'System.String') LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND Convert(Graduation, 'System.String') LIKE '%{traningDC_textBoxS.Text}%'");
        //}

        private void diploma_RN_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void stud_name_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void traningDC_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void traningDN_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void assignedQualification_Name_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void honors_comboBoxS_SelectionChangeCommitted(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        private void passport_textBoxS_TextChanged(object sender, EventArgs e)
        {
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format($"diploma_RN LIKE '%{diploma_RN_textBoxS.Text}%' AND studName LIKE '%{stud_name_textBoxS.Text}%' AND trainingDirection_code LIKE '%{traningDC_textBoxS.Text}%' AND trainingDirection_Name LIKE '%{traningDN_textBoxS.Text}%' AND assignedQualification_Name LIKE '%{assignedQualification_Name_textBoxS.Text}%' AND passport LIKE '%{passport_textBoxS.Text}%'");
        }

        //Чтение файла Excel
        private string excelFileName = string.Empty;

        private DataTableCollection excelTableCollection = null;

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dialogResult = openFileDialog1.ShowDialog();

                if (dialogResult == DialogResult.OK)
                {
                    excelFileName = openFileDialog1.FileName;
                    Text = excelFileName;
                    OpenExcelFile(excelFileName);
                }
                else
                {
                    throw new Exception("Файл не выбран");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
         
        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            DataSet excelDB = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            excelTableCollection = excelDB.Tables;
            toolStripComboBox1.Items.Clear();

            foreach (DataTable table in excelTableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = excelTableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            dataGridView3.DataSource = table;
        }

        //Создание Excel файла и заполнение через dataGridView2 (вкладка поиск)
        private void Export_button_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            Excel.Worksheet ws = (Excel.Worksheet)excelApp.ActiveSheet;

            for (int i=0; i<dataGridView2.RowCount-1; i++)
            {
                for (int j=0; j<dataGridView2.ColumnCount; j++)
                {
                    ws.Cells[i + 1, j + 1] = dataGridView2[j, i].Value.ToString();
                }
            }

            excelApp.Visible = true;
        }
    }
}
