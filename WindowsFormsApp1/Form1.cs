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
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString);
            
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

        private void add_student_button_Click(object sender, EventArgs e) //Кнопка "Добавить"
        {
            //Подключение БД
            SqlCommand command = new SqlCommand(
                $"INSERT INTO [KGEU_Diploma] (diploma_RN, studName, diplomaForm_SN, diploma_supplement_form_SN, diplomaIssue_Date, trainingDirection_code, trainingDirection_Name, assignedQualification_Name, honors, stateCommissionProtocol_Date, graduateExpulsionOrder_Date, diploma_status,admission_Year, graduation_Year, passport, student_signature, management_signature) VALUES (@diploma_RN, @studName, @diplomaForm_SN, @diploma_supplement_form_SN, @diplomaIssue_Date, @trainingDirection_code, @trainingDirection_Name, @assignedQualification_Name, @honors, @stateCommissionProtocol_Date, @graduateExpulsionOrder_Date, @diploma_status, @admission_Year, @graduation_Year, @passport, @student_signature, @management_signature)", sqlConnection);

            //Заполнение БД
            //string validformat = "dd-MM-yyyy";
            string DiplomaIssue_Date = diploma_issue_dateTimePickerAdd.Text;
            string Admission_Year = diploma_status_comboBoxAdd.Text;
            string Graduation_Year = graduation_Year_dateTimePickerAdd.Text;

            command.Parameters.AddWithValue("diploma_RN", diploma_RN_textBoxAdd.Text);
            command.Parameters.AddWithValue("studName", stud_name_textBoxAdd.Text);
            command.Parameters.AddWithValue("diplomaForm_SN", diplomaForm_SN_textBoxAdd.Text);
            command.Parameters.AddWithValue("diploma_supplement_form_SN", diploma_sup_form_SN_textBoxAdd.Text);
            command.Parameters.AddWithValue("diplomaIssue_Date", DiplomaIssue_Date);
            command.Parameters.AddWithValue("trainingDirection_code", traningDC_textBoxAdd.Text);
            command.Parameters.AddWithValue("trainingDirection_Name", traningDN_textBoxAdd.Text);
            command.Parameters.AddWithValue("assignedQualification_Name", assignedQualification_Name_textBoxAdd.Text);
            command.Parameters.AddWithValue("honors", honors_comboBoxAdd.Text);
            command.Parameters.AddWithValue("stateCommissionProtocol_Date", stateCommissionProtocol_Date_textBoxAdd.Text);
            command.Parameters.AddWithValue("graduateExpulsionOrder_Date", graduationExplusionOrder_Date_textBoxAdd.Text);
            command.Parameters.AddWithValue("diploma_status", diploma_status_comboBoxAdd.Text);
            command.Parameters.AddWithValue("admission_Year", Admission_Year);
            command.Parameters.AddWithValue("graduation_Year", Graduation_Year);
            command.Parameters.AddWithValue("passport", passport_textBoxAdd.Text);
            command.Parameters.AddWithValue("student_signature", student_signature_comboBoxAdd.Text);
            command.Parameters.AddWithValue("management_signature", managment_signature_comboBoxAdd.Text);

            //Уведомление о количестве заполненных строк
            MessageBox.Show(command.ExecuteNonQuery().ToString());

            diploma_RN_textBoxAdd.Text = null;
            stud_name_textBoxAdd.Text = null;
            diplomaForm_SN_textBoxAdd.Text = null;
            diploma_sup_form_SN_textBoxAdd.Text = null;
            traningDC_textBoxAdd.Text = null;
            traningDN_textBoxAdd.Text = null;
            assignedQualification_Name_textBoxAdd.Text = null;
            stateCommissionProtocol_Date_textBoxAdd.Text = null;
            graduationExplusionOrder_Date_textBoxAdd.Text = null;
            diploma_status_comboBoxAdd.Text = null;
            passport_textBoxAdd.Text = null;
            student_signature_comboBoxAdd.Text = null;
            managment_signature_comboBoxAdd.Text = null;
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
            string[,] data = new string[17, 1];

            for (int i=0; i<dataGridView2.RowCount-1; i++)
            {
                for (int j=0; j<dataGridView2.ColumnCount; j++)
                {
                    data[j, i] = dataGridView2[j, i].Value.ToString();
                    //MessageBox.Show(data[j, i].ToString(), j.ToString());
                    ws.Cells[i+7 , j +5] = dataGridView2[j, i].Value.ToString();
                }
            }

            // с какой строки начинаем вставлять данные из dgv
            //int iRowCount = 3;

            //for (int i = 0; i < ws.Rows.Count; i++)
            //{
                //excelApp.Cells[iRowCount, 1] = ws.Rows[i].Cells[0].Value.ToString();
                //excelApp.Cells[iRowCount, 2] = ws.Rows[i].Cells[1].Value.ToString();
                //excelApp.Cells[iRowCount, 3] = ws.Rows[i].Cells[2].Value.ToString();
                //excelApp.Cells[iRowCount, 4] = ws.Rows[i].Cells[3].Value.ToString();

                //iRowCount++;

                // Добавляем строчку ниже
                //var cellsDRnr = ws.get_Range("A" + iRowCount, "A" + iRowCount);
                //cellsDRnr.EntireRow.Insert(-4121, m_objOpt);
                excelApp.Visible = true;
            //}
        }

        private void update_button_Click(object sender, EventArgs e)
        {
            //Обновление данных в "Поиске" (dataGridView2)
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString);

            sqlConnection.Open();

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM KGEU_Diploma", sqlConnection);

            DataSet db = new DataSet();

            dataAdapter.Fill(db);

            dataGridView2.DataSource = db.Tables[0];
        }

        private void WritingToTheDataBase_button_Click(object sender, EventArgs e)
        {

        }
    }
}
