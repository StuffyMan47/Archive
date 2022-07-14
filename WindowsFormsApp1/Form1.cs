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
            try
            {
                sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString);
                sqlConnection.Open();
            }
            catch
            {
                //Проверка подключения БД
                if (sqlConnection.State != ConnectionState.Open)
                {
                    MessageBox.Show("Ошибка в подключении БД");
                }
            }
            

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
            string Admission_Year = admission_Year_dateTimePickerAdd.Text;
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
            ws.Name = "Отчёт";

            //Разметка Ecxel документа в оглавлении документа
            Excel.Range rangeHeading = ws.get_Range("A1", "G1");
            rangeHeading.Cells.Font.Name = "Times New Roman";
            rangeHeading.Cells.Font.Size = 14;
            rangeHeading.Merge(Type.Missing);
            rangeHeading.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rangeHeading.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rangeHeading.Value = "Подтверждение о наличии диплома о высшем образовании";

            //Разметка Ecxel документа в оглавлении таблицы
            Excel.Range rangeMain = ws.get_Range("A5", "F5");
            rangeMain.Cells.Font.Name = "Times New Roman";
            rangeMain.Cells.Font.Size = 10;
            rangeMain.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rangeMain.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            Excel.Range rangeColorMain = rangeMain;
            rangeColorMain.Borders.Color = ColorTranslator.ToOle(Color.Black);
            Excel.Range rowHeightMain = rangeMain;
            rowHeightMain.EntireRow.RowHeight = 70;
            ws.Cells[5, 1] = "Фамилия, имя, отчество";
            ws.Cells[5, 2] = "Год постуаления";
            ws.Cells[5, 3] = " Год выпуска ";
            ws.Cells[5, 4] = "Дата и номер протокола \nгосударственной комиссии";
            ws.Cells[5, 5] = "Код направления";
            ws.Cells[5, 6] = "Наименование направления подготовки";

            //Разметка Ecxel документа в заполняемой зоне (таблица)
            Excel.Range rangeTable = ws.get_Range("A6", "F6");
            rangeTable.Cells.Font.Name = "Times New Roman";
            rangeTable.Cells.Font.Size = 10;
            rangeTable.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rangeTable.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            Excel.Range rangeColorTable = rangeTable;
            rangeColorTable.Borders.Color = ColorTranslator.ToOle(Color.Black);
            Excel.Range rowHeightTable = rangeTable;
            rowHeightTable.EntireRow.RowHeight = 70;
            rangeTable.EntireColumn.AutoFit();

            //Заполнение полей таблицы данными из БД
            ws.Cells[6, 1] = dataGridView2[1, 0].Value.ToString();
            ws.Cells[6, 2] = dataGridView2[12, 0].Value.ToString();
            ws.Cells[6, 3] = dataGridView2[13, 0].Value.ToString();
            ws.Cells[6, 4] = dataGridView2[9, 0].Value.ToString();
            ws.Cells[6, 5] = dataGridView2[5, 0].Value.ToString();
            ws.Cells[6, 6] = dataGridView2[6, 0].Value.ToString();

            excelApp.Visible = true;
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
