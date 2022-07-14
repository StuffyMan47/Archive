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

            sqlConnection.Close();
        }

        private void add_student_button_Click(object sender, EventArgs e) //Кнопка "Добавить"
        {
            sqlConnection.Open();
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

            sqlConnection.Close();
        }

        private void button2_Click(object sender, EventArgs e) //Кнопка "Поиск"
        {
            sqlConnection.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                textBox4.Text,
                sqlConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            dataGridView1.DataSource = dataSet.Tables[0];

            sqlConnection.Close();
        }

        private void button3_Click(object sender, EventArgs e) //Кнопка "Поиск" (2)
        {
            sqlConnection.Open();

            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM KGEU_Diploma",
                sqlConnection);

            DataSet dataSet = new DataSet();

            dataAdapter.Fill(dataSet);

            dataGridView4.DataSource = dataSet.Tables[0];

            sqlConnection.Close();
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
            sqlConnection.Open();

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

            if (dataGridView3.Rows.Count > 2)
                WritingToTheDataBase_button.Enabled = true;


            sqlConnection.Close();
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

            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    data[j, i] = dataGridView2[j, i].Value.ToString();
                    //MessageBox.Show(data[j, i].ToString(), j.ToString());
                    ws.Cells[i + 7, j + 5] = dataGridView2[j, i].Value.ToString();
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
            sqlConnection.Open();
            //Обновление данных в "Поиске" (dataGridView2)
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString);

            sqlConnection.Open();

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM KGEU_Diploma", sqlConnection);

            DataSet db = new DataSet();

            dataAdapter.Fill(db);

            dataGridView2.DataSource = db.Tables[0];

            sqlConnection.Close();
        }

        private void WritingToTheDataBase_button_Click(object sender, EventArgs e)
        {
            //Заполнение БД
            //string validformat = "dd-MM-yyyy";
            sqlConnection.Open();

            SqlCommand command = new SqlCommand(
              $"INSERT INTO [KGEU_Diploma] (diploma_RN, studName, diplomaForm_SN, diploma_supplement_form_SN, diplomaIssue_Date, trainingDirection_code, trainingDirection_Name, assignedQualification_Name, honors, stateCommissionProtocol_Date, graduateExpulsionOrder_Date, diploma_status,admission_Year, graduation_Year, passport, student_signature, management_signature) VALUES (@diploma_RN, @studName, @diplomaForm_SN, @diploma_supplement_form_SN, @diplomaIssue_Date, @trainingDirection_code, @trainingDirection_Name, @assignedQualification_Name, @honors, @stateCommissionProtocol_Date, @graduateExpulsionOrder_Date, @diploma_status, @admission_Year, @graduation_Year, @passport, @student_signature, @management_signature)", sqlConnection);

            command.Parameters.Add("@diploma_RN", SqlDbType.NVarChar);
            command.Parameters.Add("@studName", SqlDbType.NVarChar);
            command.Parameters.Add("@diplomaForm_SN", SqlDbType.NChar);
            command.Parameters.Add("@diploma_supplement_form_SN", SqlDbType.NChar);
            command.Parameters.Add("@diplomaIssue_Date", SqlDbType.NVarChar);
            command.Parameters.Add("@trainingDirection_code", SqlDbType.NChar);
            command.Parameters.Add("@trainingDirection_Name", SqlDbType.NVarChar);
            command.Parameters.Add("@assignedQualification_Name", SqlDbType.NVarChar);
            command.Parameters.Add("@honors", SqlDbType.NVarChar);
            command.Parameters.Add("@stateCommissionProtocol_Date", SqlDbType.NVarChar);
            command.Parameters.Add("@graduateExpulsionOrder_Date", SqlDbType.NVarChar);
            command.Parameters.Add("@diploma_status", SqlDbType.NVarChar);
            command.Parameters.Add("@admission_Year", SqlDbType.NVarChar);
            command.Parameters.Add("@graduation_Year", SqlDbType.NVarChar);
            command.Parameters.Add("@passport", SqlDbType.NVarChar);
            command.Parameters.Add("@student_signature", SqlDbType.NVarChar);
            command.Parameters.Add("@management_signature", SqlDbType.NVarChar);

            int a = 0;
            try
            {
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    a = i;
                    command.Parameters["@diploma_RN"].Value = dataGridView3["diploma_RN", i].Value;
                    command.Parameters["@studName"].Value = dataGridView3["studName", i].Value;
                    command.Parameters["@diplomaForm_SN"].Value = dataGridView3["diplomaForm_SN", i].Value;
                    command.Parameters["@diploma_supplement_form_SN"].Value = dataGridView3["diploma_supplement_form_SN", i].Value;
                    command.Parameters["@diplomaIssue_Date"].Value = dataGridView3["diplomaIssue_Date", i].Value;
                    command.Parameters["@trainingDirection_code"].Value = dataGridView3["trainingDirection_code", i].Value;
                    command.Parameters["@trainingDirection_Name"].Value = dataGridView3["trainingDirection_Name", i].Value;
                    command.Parameters["@assignedQualification_Name"].Value = dataGridView3["assignedQualification_Name", i].Value;
                    command.Parameters["@honors"].Value = dataGridView3["honors", i].Value;
                    command.Parameters["@stateCommissionProtocol_Date"].Value = dataGridView3["stateCommissionProtocol_Date", i].Value;
                    command.Parameters["@graduateExpulsionOrder_Date"].Value = dataGridView3["graduateExpulsionOrder_Date", i].Value;
                    command.Parameters["@diploma_status"].Value = dataGridView3["diploma_status", i].Value;
                    command.Parameters["@admission_Year"].Value = dataGridView3["admission_Year", i].Value;
                    command.Parameters["@graduation_Year"].Value = dataGridView3["graduation_Year", i].Value;
                    command.Parameters["@passport"].Value = dataGridView3["passport", i].Value;
                    command.Parameters["@student_signature"].Value = dataGridView3["student_signature", i].Value;
                    command.Parameters["@management_signature"].Value = dataGridView3["management_signature", i].Value;
                    command.ExecuteNonQuery();
                }

                MessageBox.Show(command.ExecuteNonQuery().ToString());
            }
            catch
            {
                object b = dataGridView3["diploma_RN", a].Value;
                MessageBox.Show($"Человек с номером {b} уже существует в БД"); ;
            }
            finally
            {
                sqlConnection.Close();
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application crFile = new Excel.Application();
            crFile.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)crFile.ActiveSheet;

            string[,] nColimns = new string[,] { { "", "diploma_RN", "studName", "diplomaForm_SN", "diploma_supplement_form_SN", "diplomaIssue_Date", "trainingDirection_code", "trainingDirection_Name", "assignedQualification_Name", "honors", "stateCommissionProtocol_Date", "graduateExpulsionOrder_Date", "diploma_status", "admission_Year", "graduation_Year", "passport", "student_signature", "management_signature" } };

            for (int i = 1; i < nColimns.Length - 1; i++)
                sheet.Cells[1, i] = nColimns[0, i];

            crFile.Visible = true;

        }

        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            sqlConnection.Close();
        }
    }
}
