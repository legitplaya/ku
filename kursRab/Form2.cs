using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;
using MaterialSkin;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;


namespace kursRab
{
    public partial class Form2 : MaterialForm
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet2.logpass". При необходимости она может быть перемещена или удалена.
            this.logpassTableAdapter.Fill(this.ip521_Sofronov_KRDataSet2.logpass);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.Фильмы_для_лиц_старше_17_лет". При необходимости она может быть перемещена или удалена.
            this.фильмы_для_лиц_старше_17_летTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.Фильмы_для_лиц_старше_17_лет);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.самый_первый_купленный_билет". При необходимости она может быть перемещена или удалена.
            this.самый_первый_купленный_билетTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.самый_первый_купленный_билет);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.Посетители_старше_18_лет". При необходимости она может быть перемещена или удалена.
            this.посетители_старше_18_летTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.Посетители_старше_18_лет);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.Минимальная_стоимость_билета". При необходимости она может быть перемещена или удалена.
            this.минимальная_стоимость_билетаTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.Минимальная_стоимость_билета);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.Максимальная_стоимость_билета". При необходимости она может быть перемещена или удалена.
            this.максимальная_стоимость_билетаTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.Максимальная_стоимость_билета);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1._Кол_во_сеансов_в_1_зале_за_все_время". При необходимости она может быть перемещена или удалена.
            this.кол_во_сеансов_в_1_зале_за_все_времяTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1._Кол_во_сеансов_в_1_зале_за_все_время);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1._Кол_во_свободных_залов". При необходимости она может быть перемещена или удалена.
            this.кол_во_свободных_заловTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1._Кол_во_свободных_залов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1._Кол_во_купленных_билетов_дороже_500_рублей". При необходимости она может быть перемещена или удалена.
            this.кол_во_купленных_билетов_дороже_500_рублейTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1._Кол_во_купленных_билетов_дороже_500_рублей);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1._Кол_во_занятых_залов". При необходимости она может быть перемещена или удалена.
            this.кол_во_занятых_заловTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1._Кол_во_занятых_залов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet1.Билеты_купленные_за_20_лет". При необходимости она может быть перемещена или удалена.
            this.билеты_купленные_за_20_летTableAdapter.Fill(this.ip521_Sofronov_KRDataSet1.Билеты_купленные_за_20_лет);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet.tickets". При необходимости она может быть перемещена или удалена.
            this.ticketsTableAdapter.Fill(this.ip521_Sofronov_KRDataSet.tickets);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet.rooms". При необходимости она может быть перемещена или удалена.
            this.roomsTableAdapter.Fill(this.ip521_Sofronov_KRDataSet.rooms);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet.customers". При необходимости она может быть перемещена или удалена.
            this.customersTableAdapter.Fill(this.ip521_Sofronov_KRDataSet.customers);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "ip521_Sofronov_KRDataSet.films". При необходимости она может быть перемещена или удалена.
            this.filmsTableAdapter.Fill(this.ip521_Sofronov_KRDataSet.films);
            
           
             /*
            PC1 = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\kosty\\source\\repos\\PC1.mdf;Integrated Security=True;Connect Timeout=30");
            PC1.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Computers", PC1);
            DataSet db = new DataSet();
            dataAdapter.Fill(db);
            dataGridView1.DataSource = db.Tables[0];*/

            // Заполняем ComboBox названиями столбцов ----------------НАСТРОИТЬ ДЛЯ МАТЕРИАЛ------------
            /*
            foreach (DataColumn column in db.Tables[0].Columns)
            {
                materialComboBox1.Items.Add(column.ColumnName);
            }*/
            

            if (Form1.User == "user")
            {
                //tabControl1.Visible = false;
                //tabControl2.Visible = true;
                

                dataGridView1.ReadOnly = true;
                dataGridView2.ReadOnly = true;
                dataGridView3.ReadOnly = true;
                dataGridView4.ReadOnly = true;

                dataGridView1.AllowUserToAddRows = false;
                dataGridView2.AllowUserToAddRows = false;
                dataGridView3.AllowUserToAddRows = false;
                dataGridView4.AllowUserToAddRows = false;

                dataGridView2.Visible = false;
                dataGridView4.Visible = false;

                
                /*
                Size = new Size(810, 789);
                MaximumSize = new Size(810, 789);
                MinimumSize = new Size(810, 789);*/
            }
            else
             if (Form1.User == "admin")
            {

            }

            
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void таблицыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl2.Visible = false;
        }

        private void представленияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = false;
            tabControl2.Visible = true;
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            filmsBindingSource.EndEdit();
            filmsTableAdapter.Update(ip521_Sofronov_KRDataSet.films);
            
            customersBindingSource.EndEdit();
            customersTableAdapter.Update(ip521_Sofronov_KRDataSet.customers);

            roomsBindingSource.EndEdit();
            roomsTableAdapter.Update(ip521_Sofronov_KRDataSet.rooms);

            ticketsBindingSource.EndEdit();
            ticketsTableAdapter.Update(ip521_Sofronov_KRDataSet.tickets);
            
        }

        private void фильтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void materialMaskedTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void materialComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        

        private void materialButton1_Click_1(object sender, EventArgs e)
        {
            filmsBindingSource.Filter = string.Empty;
            roomsBindingSource.Filter = string.Empty;
            ticketsBindingSource.Filter = string.Empty;
            customersBindingSource.Filter = string.Empty;

            customersBindingSource.RemoveFilter();
            ticketsBindingSource.RemoveFilter();
            roomsBindingSource.RemoveFilter();
            filmsBindingSource.RemoveFilter();

            materialTextBox3.Clear();
            materialTextBox5.Clear();
            
        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
            string columnName = materialTextBox3.Text;
            string filterValue = materialTextBox5.Text;
            string a = materialComboBox2.Text.ToString();
            string secondColumnName = materialTextBox4.Text;
            string secondFilterValue = materialTextBox2.Text;

            bool useOrCondition = materialSwitch1.Checked;

            if (materialSwitch5.Checked == false)
            {
                switch (a)
                {
                    {
					case "": filmsBindingSource.Filter = materialTextBox3.Text + " = ('" + materialTextBox5.Text + "')"; break;
					case "Товары": ticketsBindingSource.Filter = materialTextBox3.Text + " = ('" + materialTextBox5.Text + "')"; break;
					case "Поставки": customersBindingSource.Filter = materialTextBox3.Text + " = ('" + materialTextBox5.Text + "')"; break;
					case "Склады": roomsBindingSource.Filter = materialTextBox3.Text + " = ('" + materialTextBox5.Text + "')"; break;


                    default: MessageBox.Show("Таблица не найдена", "Ошибка"); break;
                }
            }
            

        }
        
        
        string filepath = "";
        private void exportToExcel(DataGridView dataGridView, string filepath)
        {
            try
            {
                if (materialComboBox1.SelectedItem != null)
                {
                    string tableName = materialComboBox1.SelectedItem.ToString();

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(tableName);

                        for (int column = 0; column < dataGridView.ColumnCount; column++)
                        {
                            worksheet.Cells[1, column + 1].Value = dataGridView.Columns[column].HeaderText;
                        }

                        for (int row = 0; row < dataGridView.Rows.Count; row++)
                        {
                            for (int column = 0; column < dataGridView.ColumnCount; column++)
                            {
                                worksheet.Cells[row + 2, column + 1].Value = dataGridView.Rows[row].Cells[column].Value;
                            }
                        }
                        using (ExcelRange range = worksheet.Cells[1, 1, dataGridView.Rows.Count + 1, dataGridView.ColumnCount])
                        {
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range.AutoFitColumns();
                        }
                        FileInfo excelFile = new FileInfo(filepath);
                        excelPackage.SaveAs(excelFile);
                    }
                    MessageBox.Show($"Таблица {tableName} успешно экспортирована!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    MessageBox.Show("Выберите таблицу для экспорта.", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                MessageBox.Show("При экспорте таблицы произошла ошибка", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string GetFilterCondition(string columnName, string filterValue, bool isSecondColumn = false, bool useOr = false, string secondColumnName = null, string secondFilterValue = null)
        {
            string condition = isSecondColumn ? (useOr ? " or " : " and ") : "";

            if (!string.IsNullOrEmpty(secondColumnName) && !string.IsNullOrEmpty(secondFilterValue))
            {
                // Add the second condition if provided
                condition += $"{secondColumnName} = '{secondFilterValue}'";
            }

            if (int.TryParse(filterValue, out _))
            {
                // Numeric value
                return $"{columnName} = {filterValue}" + condition;
            }
            else
            {
                // String value
                return $"{columnName} = '{filterValue}'" + condition;
            }
        }

        private void materialButton3_Click(object sender, EventArgs e)
        {
            string table = materialComboBox1.Text;
            switch (table)
            {
                case "Films":
                    filepath = $"{materialTextBox1.Text}.xlsx";
                    exportToExcel(dataGridView1, filepath);
                    break;
                case "Customers":
                    filepath = $"{materialTextBox1.Text}.xlsx";
                    exportToExcel(dataGridView2, filepath);
                    break;
                case "Tickets":
                    filepath = $"{materialTextBox1.Text}.xlsx";
                    exportToExcel(dataGridView4, filepath);
                    break;
                case "Rooms":
                    filepath = $"{materialTextBox1.Text}.xlsx";
                    exportToExcel(dataGridView3, filepath);
                    break;
                default:
                    MessageBox.Show("Такой таблицы не существует!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        private void materialSwitch5_CheckedChanged(object sender, EventArgs e)
        {
            if (materialSwitch5.Checked == true)
            {
                groupBox5.Visible = true;
            }
            else
            {
                groupBox5.Visible = false;
            }
        }

        private void materialSwitch1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
