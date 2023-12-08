using System;
using System.Collections.Generic;
using System.ComponentModel;
using SD = System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Net.Sockets;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading;
using System.Runtime.InteropServices;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Drawing.Drawing2D;

namespace PraktikaTRPO
{
    public partial class Form1 : Form
    {
        public static int idClient;
        public static int idDeposits;
        public static int idTypeOfDeposits;
        public static string mainqueryClients = "Select Id, Name as 'Имя',Surname as 'Фамилия', Lastname 'Отчество', Format(Birthdate,'dd/MM/yyyy') as 'Дата рождения', PassportNumber as 'Номер паспорта', Phone as 'Телефон', Address as 'Адрес' From Clients";
        public static string mainqueryDeposits = "Select Deposits.Id, TypeOfDeposits.Title as 'Вид вклада', CONCAT(Clients.Name,' ',Clients.Surname,' ',Clients.Lastname) as 'Клиент',TypeOfDeposits.TypeMoneyDeposit as 'Валюта вклада',Deposits.DepositMoneyAmount as 'Депозит', Deposits.DateOpen as 'Открыт',\r\nDeposits.DateClose as 'Закрыт',TypeOfDeposits.YearRate as 'Годовая % ставка' From Deposits\r\nleft join Clients on Deposits.IDClient = Clients.Id\r\nleft join TypeOfDeposits on Deposits.IDTypeOfDeposit = TypeOfDeposits.Id";
        public static string mainqueryTypeOfDeposits = "Select TypeOfDeposits.Id, TypeOfDeposits.Title as 'Название', TypeOfDeposits.TypeMoneyDeposit as 'Валюта вклада', TypeOfDeposits.TypeDeposit as 'Тип депозита', TypeOfDeposits.YearRate as 'Ставка',\r\nTypeOfDeposits.MinMoneyDeposit as 'Мин. сумма вклада', TypeOfDeposits.DepositDate 'Срок г.' From TypeOfDeposits";
        public int IsDepositDate = 0, IsDepositMoney = 0, IsDepositTypeOfMoney = 0;
        public int IsTypeOfDepositTypeOfMoney = 0, IsTypeOfDepositTypeDeposit = 0, IsTypeOfDepositYearRate = 0, IsTypeOfDepositMoney = 0;
        public int IsClientName = 0, IsClientSurname = 0, IsClientLastname = 0, IsClientDate = 0, IsClientPasssportNumber = 0;
        private string query;
        public static SqlConnection DBBankConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["DBBankDeposits"].ToString());
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Clients_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            NumberFormatInfo numberFormatInfo = new NumberFormatInfo()
            {
                NumberDecimalSeparator = ".",
            };
            datePickerMinDate.Enabled = false;
            datePickerMaxDate.Enabled = false;
            checkBoxMinDate.Checked = false;
            checkBoxMaxDate.Checked = false;
            guna2DateTimePicker1.Enabled = false;
            checkBox1.Checked = false;
             DBBankConnection.Open();
            DBManager.LoadData(mainqueryClients, dataGridViewClients);
            dataGridViewClients.Columns[0].Visible = false;
            DBManager.LoadData(mainqueryDeposits, dataGridViewDeposits);
            dataGridViewDeposits.Columns[0].Visible = false;
            DBManager.LoadData(mainqueryTypeOfDeposits, dataGridViewTypeOfDeposits);
            dataGridViewTypeOfDeposits.Columns[0].Visible = false;
            query = "Select distinct  TypeOfDeposits.TypeMoneyDeposit From TypeOfDeposits";
            DBManager.ComboBoxLoadData(query, comboBoxTypeOfMoney);
            DBManager.ComboBoxLoadData(query, comboBoxTypeOfDepositsFilterDeposits);
            comboBoxTypeOfMoney.Items.Add("");
            comboBoxTypeOfDepositsFilterDeposits.Items.Add("");
            dataGridViewClients.ClearSelection();
            dataGridViewDeposits.ClearSelection();
            dataGridViewTypeOfDeposits.ClearSelection();
            guna2DateTimePicker1.MaxDate = DateTime.Now;
            guna2DateTimePicker1.MinDate = Convert.ToDateTime("01.01.1900");
            datePickerMinDate.MaxDate = DateTime.Now;
            datePickerMinDate.MinDate = Convert.ToDateTime("01.01.1900");
            datePickerMaxDate.MaxDate = DateTime.Now;
            datePickerMaxDate.MinDate = Convert.ToDateTime("01.01.1900");
            textBoxFilterName.MaxLength = 30;
            textBoxFilterSurname.MaxLength = 30;
            textBoxFilterLastname.MaxLength = 30;
            textBoxFilterPassportNumber.MaxLength = 9;
            TextBoxMinDeposit.MaxLength = 10;
            textBoxMaxDeposit.MaxLength = 10;
            textBoxYearRateMin.MaxLength = 6;
            textBoxYearRateMax.MaxLength = 6;
            textBoxMoneyMin.MaxLength = 10;
            textBoxMoneyMax.MaxLength = 10;

        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void TextBoxSearchClients_TextChanged(object sender, EventArgs e)
        {
            Methods.SearchData(dataGridViewClients, textBoxSearchClients);
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            Methods.SearchData(dataGridViewDeposits, textBoxSearchDeposits);
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {
            Methods.SearchData(dataGridViewTypeOfDeposits, textBoxSearchTypeOfDeposits);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBoxMinMoneyDeposits_TextChanged(object sender, EventArgs e)
        {

        }

  

        private void guna2DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.SendKeys.Send("%{DOWN}");
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if ((dataGridViewClients?.CurrentRow?.Index ?? -1) != -1 || (dataGridViewClients?.CurrentRow?.Index ?? 0) != 0)
                {                    
                    query = $"Delete Clients Where Id = {dataGridViewClients.CurrentRow.Cells[0].Value}";
                    DBManager.ExecuteQuery(query);
                    dataGridViewClients.Rows.RemoveAt(dataGridViewClients.CurrentRow.Index);
                }
                else
                    MessageBox.Show("Невозможно удалить данную строку", "Ошибка", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Не возможно удалить данного клиента у него есть депозиты");
            }
            
        }

        private void dataGridViewClients_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if(e.RowIndex != -1) 
                { 
                dataGridViewClients.ClearSelection();
                dataGridViewClients[e.ColumnIndex, e.RowIndex].Selected = true;
                }
            }
        }

        private void удалитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                if ((dataGridViewTypeOfDeposits?.CurrentRow?.Index ?? -1) != -1 || (dataGridViewTypeOfDeposits?.CurrentRow?.Index ?? 0) != 0)
                {
                    query = $"Delete TypeOfDeposits Where Id = {dataGridViewTypeOfDeposits.CurrentRow.Cells[0].Value}";
                    DBManager.ExecuteQuery(query);
                    dataGridViewTypeOfDeposits.Rows.RemoveAt(dataGridViewTypeOfDeposits.CurrentRow.Index);
                }
                else
                    MessageBox.Show("Невозможно удалить данную строку", "Ошибка", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Данный вид вклада имеет открытые депозиты");
            }
        }

        private void dataGridViewTypeOfDeposits_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex != -1)
                {
                    dataGridViewTypeOfDeposits.ClearSelection();
                    dataGridViewTypeOfDeposits[e.ColumnIndex, e.RowIndex].Selected = true;
                }
            }
        }

        private void dataGridViewDeposits_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex != -1)
                {
                    dataGridViewDeposits.ClearSelection();
                    dataGridViewDeposits[e.ColumnIndex, e.RowIndex].Selected = true;
                }
            }
        }

        private void guna2DateTimePicker1_CloseUp(object sender, EventArgs e)
        {
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void guna2TextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void ButtonAddClient_Click(object sender, EventArgs e)
        {
            AddClient addClient = new AddClient();
            addClient.Owner = this;
            addClient.ButtonApply.Text = "Добавить";
            addClient.groupFilterClients.Text = "Добавить клиента";

            addClient.ShowDialog();
            DBManager.LoadData(mainqueryClients, dataGridViewClients);
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
                AddClient addClient = new AddClient();
                addClient.Owner = this;
                idClient = Convert.ToInt32(dataGridViewClients.CurrentRow.Cells[0].Value);
                addClient.ButtonApply.Text = "Изменить";
                addClient.groupFilterClients.Text = "Изменить данные клиента";
                addClient.ShowDialog();
                DBManager.LoadData(mainqueryClients, dataGridViewClients);
        }

        private void guna2DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.SendKeys.Send("%{DOWN}");
        }

        private void guna2DateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.SendKeys.Send("%{DOWN}");
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            string date = guna2DateTimePicker1.Enabled ? guna2DateTimePicker1.Value.ToString("yyyy/MM/dd") : "";
            //for print excel filter
            if (textBoxFilterName.Text != "") IsClientName = 1;
            if(textBoxFilterSurname.Text != "") IsClientSurname= 1;
            if (textBoxFilterLastname.Text != "") IsClientLastname = 1;
            if (guna2DateTimePicker1.Enabled == true) IsClientDate = 1;

            if (textBoxFilterPassportNumber.Text != "") IsClientPasssportNumber = 1; 
            query = "Select Id, Name as 'Имя',Surname as 'Фамилия', Lastname 'Отчество', Format(Birthdate,'dd/MM/yyyy') as 'Дата рождения'" +
                ", PassportNumber as 'Номер паспорта', Phone as 'Телефон', Address as 'Адрес' From Clients " +
                $"Where Name Like N'%{textBoxFilterName.Text}%' And Surname Like N'%{textBoxFilterSurname.Text}%' And Lastname Like N'%{textBoxFilterLastname.Text}%'" +
                $" And PassportNumber Like N'%{textBoxFilterPassportNumber.Text}%'";
            if (guna2DateTimePicker1.Enabled) { query += $"And Birthdate = '{date}'"; }
                DBManager.LoadData(query, dataGridViewClients);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
                guna2DateTimePicker1.Enabled = true;
            else
                guna2DateTimePicker1.Enabled = false;
        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((dataGridViewDeposits?.CurrentRow?.Index ?? -1) != -1 || (dataGridViewDeposits?.CurrentRow?.Index ?? 0) != 0)
                {
                    query = $"Delete Deposits Where Id = {dataGridViewDeposits.CurrentRow.Cells[0].Value}";
                    DBManager.ExecuteQuery(query);
                    dataGridViewDeposits.Rows.RemoveAt(dataGridViewDeposits.CurrentRow.Index);
                }
                else
                    MessageBox.Show("Невозможно удалить данную строку", "Ошибка", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось удалить данный депозит");
            }
        }

        private void ButtonAddTypeOfDeposits_Click(object sender, EventArgs e)
        {
            AddTypeOfDeposits addTypeOfDeposits = new AddTypeOfDeposits();
            addTypeOfDeposits.Owner= this;
            idTypeOfDeposits = Convert.ToInt32(dataGridViewTypeOfDeposits.CurrentRow.Cells[0].Value);
            addTypeOfDeposits.buttonApply.Text = "Добавить";
            addTypeOfDeposits.groupFilterTypeOfDeposits.Text = "Добавить депозит";
            addTypeOfDeposits.ShowDialog();
            DBManager.LoadData(mainqueryTypeOfDeposits, dataGridViewTypeOfDeposits);
        }

        private void изменитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            AddTypeOfDeposits addTypeOfDeposits = new AddTypeOfDeposits();
            addTypeOfDeposits.Owner = this;
            idTypeOfDeposits = Convert.ToInt32(dataGridViewTypeOfDeposits.CurrentRow.Cells[0].Value);
            addTypeOfDeposits.buttonApply.Text = "Изменить";
            addTypeOfDeposits.groupFilterTypeOfDeposits.Text = "Изменить депозит";
            addTypeOfDeposits.ShowDialog();
            DBManager.LoadData(mainqueryTypeOfDeposits, dataGridViewTypeOfDeposits);
        }

        private void ApplyFilterTypeOfDeposits_Click(object sender, EventArgs e)
        {
            try
            {
                string minRate = textBoxYearRateMin.Text == "" ? "0" : textBoxYearRateMin.Text;
                string maxRate = textBoxYearRateMax.Text == "" ? $"{int.MaxValue}" : textBoxYearRateMax.Text;
                string moneyMin = textBoxMoneyMin.Text == "" ? "0" : textBoxMoneyMin.Text;
                string moneyMax = textBoxMoneyMax.Text == "" ? $"{int.MaxValue}" : textBoxMoneyMax.Text;
                Thread.CurrentThread.CurrentCulture=CultureInfo.GetCultureInfo("en-US");
                if (double.Parse(minRate) < double.Parse(maxRate) && int.Parse(moneyMin) < int.Parse(moneyMax)) {
                    //for print excel filter
                    if (comboBoxTypeOfMoney.Text != "") IsDepositTypeOfMoney = 1;
                    if (comboBoxTypeDeposit.Text != "") IsTypeOfDepositTypeDeposit = 1;
                    if (textBoxYearRateMin.Text != "" || textBoxYearRateMax.Text != "") IsTypeOfDepositYearRate = 1;
                    if (textBoxMoneyMin.Text != "" || textBoxMoneyMax.Text != "") IsTypeOfDepositMoney = 1;
                          
                query = mainqueryTypeOfDeposits + $" Where TypeMoneyDeposit Like N'%{comboBoxTypeOfMoney.Text}%' AND TypeDeposit Like N'%{comboBoxTypeDeposit.Text}%' " +
                        $" AND YearRate between {minRate} AND {maxRate} " +
                            $" AND MinMoneyDeposit between {moneyMin} AND {moneyMax}";
                DBManager.LoadData(query, dataGridViewTypeOfDeposits);
            }
                else
            {
                MessageBox.Show("Минимальное граница не может быть больше максимальной");
            }
        }
            catch (Exception)
            {
                MessageBox.Show("Введите число");
            }

}

        private void buttonClearTypeOfDeposits_Click(object sender, EventArgs e)
        {
            DBManager.LoadData(mainqueryTypeOfDeposits, dataGridViewTypeOfDeposits);
            comboBoxTypeOfMoney.SelectedItem = null;
            comboBoxTypeDeposit.SelectedItem = null;
            textBoxYearRateMin.Text = "";
            textBoxYearRateMax.Text = "";
            textBoxMoneyMin.Text = "";
            textBoxMoneyMax.Text = "";

            IsDepositTypeOfMoney = 0;
            IsTypeOfDepositTypeDeposit = 0;
            IsTypeOfDepositYearRate = 0;
            IsTypeOfDepositMoney = 0;
        }

        private void buttonCancelDeposits_Click(object sender, EventArgs e)
        {
            DBManager.LoadData(mainqueryDeposits, dataGridViewDeposits);
            IsDepositMoney = 0;
            IsDepositTypeOfMoney = 0;
            IsDepositDate = 0;
            datePickerMinDate.Enabled= false;
            datePickerMaxDate.Enabled= false;
            checkBoxMinDate.Checked = false;
            checkBoxMaxDate.Checked = false;
            textBoxFilterName.Text = "";
            textBoxFilterSurname.Text = "";
            textBoxFilterLastname.Text = "";
            textBoxFilterPassportNumber.Text = "";
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            DBManager.LoadData(mainqueryClients, dataGridViewClients);
            textBoxFilterName.Text = "";
            textBoxFilterSurname.Text = "";
            textBoxFilterLastname.Text = "";
            guna2DateTimePicker1.Enabled = false;
            checkBox1.Checked = false;
            textBoxFilterPassportNumber.Text = "";

            IsClientName = 1;
            IsClientSurname = 1;
            IsClientLastname = 1;
            IsClientDate = 1;
            IsClientPasssportNumber = 1;
        }

        private void textBoxFilterSurname_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsLetter(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void TextBoxMinDeposit_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void textBoxMaxDeposit_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void textBoxYearRateMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 46)
            {
                e.Handled = true;
            }
        }

        private void textBoxYearRateMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 46)
            {
                e.Handled = true;
            }
        }

        private void textBoxMoneyMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void textBoxMoneyMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void ButtonOpenDeposits_Click(object sender, EventArgs e)
        {
            AddDeposits addDeposits = new AddDeposits();
            addDeposits.Owner = this;
            addDeposits.ShowDialog();
            DBManager.LoadData(mainqueryDeposits,dataGridViewDeposits);
        }

        private void изменитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string nameDeposit = dataGridViewDeposits.CurrentRow.Cells["Вид вклада"].Value.ToString();
            string query = $"Select TypeDeposit From TypeOfDeposits Where Title = N'{nameDeposit}'";
            SqlCommand command = new SqlCommand(query, DBBankConnection);
            string typeDeposit = Convert.ToString(command.ExecuteScalar());
            query = $"Select DepositDate From TypeOfDeposits Where Title = N'{nameDeposit}'";
            command = new SqlCommand(query, Form1.DBBankConnection);
            int dateDeposit = Convert.ToInt32(command.ExecuteScalar());
            DateTime dateTime = (Convert.ToDateTime(dataGridViewDeposits.CurrentRow.Cells["Открыт"].Value)).AddYears(dateDeposit);

            if (typeDeposit == "Возвращаемый")
            {
                query = $"Update Deposits Set DateClose = '{DateTime.Now.ToString("yyyy/MM/dd")}' Where Id = {dataGridViewDeposits.CurrentRow.Cells[0].Value}";
                DBManager.ExecuteQuery(query);
                DBManager.LoadData(mainqueryDeposits, dataGridViewDeposits);
            }
            else if (typeDeposit == "Не возвращаемый")
            {
                if (dateTime <= DateTime.Now)
                {
                    query = $"Update Deposits Set DateClose = '{DateTime.Now.ToString("yyyy/MM/dd")}' Where Id = {dataGridViewDeposits.CurrentRow.Cells[0].Value}";
                    DBManager.ExecuteQuery(query);
                    DBManager.LoadData(mainqueryDeposits, dataGridViewDeposits);
                }
                else
                {
                    MessageBox.Show("Данный вклад невозможно закрыть в данный момет");
                }
            }
            else
                MessageBox.Show("Данный вклад не может быть возвращён из-за его типа");
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            string mindate = datePickerMinDate.Enabled ? datePickerMinDate.Value.ToString("yyyy/MM/dd") : "0001.01.01";
            string maxdate = datePickerMaxDate.Enabled ? datePickerMaxDate.Value.ToString("yyyy/MM/dd") : "2040.10.10";
            string minDeposit = TextBoxMinDeposit.Text == "" ? "0" : TextBoxMinDeposit.Text;
            string maxDeposit = textBoxMaxDeposit.Text == "" ? $"{int.MaxValue}" : textBoxMaxDeposit.Text;
            if (int.TryParse(minDeposit, out int res) && int.TryParse(maxDeposit, out int result))
            {
                if (Convert.ToInt32(minDeposit) <= Convert.ToInt32(maxDeposit))
                {
                    if (Convert.ToDateTime(mindate) <= Convert.ToDateTime(maxdate))
                    {
                        if (comboBoxTypeOfDepositsFilterDeposits.Text != "" ) IsDepositTypeOfMoney = 1;
                        if (TextBoxMinDeposit.Text != ""|| textBoxMaxDeposit.Text != "") IsDepositMoney = 1;
                        if (datePickerMinDate.Enabled == true || datePickerMaxDate.Enabled == true) IsDepositDate = 1;
                            string query = $"Select Deposits.Id, TypeOfDeposits.Title as 'Вид вклада', CONCAT(Clients.Name, ' ', Clients.Surname, ' ', Clients.Lastname) " +
                        $"as 'Клиент',TypeOfDeposits.TypeMoneyDeposit as 'Валюта вклада',Deposits.DepositMoneyAmount as 'Депозит',  " +
                        $"Deposits.DateOpen as 'Открыт', Deposits.DateClose as 'Закрыт',TypeOfDeposits.YearRate as 'Годовая % ставка' From Deposits " +
                        $"left join Clients on Deposits.IDClient = Clients.Id " +
                        $"left join TypeOfDeposits on Deposits.IDTypeOfDeposit = TypeOfDeposits.Id " +
                        $"Where Deposits.DepositMoneyAmount between {minDeposit} AND {maxDeposit}" +
                        $" AND Deposits.DateOpen between '{mindate}' AND '{maxdate}' " +
                        $"AND TypeOfDeposits.TypeMoneyDeposit Like '%{comboBoxTypeOfDepositsFilterDeposits.Text}%'";
                        DBManager.LoadData(query, dataGridViewDeposits);
                    }
                    else
                        MessageBox.Show("начальная дата не может быть больше максимальной");
                }
                else
                    MessageBox.Show("Минимальный депозит не может быть больше максимального");
            }
            else
                MessageBox.Show("Требуется ввести числовое значение");
           
        }

        private void checkBoxMinDate_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBoxMinDate.Checked)
            {
                datePickerMinDate.Enabled = true;
            }
            else 
            {
                datePickerMinDate.Enabled = false;
            }
        }

        private void checkBoxMaxDate_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMaxDate.Checked)
            {
                datePickerMaxDate.Enabled = true;
            }
            else
            {
                datePickerMaxDate.Enabled = false;
            }
        }

        private void buttonExportExcel_Click(object sender, EventArgs e)
        {
            string ClientFilterText = "";
            if (IsClientName == 1) ClientFilterText += $" Имя: {textBoxFilterName.Text} ";
            if (IsClientSurname == 1) ClientFilterText += $" Фамилия: {textBoxFilterSurname.Text} ";
            if (IsClientLastname == 1) ClientFilterText += $" Отчёство: {textBoxFilterLastname.Text} ";
            if (IsClientDate == 1) ClientFilterText += $" Дата рождения: {guna2DateTimePicker1.Value} ";
            if (IsClientPasssportNumber == 1) ClientFilterText += $" Номер паспорта: {textBoxFilterPassportNumber.Text} ";
            
            Methods.ExportExcel(dataGridViewClients, ClientFilterText, "Клиенты");
        }

        private void buttonExportExcelDeposits_Click(object sender, EventArgs e)
        {
            string DepositFilterText = "";           
            if (IsDepositDate == 1)
            {
                DepositFilterText += " Дата:";
                if (datePickerMinDate.Enabled == true) { DepositFilterText += " от " + datePickerMinDate.Value; }
                if (datePickerMaxDate.Enabled == true) DepositFilterText += " до " + datePickerMaxDate.Value;
                       
            }
            if (IsDepositMoney == 1)
            {
                DepositFilterText += " Сумма вклада:";
                if (TextBoxMinDeposit.Text != "") DepositFilterText += " от " + TextBoxMinDeposit.Text;
                if (textBoxMaxDeposit.Text != "") DepositFilterText += " до " + textBoxMaxDeposit.Text;
            }
            if (IsDepositTypeOfMoney == 1) DepositFilterText += " Вид валюты: " + comboBoxTypeOfDepositsFilterDeposits.Text;
            Methods.ExportExcel(dataGridViewDeposits, DepositFilterText, "Депозиты");
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            string typeOfDepositFilter = "";
            if (comboBoxTypeOfMoney.Text != "") IsDepositTypeOfMoney = 1;
            if (comboBoxTypeDeposit.Text != "") IsTypeOfDepositTypeDeposit = 1;
            if (textBoxYearRateMin.Text != "" || textBoxYearRateMax.Text != "") IsTypeOfDepositYearRate = 1;
            if (textBoxMoneyMin.Text != "" || textBoxMoneyMax.Text != "") IsTypeOfDepositMoney = 1;

            if (IsDepositTypeOfMoney == 1) typeOfDepositFilter += $" Валюта вклада: {comboBoxTypeOfMoney.Text} ";
            if (IsTypeOfDepositTypeDeposit == 1) typeOfDepositFilter += $" Вид депозита: {comboBoxTypeDeposit.Text} ";
            if(IsTypeOfDepositYearRate == 1)
            {
                typeOfDepositFilter += " Годовая % ставка: ";
                if (textBoxYearRateMin.Text != "") typeOfDepositFilter += $" от {textBoxYearRateMin.Text} ";
                if (textBoxYearRateMax.Text != "") typeOfDepositFilter += $" до {textBoxYearRateMax.Text} ";
            }
            if (IsTypeOfDepositMoney == 1)
            {
                typeOfDepositFilter += " Сумма вклада: ";
                if (textBoxMoneyMin.Text != "") typeOfDepositFilter += $" от {textBoxMoneyMin.Text} ";
                if (textBoxMoneyMax.Text != "") typeOfDepositFilter += $" от {textBoxMoneyMax.Text} ";
            }

            Methods.ExportExcel(dataGridViewTypeOfDeposits,typeOfDepositFilter, "Виды вкладов");
        }

        private void buttonExportWord_Click(object sender, EventArgs e)
        {
            FileInfo fileInfo;
            string fileName = "ДОГОВОР.docx";
            if (File.Exists(fileName))
            {
                fileInfo = new FileInfo(fileName);

                string nameDeposit = dataGridViewDeposits.CurrentRow.Cells["Вид вклада"].Value.ToString();
                string query = $"Select TypeDeposit From TypeOfDeposits Where Title = N'{nameDeposit}'";
                SqlCommand command = new SqlCommand(query, DBBankConnection);
                string typeDeposit = Convert.ToString(command.ExecuteScalar());
                query = $"Select DepositDate From TypeOfDeposits Where Title = N'{nameDeposit}'";
                command = new SqlCommand(query, Form1.DBBankConnection);
                int dateDeposit = Convert.ToInt32(command.ExecuteScalar());
                DateTime dateTime = (Convert.ToDateTime(dataGridViewDeposits.CurrentRow.Cells["Открыт"].Value)).AddYears(dateDeposit);

                var items = new Dictionary<string, string>
            {
                {"<DATE>", DateTime.Now.ToString("dd/MM/yyyy") },
                {"<FIO>", dataGridViewDeposits.CurrentRow.Cells["Клиент"].Value.ToString()},
                {"<TYPE>", typeDeposit },
                {"<DATEEND>", dateTime.ToString("dd/MM/yyyy") },
                {"<MONEY>", dataGridViewDeposits.CurrentRow.Cells["Депозит"].Value.ToString()},
                {"<YEARRATE>", dataGridViewDeposits.CurrentRow.Cells["Годовая % ставка"].Value.ToString() }

            };

                try
                {
                    Word.Application app = new Word.Application();
                    Object file = fileInfo.FullName;

                    Object missing = Type.Missing;

                    app.Documents.Open(@"D:\WebSharp\PraktikaTRPO\PraktikaTRPO\bin\Debug\ДОГОВОР.docx");
                    app.Visible = false;
                    foreach (var item in items)
                    {
                        Word.Find find = app.Selection.Find;
                        find.Text = item.Key.ToString();
                        find.Replacement.Text = item.Value.ToString();

                        object wrap = Word.WdFindWrap.wdFindContinue;
                        object replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing,
                            MatchCase: false,
                            MatchWholeWord: false,
                            MatchWildcards: false,
                            MatchSoundsLike: missing,
                            MatchAllWordForms: false,
                            Forward: true,
                            Wrap: wrap,
                            Format: false,
                            ReplaceWith: missing, Replace: replace);
                    }
                    app.Visible = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка");
                }                
            }
            else
            {
                MessageBox.Show("Файл не найден");
            }

            
        }
    }
}