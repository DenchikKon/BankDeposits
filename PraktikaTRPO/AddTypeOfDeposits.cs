using Guna.UI2.WinForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PraktikaTRPO
{
    public partial class AddTypeOfDeposits : Form
    {
        public AddTypeOfDeposits()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void AddTypeOfDeposits_Load(object sender, EventArgs e)
        {
            textBoxYearRate.MaxLength = 5;
            textBoxDepDate.MaxLength = 1;
            textBoxMoneyMin.MaxLength = 9;
            textBoxTitle.MaxLength = 20;
            Form1 form1 = this.Owner as Form1;
            string query = "Select distinct  TypeOfDeposits.TypeMoneyDeposit From TypeOfDeposits";
            DBManager.ComboBoxLoadData(query, comboBoxTypeOfMoney);
            if (buttonApply.Text == "Изменить")
            {
                textBoxTitle.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Название"].Value.ToString();
                comboBoxTypeOfMoney.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Валюта вклада"].Value.ToString();
                comboBoxTypeDeposit.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Тип депозита"].Value.ToString();
                textBoxYearRate.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Ставка"].Value.ToString();
                textBoxMoneyMin.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Мин. сумма вклада"].Value.ToString();
                textBoxDepDate.Text = form1.dataGridViewTypeOfDeposits.CurrentRow.Cells["Срок г."].Value.ToString();
            }
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            Form1 form1 = this.Owner as Form1;
            string query;                      
            int count;
            SqlCommand command;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            if (textBoxTitle.Text != "" && comboBoxTypeOfMoney.SelectedIndex != -1 && comboBoxTypeDeposit.SelectedIndex != -1 &&
                textBoxYearRate.Text != "" && textBoxMoneyMin.Text != "" && textBoxDepDate.Text != "")
            {
                if(double.TryParse(textBoxYearRate.Text,out double result) && double.Parse(textBoxYearRate.Text) > 0
                    && int.TryParse(textBoxMoneyMin.Text, out int  res) && double.Parse(textBoxMoneyMin.Text) > 0 
                    && int.TryParse(textBoxDepDate.Text, out int r) && int.Parse(textBoxDepDate.Text) > 0)
                {
                    switch (buttonApply.Text)
                    {
                        case "Добавить":
                             query = $"if (select count(Title) From TypeOfDeposits Where Title =N'{textBoxTitle.Text}')=0" +
                                " Insert into TypeOfDeposits(Title,TypeMoneyDeposit,TypeDeposit,YearRate,MinMoneyDeposit,DepositDate) Values " +
                        $"(N'{textBoxTitle.Text}',N'{comboBoxTypeOfMoney.Text}',N'{comboBoxTypeDeposit.Text}',{textBoxYearRate.Text}, {textBoxMoneyMin.Text}, {textBoxDepDate.Text})";
                            command= new SqlCommand(query,Form1.DBBankConnection);
                            count = command.ExecuteNonQuery();
                            if (count == -1)
                                MessageBox.Show("Данное название уже имеется в базе");
                            else
                                Hide();
                            break;
                        case "Изменить":
                            query = $"if (select count(Title) From TypeOfDeposits Where Title =N'{textBoxTitle.Text}'" +
                                $" and Id != {form1.dataGridViewTypeOfDeposits.CurrentRow.Cells[0].Value})=0 " +
                                $" Update TypeOfDeposits Set Title = N'{textBoxTitle.Text}', TypeMoneyDeposit = N'{comboBoxTypeOfMoney.Text}'," +
                                $" TypeDeposit = N'{comboBoxTypeDeposit.Text}', YearRate = {textBoxYearRate.Text.Replace(',','.')}," +
                                $" MinMoneyDeposit = {textBoxMoneyMin.Text}, DepositDate = {textBoxDepDate.Text} Where Id = {Form1.idTypeOfDeposits}";
                            command = new SqlCommand(query, Form1.DBBankConnection);
                            count = command.ExecuteNonQuery();
                            if (count == -1)
                                MessageBox.Show("Данное название уже имеется в базе");
                            else
                                Hide();
                            break;
                    }
                    
                }
                else
                {
                    MessageBox.Show("Выдимо вы ввели не число в требуемое поле(процентная ставка, сумма депозита, срок депозита) и данное значение не может равняться 0");
                }
            }
            else
            {
                MessageBox.Show("Заполните все поля");
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

        private void textBoxDepDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void textBoxYearRate_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8 && number != 46)
            {
                e.Handled = true;
            }
        }
    }
}
