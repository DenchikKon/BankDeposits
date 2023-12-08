using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PraktikaTRPO
{
    public partial class AddClient : Form
    {
        public AddClient()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void groupFilterClients_Click(object sender, EventArgs e)
        {

        }

        private void AddClient_Load(object sender, EventArgs e)
        {
            guna2DateTimePicker1.MaxDate = DateTime.Now;
            guna2DateTimePicker1.MinDate = Convert.ToDateTime("01.01.1900");
            textBoxPhone.MaxLength = 15;
            textBoxName.MaxLength = 15;
            textBoxSurname.MaxLength = 20;
            textBoxLastname.MaxLength = 20;
            textBoxPassportnumber.MaxLength = 9;
            textBoxAddress.MaxLength = 30;
            Form1 form1 = this.Owner as Form1;
            if(ButtonApply.Text == "Изменить")
            {
                textBoxName.Text = form1.dataGridViewClients.CurrentRow.Cells["Имя"].Value.ToString();
                textBoxSurname.Text = form1.dataGridViewClients.CurrentRow.Cells["Фамилия"].Value.ToString();
                textBoxLastname.Text = form1.dataGridViewClients.CurrentRow.Cells["Отчество"].Value.ToString();
                guna2DateTimePicker1.Value = Convert.ToDateTime(form1.dataGridViewClients.CurrentRow.Cells["Дата рождения"].Value);
                textBoxPassportnumber.Text = form1.dataGridViewClients.CurrentRow.Cells["Номер паспорта"].Value.ToString();
                textBoxPhone.Text = form1.dataGridViewClients.CurrentRow.Cells["Телефон"].Value.ToString();
                textBoxAddress.Text = form1.dataGridViewClients.CurrentRow.Cells["Адрес"].Value.ToString();
            }
        }

        private void ButtonApply_Click(object sender, EventArgs e)
        {
            Form1 form1 = this.Owner as Form1;
            Regex Isnumber = new Regex(@"^(80|375)\((44|29|33|25)\)[0-9]{7}$");
            Regex Ispassportnumber = new Regex(@"[A-Z]{2}[0-9]{7}$");            
            string query;
            SqlCommand command;
            int count;
            string passportNumber = textBoxPassportnumber.Text;
            switch (ButtonApply.Text)
            {
                case "Добавить":
                    if (textBoxName.Text != "" && textBoxSurname.Text != "" && textBoxLastname.Text != ""
                && textBoxPassportnumber.Text != "" && textBoxPhone.Text != "" && textBoxAddress.Text != "")
                    {
                        if ((((uint)guna2DateTimePicker1.Value.Year)) - (DateTime.Now.Year - 14) <= 0)
                        {
                            if (Ispassportnumber.IsMatch(textBoxPassportnumber.Text))
                            {
                                if (Isnumber.IsMatch(textBoxPhone.Text))
                                {
                                    query = $"if (select count(PassportNumber) From Clients Where PassportNumber =N'{textBoxPassportnumber.Text}')=0" +
                                        $" And (select count(Phone) From Clients Where Phone =N'{textBoxPhone.Text}')=0" +
                                        $" Insert Into Clients(Name,Surname,Lastname,Birthdate,PassportNumber,Phone,Address)" +
                                       $" Values (N'{textBoxName.Text}', N'{textBoxSurname.Text}', N'{textBoxLastname.Text}', N'{guna2DateTimePicker1.Value.ToString("yyyy/MM/dd")}', " +
                                       $" N'{textBoxPassportnumber.Text}', N'{textBoxPhone.Text}', N'{textBoxAddress.Text}')";
                                    command = new SqlCommand(query, Form1.DBBankConnection);
                                    count = command.ExecuteNonQuery();
                                    if (count == -1)
                                        MessageBox.Show("Данный номер телефона или номер паспорта уже имеется в базе");
                                    else
                                        Hide();
                                }
                                else
                                    MessageBox.Show("Некоректно набран номер");
                            }
                            else
                                MessageBox.Show("Укажите верный номер паспорта");
                        }
                        else
                            MessageBox.Show("Клиенту должно быть больше 14 лет");
                    }
                    else
                        MessageBox.Show("Заполните все требуемые поля");
                    
                    break;
                case "Изменить":
                    if (textBoxName.Text != "" && textBoxSurname.Text != "" && textBoxLastname.Text != ""
               && textBoxPassportnumber.Text != "" && textBoxPhone.Text != "" && textBoxAddress.Text != "")
                    {
                        if ((((uint)guna2DateTimePicker1.Value.Year)) - (DateTime.Now.Year - 14) <= 0)
                        {
                            if (Ispassportnumber.IsMatch(textBoxPassportnumber.Text))
                            {
                                if (Isnumber.IsMatch(textBoxPhone.Text))
                                {
                                    query = $"if (select count(PassportNumber) From Clients Where PassportNumber =N'{textBoxPassportnumber.Text}' and " +
                                    $"Id != {form1.dataGridViewClients.CurrentRow.Cells[0].Value})=0" +
                                    $" And (select count(Phone) From Clients Where Phone =N'{textBoxPhone.Text}' and " +
                                    $"Id != {form1.dataGridViewClients.CurrentRow.Cells[0].Value})=0 " +
                                    $"Update Clients Set Name = N'{textBoxName.Text}', Surname = N'{textBoxSurname.Text}', Lastname = N'{textBoxLastname.Text}'," +
                                    $" Birthdate = N'{guna2DateTimePicker1.Value.ToString("yyyy/MM/dd")}', PassportNumber = N'{textBoxPassportnumber.Text}'," +
                                    $" Phone = N'{textBoxPhone.Text}', Address = N'{textBoxAddress.Text}' Where id = {Form1.idClient}";
                                    command = new SqlCommand(query, Form1.DBBankConnection);
                                    count = command.ExecuteNonQuery();
                                    if (count == -1)
                                        MessageBox.Show("Данный номер телефона или номер паспорта уже имеется в базе");
                                    else
                                        Hide();
                                }
                                else
                                    MessageBox.Show("Некоректно набран номер");
                            }
                            else
                                MessageBox.Show("Укажите верный номер паспорта");
                        }
                        else
                            MessageBox.Show("Клиенту должно быть больше 14 лет");
                    }
                    else
                        MessageBox.Show("Заполните все требуемые поля");
                    break;
            }
            
        }

        private void guna2DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.SendKeys.Send("%{DOWN}");
        }

        private void textBoxName_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void textBoxSurname_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxSurname_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsLetter(number) && number != 8)
            {
                e.Handled = true;
            }
        }
    }
}
