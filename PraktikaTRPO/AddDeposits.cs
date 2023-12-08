using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PraktikaTRPO
{
    public partial class AddDeposits : Form
    {
        public AddDeposits()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void AddDeposits_Load(object sender, EventArgs e)
        {
            string query = "Select id, TypeOfDeposits.Title From TypeOfDeposits";
            DBManager.LoadComboBoxNew(query, comboBoxTypeOfDeposit, "id", "Title");
            query = "Select id, Concat(Name,' ',Surname,' ',Lastname) as 'Fio' From Clients";
            DBManager.LoadComboBoxNew(query, comboBoxClients, "id", "Fio");
            comboBoxTypeOfDeposit.SelectedIndex = -1;
            comboBoxClients.SelectedIndex = -1;
            textBoxMoney.MaxLength = 9;
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            if(textBoxMoney.Text != "" && comboBoxClients.SelectedIndex != -1 && comboBoxTypeOfDeposit.SelectedIndex != -1)
            {
                int idTypeOfDeposits = Convert.ToInt32(comboBoxTypeOfDeposit.SelectedValue);
                string query = $"Select TypeOfDeposits.MinMoneyDeposit From TypeOfDeposits Where Id = {idTypeOfDeposits}";
                SqlCommand command= new SqlCommand(query,Form1.DBBankConnection);
                int minMoney = Convert.ToInt32(command.ExecuteScalar());
                if (int.TryParse(textBoxMoney.Text, out int res) && int.Parse(textBoxMoney.Text) > 0)
                {
                    if (Convert.ToInt32(textBoxMoney.Text) >= minMoney)
                    {
                        query = $"Insert Into Deposits(IDTypeOfDeposit, IDClient,DepositMoneyAmount,DateOpen) " +
                            $"Values({comboBoxTypeOfDeposit.SelectedValue},{comboBoxClients.SelectedValue},{textBoxMoney.Text},'{DateTime.Now.ToString("yyyy/MM/dd")}')";
                        DBManager.ExecuteQuery(query);
                        Close();
                    }
                    else
                        MessageBox.Show("Введённая сумма меньше минимальной суммы данного вклада");
                }
                else
                    MessageBox.Show("В поле сумма вклада должно быть введено число отличное от 0");
            }
            else
                MessageBox.Show("Заполните все поля");
        }

        private void textBoxMoney_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }
    }
}
