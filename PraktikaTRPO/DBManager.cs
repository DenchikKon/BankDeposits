using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace PraktikaTRPO
{
    internal class DBManager
    {
        public static void LoadData(string query,DataGridView dataGrid)
        {
            SqlDataAdapter adapter = new SqlDataAdapter(query,Form1.DBBankConnection);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            dataGrid.DataSource = dataTable;
        }
        public static void ComboBoxLoadData(String query,ComboBox comboBox)
        {
            SqlCommand command = new SqlCommand(query,Form1.DBBankConnection);
            SqlDataReader dataReader= command.ExecuteReader();
            while (dataReader.Read())
            {
                comboBox.Items.Add(dataReader.GetValue(0));
            }
            dataReader.Close();
        }
        public static void ExecuteQuery(string query)
        {
            SqlCommand command = new SqlCommand(query, Form1.DBBankConnection);
            command.ExecuteNonQuery();
        }
        public static void LoadComboBoxNew(String query, ComboBox comboBox, string valueMember,string displayMember)
        {
            SqlDataAdapter adapter = new SqlDataAdapter(query, Form1.DBBankConnection);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            comboBox.DataSource= dataTable;
            comboBox.DisplayMember = $"{displayMember}";
            comboBox.ValueMember = $"{valueMember}";

        }
    }
}
