//Код функций формы для добавления
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace Year2
{
    public partial class Add : Form
    {
        static int counter = 0;
        public Add()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Regex r = new Regex(@"[A-z]+");
            if (r.IsMatch(textBox1.Text) && r.IsMatch(textBox2.Text))
            {
                User u = new User(textBox1.Text + " " + textBox2.Text, comboBox1.Text);
                string sql = "Insert into Users values (" + Main.uid + "," + "\'" + u.FullName + "\'," + "\'" + u.Role + "\');";
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Main.conn;
                cmd.CommandText = sql;
                int bin = cmd.ExecuteNonQuery();
                counter++;
                this.Close();
            }
            else
            {
                MessageBox.Show("Please, check the spelling of both first name and last name. Both must only include letters.");
            }
        }

        private void Add_Load(object sender, EventArgs e)
        {

        }
    }
}
