//Код основной формы
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
using Microsoft.Office.Interop.Word;

namespace Year2
{
    public partial class Main : Form
    {
        static int res = 1;
        static int sm = 1;
        static int inv = 1;
        static int it = 0;
        static int an = 1;
        static int question = 1;
        static int questionnaire = 1;
        static int project = 1;
        public static int uid=1;
        public static SqlConnection conn;
        public Main()
        {
            InitializeComponent();
        }
        public List<string[]> sqldata;
        public List<string[]> GetData(string sql,int num)
        {
            List<string[]> data = new List<string[]>();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                data.Add(new string[num]);
                for (int i = 0; i < num; i++)
                {
                    data[data.Count - 1][i] = reader[i].ToString();
                }
            }
            reader.Close();
            return data;
        }
        public void Execute(string sql)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            int bin = cmd.ExecuteNonQuery();
        }
        public int GetNum(string sql,int variable)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            SqlDataReader reader = cmd.ExecuteReader();
            try
            {
                while (reader.Read())
                {
                    if (variable < Convert.ToInt32(reader[0].ToString()))
                        variable = Convert.ToInt32(reader[0].ToString());
                }
            }
            catch
            {
                variable = 0;
            }
            variable++;
            reader.Close();
            return variable;
        }
        private void Delu_Click(object sender, EventArgs e)
        {
            try
            {
                string[] txt = Ubox.Text.Split('|');
                string sql = "delete from users where userid=" + txt[0] + ";";
                Execute(sql);
                sql = "SELECT UserID,FullName,Role From Users";
                Ubox.Items.Clear();
                sqldata = GetData(sql, 3);
                foreach (string[] x in sqldata)
                {
                    Ubox.Items.Add(x[0] + "|" + x[1] + "|" + x[2]);
                }
                MessageBox.Show("Don't forget to choose a user againn.");
                sqldata.Clear();
            }
            catch
            {
                MessageBox.Show("Make sure you chose a user.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sqldata = new List<string[]>();
            qgrid.Columns.Add("Question Text", "Question Text");
            qgrid.Columns.Add("Unit", "Unit");
            qgrid.Columns.Add("Max", "Max");
            qgrid.Columns.Add("Min", "Min");
            qgrid.Rows.Add();
            conn = DBUtils.GetDBConnection();
            conn.Open();
            string[] txt = Ubox.Text.Split('|');
            string sql = "SELECT UserID,FullName,Role From Users";
            sqldata = GetData(sql, 3);
            foreach (string[] x in sqldata)
            {
                Ubox.Items.Add(x[0] + "|" + x[1] + "|" + x[2]);
            }
            sqldata.Clear();
            sql = "select Max(QuestionID) from question;";
            question=GetNum(sql, question);
            sql = "select Max(QID) from questionnaire;";
            questionnaire=GetNum(sql, questionnaire);
            sql = "select Max(UserID) from Users;";
            uid=GetNum(sql, uid);
            sql = "select Max(ProjectID) from Project;";
            project=GetNum(sql, project);
            sql = "select * from project;";
            sqldata = GetData(sql, 2);
            comboBox1.Items.Clear();
            foreach (string[] x in sqldata)
            {
                comboBox1.Items.Add(x[0] + "|" + x[1]);
            }
            sqldata.Clear();
            sql = "select Max(InvolvmentID) from Involvment;";
            inv=GetNum(sql, inv);
        }
        private void Addu_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Don't forget to press the \'Refresh\' button after adding new user.");
            Add a = new Add();
            a.Show();
        }
        public void FullUpdate()
        {
            if (Ubox.Text.Contains("Administrator"))
            {
                string[] txt = Ubox.Text.Split('|');
                string sql = "Select UserID, FullName from users where(Role='Expert');";
                Aclb1.Items.Clear();
                sqldata = GetData(sql, 2);
                foreach (string[] x in sqldata)
                {
                    Aclb1.Items.Add(x[0] + "|" + x[1]);
                }
                sqldata.Clear();
                comboBox1.Items.Clear();
                sql = "select * from project";
                sqldata = GetData(sql, 2);
                foreach(string[] x in sqldata)
                {
                    comboBox1.Items.Add(x[0] + "|" + x[1]);
                }
                Acb.Items.Clear();
                itbox.Items.Clear();
                sqldata.Clear();
                Art.Clear();
            }
            else if (Ubox.Text.Contains("Expert"))
            {
                other.Rows.Clear();other.Columns.Clear();
                string[] txt = Ecb.Text.Split('|');
                Eqgrid.Rows.Clear(); Eqgrid.Columns.Clear();
                Eagrid.Rows.Clear(); Eagrid.Columns.Clear();
                Eqgrid.Columns.Add("Question Text", "Question Text");
                Eqgrid.Columns.Add("Unit", "Unit");
                Eqgrid.Columns.Add("Max", "Max");
                Eqgrid.Columns.Add("Min", "Min");
                string sql = "Select questionid,unit,maximum,minimum,text from question where qid=" + txt[0] + ";";
                sqldata = GetData(sql, 5);
                sql = "select Max(iteration) from answer where (userid=" + Ubox.Text[0] + " and questionid=(select top 1 questionid from question where qid="+Ecb.Text[0]+"));";
                var cmd = new SqlCommand();
                cmd.Connection = Main.conn;
                cmd.CommandText = sql;
                SqlDataReader reader = cmd.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (it < Convert.ToInt32(reader[0].ToString()))
                            it = Convert.ToInt32(reader[0].ToString());
                    }

                }
                catch
                {
                    it = 0;
                }
                reader.Close();
                if (it < 1)
                {
                    El3.Visible = true;
                    etb.Visible = true;

                }
                else
                {
                   
                }
                Eagrid.Columns.Add("Answer", "Answer");Eagrid.Columns.Add("Explanation", "Explanation");
                for (int i = 0; i < sqldata.Count; i++)
                {
                    Eqgrid.Rows.Add(); Eqgrid.Rows[i].HeaderCell.Value = sqldata[i][0];
                    Eqgrid.Rows[i].Cells[0].Value = sqldata[i][4];
                    Eqgrid.Rows[i].Cells[1].Value = sqldata[i][1];
                    Eqgrid.Rows[i].Cells[2].Value = sqldata[i][2];
                    Eqgrid.Rows[i].Cells[3].Value = sqldata[i][3];
                    Eagrid.Rows.Add(); Eagrid.Rows[i].HeaderCell.Value = sqldata[i][0];
                }
                sqldata.Clear();
                Sb.Visible = true;
                other.Columns.Clear();other.Rows.Clear();
                sql = "select questionid from question where qid=" + Ecb.Text[0] + ";";
                sqldata = GetData(sql, 1);
                int max=0;
                sql = "select max(iteration) from answer where questionid=" + sqldata[0][0] + ";";
                cmd = new SqlCommand();
                cmd.Connection = Main.conn;
                cmd.CommandText = sql;
                reader = cmd.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (max < Convert.ToInt32(reader[0].ToString()))
                            max = Convert.ToInt32(reader[0].ToString());
                    }
                }
                catch
                {
                    max = 0;
                }
                reader.Close();
                if (max >= 1)
                {
                    other.Columns.Add("Answer", "Answer");
                    other.Columns.Add("Explanation", "Explanation");
                    foreach(string[] x in sqldata)
                    {
                        sql = "select answer, explanation from answer where questionid=" + x[0] + " and iteration=" + max + " and userid!="+Ubox.Text[0]+";";
                        List<string[]> ans=GetData(sql,2);
                        foreach (string[] y in ans)
                        {
                            other.Rows.Add();other.Rows[other.Rows.Count - 1].HeaderCell.Value = x[0];
                            other.Rows[other.Rows.Count - 1].Cells[0].Value = y[0];
                            other.Rows[other.Rows.Count - 1].Cells[1].Value = y[1];
                        }
                    }
                    sqldata.Clear();
                }
            }
        } 

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Press the \"Add\" button to open a form for new user registration. \n Press the \"Delete\" button to delete the chosen user from the memory.");
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Ubox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Acb.Items.Clear();
            Questtab.Rows.Clear(); Questtab.Columns.Clear();Aagrid.Rows.Clear();Aagrid.Columns.Clear();
            other.Rows.Clear();other.Columns.Clear();
            string[] txt = Ubox.Text.Split('|');
            if (txt[2] == "Administrator")
            {
                Acb.Text = "";
                comboBox1.Text = "";
                Al.Visible = true;
                Al1.Visible = true;
                Al2.Visible = true;
                Al3.Visible = true;
                Al4.Visible = true;
                Al5.Visible = true;
                Ab.Visible = true;
                Ab1.Visible = true;
                qgrid.Visible = true;
                Aagrid.Visible = true;
                Aclb1.Visible = true;
                Acb.Visible = true;
                Art.Visible = true;
                itbox.Visible = true;
                textBox1.Visible = true;
                CQ.Visible = true;
                button2.Visible = true;
                button4.Visible = true;
                button3.Visible = true;
                label2.Visible = true;
                comboBox1.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                Questtab.Visible = true;
                Ecb.Visible = false;
                El.Visible = false;
                Eqgrid.Visible = false;
                El1.Visible = false;
                Eagrid.Visible = false;
                Sb.Visible = false;
                El3.Visible = false;
                etb.Visible = false;
                other.Visible = false;
                A1.Visible = true;
                A2.Visible = true;
                A3.Visible = true;
                A4.Visible = false;
                A5.Visible = false;
                label3.Visible = false;
                textBox2.Visible = false;
                button7.Visible = false;
                FullUpdate();
            }
            if (txt[2] == "Expert")
            {
                Eqgrid.Rows.Clear(); Eqgrid.Columns.Clear();
                Eagrid.Rows.Clear(); Eagrid.Columns.Clear();
                A4.Visible = true;
                A5.Visible = true;
                label3.Visible = false;
                textBox2.Visible = false;
                button7.Visible = false;
                Al.Visible = false;
                Al1.Visible = false;
                Al2.Visible = false;
                Al3.Visible = false;
                Al4.Visible = false;
                Al5.Visible = false;
                Ab.Visible = false;
                Ab1.Visible = false;
                qgrid.Visible = false;
                Aagrid.Visible = false;
                Aclb1.Visible = false;
                Questtab.Visible = false;
                Acb.Visible = false;
                Art.Visible = false;
                itbox.Visible = false;
                CQ.Visible = false;
                button2.Visible = false;
                button4.Visible = false;
                button3.Visible = false;
                textBox1.Visible = false;
                label2.Visible = false;
                comboBox1.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                Ecb.Visible = true;
                El.Visible = true;
                Eqgrid.Visible = true;
                El1.Visible = true;
                Eagrid.Visible = true;
                Sb.Visible = false ;
                other.Visible = true;
                A1.Visible = false;
                A2.Visible = false;
                A3.Visible = false;
                Ecb.Items.Clear();
                txt = Ubox.Text.Split('|');
                string sql = "SELECT QID from Involvment where UID=" + txt[0] + ";";
                sqldata = GetData(sql,1);
                foreach (string[] x in sqldata)
                {
                    sql = "Select QID,Name from questionnaire where qid=" + x[0] + ";";
                    List<string[]> tem = GetData(sql,2);
                    foreach (string[] y in tem)
                    { 
                        Ecb.Items.Add(y[0] + "|" + y[1]);
                    }
                    tem.Clear();
                }
                sqldata.Clear();
            }
        }
        class DBSQLServerUtils
        {

            public static SqlConnection
                     GetDBConnection(string datasource, string database)
            {
                string connString = @"Data Source=" + datasource + ";Initial Catalog="
                            + database + ";Integrated Security=SSPI";

                SqlConnection conn = new SqlConnection(connString);

                return conn;
            }
        }
        public class DBUtils
        {
            public static SqlConnection GetDBConnection()
            {
                string datasource = @"laptop-j5h40kmv";

                string database = "Year2";

                return DBSQLServerUtils.GetDBConnection(datasource, database);
            }
        }

        private void Refresh_Click(object sender, EventArgs e)
        {
            string[] txt = Ubox.Text.Split('|');
            string sql = "SELECT UserID,FullName,Role From Users";
            sqldata = GetData(sql, 3);
            Ubox.Items.Clear();
            foreach (string[] x in sqldata)
            {
                Ubox.Items.Add(x[0] + "|" + x[1] + "|" + x[2]);
            }
            sqldata.Clear();
            MessageBox.Show("Don't forget to choose a user againn.");
        }

        private void Ab_Click(object sender, EventArgs e)
        {
            MessageBox.Show("For each field add information to the question. \n After adding all the needed questions press the create button.");
        }
        private void CQ_Click(object sender, EventArgs e)
        {
            Regex r = new Regex(@"[\d]*[A-z]+[\w]*[\d]*");
            if (r.IsMatch(comboBox1.Text))
            {
                bool ok = true;
                Regex r1 = new Regex(@"[\w]+[\d]*");
                Regex r2 = new Regex(@"\-*[\d]+\,*[\d]*");
                if (r1.IsMatch(textBox1.Text))
                {
                    for (int i = 0; i < qgrid.Rows.Count - 1; i++)
                    {
                        string s1 = qgrid.Rows[i].Cells[2].Value.ToString().Replace('.',',');
                        string s2 = qgrid.Rows[i].Cells[3].Value.ToString().Replace('.', ',');
                        if (r1.IsMatch(qgrid.Rows[i].Cells[1].Value.ToString()) && r1.IsMatch(qgrid.Rows[i].Cells[0].Value.ToString()) && r2.IsMatch(s1) && r2.IsMatch(s2) && Convert.ToDouble(s1) > Convert.ToDouble(s2))
                        {

                        }
                        else
                        {
                            MessageBox.Show("Please, check your input data again.");
                            ok = false;
                        }
                    }
                    if (ok)
                    {
                        string[] txt = comboBox1.Text.Split('|');
                        string sql = "Insert into Questionnaire values (" + questionnaire + ",\'" + textBox1.Text + "\'," + txt[0] + ");";
                        project++;
                        Execute(sql);
                        for (int i = 0; i < qgrid.Rows.Count - 1; i++)
                        {
                            string s1 = qgrid.Rows[i].Cells[2].Value.ToString().Replace(',', '.');
                            string s2= qgrid.Rows[i].Cells[3].Value.ToString().Replace(',', '.');
                            sql = "Insert into question values (" + question + "," + questionnaire + ",\'" + qgrid.Rows[i].Cells[1].Value.ToString() + "\'," + s1 + "," + s2 + ",\'" + qgrid.Rows[i].Cells[0].Value.ToString() + "\');";
                            Execute(sql);
                            question++;
                        }
                        questionnaire++;
                        txt = comboBox1.Text.Split('|');
                        sql = "SELECT QID,Name From Questionnaire where ProjectID=" + Convert.ToInt32(txt[0]);
                        sqldata = GetData(sql,2);
                        Acb.Items.Clear();
                        foreach (string[] x in sqldata)
                        {
                            Acb.Items.Add(x[0] + "|" + x[1]);
                        }
                        sqldata.Clear();
                    }
                    qgrid.Rows.Clear();qgrid.Columns.Clear();
                    qgrid.Columns.Add("Question Text", "Question Text");
                    qgrid.Columns.Add("Unit", "Unit");
                    qgrid.Columns.Add("Max", "Max");
                    qgrid.Columns.Add("Min", "Min");
                    qgrid.Rows.Add();
                }
                else
                {
                    MessageBox.Show("Please, check the name again.");
                }
            }
            else
            {
                MessageBox.Show("Please check if you selected a project.");
            }
        }

        private void Acb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                itbox.Items.Clear();
                Questtab.Rows.Clear(); Questtab.Columns.Clear();
                int maxi = 0;
                string sql = "Select questionid from question where qid=" + Acb.Text[0] + ";";
                sqldata = GetData(sql, 1);
                sql = "select max(iteration) from answer where questionid=" + sqldata[0][0] + ";";
                maxi = GetNum(sql, maxi) - 1;
                sqldata.Clear();
                sql = "Select questionid,unit,maximum,minimum,text from question where qid=" + Acb.Text[0] + ";";
                sqldata = GetData(sql, 5);
                Questtab.Columns.Clear();Questtab.Rows.Clear();
                Questtab.Columns.Add("Question text", "Question text");
                Questtab.Columns.Add("Unit", "Unit");
                Questtab.Columns.Add("Maximum", "Maximum");
                Questtab.Columns.Add("Minimum", "Minimum");
                foreach (string[] p in sqldata)
                {
                    Questtab.Rows.Add();Questtab.Rows[Questtab.Rows.Count - 2].HeaderCell.Value = p[0];
                    Questtab.Rows[Questtab.Rows.Count - 2].Cells[0].Value = p[4];
                    Questtab.Rows[Questtab.Rows.Count - 2].Cells[1].Value = p[1];
                    Questtab.Rows[Questtab.Rows.Count - 2].Cells[2].Value = p[2];
                    Questtab.Rows[Questtab.Rows.Count - 2].Cells[3].Value = p[3];
                }
                sqldata.Clear();
                if (maxi > 0)
                {
                    Aclb1.Items.Clear();
                    Aagrid.Rows.Clear(); Aagrid.Columns.Clear();
                    for (int i = 1; i <= maxi; i++)
                    {
                        itbox.Items.Add(i);
                    }
                    sql = "select questionid from question where qid=" + Acb.Text[0] + ";";
                    sqldata = GetData(sql,1);
                    sql = "select userid from answer where questionid=" + sqldata[0][0] + ";";
                    sqldata.Clear();
                    sqldata = GetData(sql, 1);
                    List<string[]> temp=new List<string[]>();
                    foreach (string[] p in sqldata)
                    {
                        sql = "select Fullname from users where userid=" + p[0] + ";";
                        temp = GetData(sql, 1);
                        Aclb1.Items.Add(p[0] + "|" + temp[0][0]);
                        temp.Clear();
                    }
                    for(int i = 0; i < Aclb1.Items.Count; i++)
                    {
                        Aclb1.SetItemChecked(i, true);
                    }
                    sqldata.Clear();
                } 
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void Aclb1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Questtab.Rows.Clear(); Questtab.Columns.Clear();Aagrid.Rows.Clear();Aagrid.Columns.Clear();
            string[] txt = Acb.Text.Split('|');
            string sql = "delete from questionnaire where QID=" + txt[0] + ";";
            project++;
            Execute(sql);
            sql= "delete from selfmark where QID = " + txt[0] + "; ";
            Execute(sql);
            txt = comboBox1.Text.Split('|');
            sql = "SELECT QID,Name From Questionnaire where ProjectID=" + Convert.ToInt32(txt[0]);
            sqldata = GetData(sql, 2);
            Acb.Items.Clear();
            foreach (string[] x in sqldata)
            {
                Acb.Items.Add(x[0] + "|" + x[1]);
            }
            sqldata.Clear();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Questtab.Rows.Clear(); Questtab.Columns.Clear();
                string[] txt = comboBox1.Text.Split('|');
                string sql = "SELECT QID,Name From Questionnaire where ProjectID=" + Convert.ToInt32(txt[0]);
                Acb.Items.Clear();
                sqldata = GetData(sql, 2);
                foreach (string[] x in sqldata)
                {
                    Acb.Items.Add(x[0] + "|" + x[1]);
                }
                itbox.Items.Clear();
                sqldata.Clear();
            }
            catch {
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Regex r = new Regex(@"[\d]*[\w]+[\d]*");
            if (r.IsMatch(comboBox1.Text))
            {
                Questtab.Rows.Clear();Questtab.Columns.Clear();Aagrid.Rows.Clear();Aagrid.Columns.Clear();
                string[] txt = comboBox1.Text.Split('|');
                string sql = "select qid from questionnaire where projectid=" + txt[0] + ";";
                sqldata = GetData(sql, 1);
                foreach(string[] x in sqldata)
                {
                    sql = "delete from selfmark where qid=" + x[0] + ";";
                    Execute(sql);
                }
                sqldata.Clear();
                sql = "delete from project where projectid=" + txt[0] + ";";
                Execute(sql);
                FullUpdate();
            }
            else
            {
                MessageBox.Show("Please choose a project.");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            textBox2.Visible = true;
            button7.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Regex r = new Regex(@"[\d]*[A-z]+[\w]*[\d]*");
            if (r.IsMatch(textBox2.Text))
            {
                SqlCommand cmd = new SqlCommand();
                string sql = "select Max(ProjectID) from Project;";
                project=GetNum(sql, 0);
                sql = "insert into project values (" + project + ",\'" + textBox2.Text + "\');";
                Execute(sql);
                label3.Visible = false;
                textBox2.Visible = false;
                button7.Visible = false;
                sql = "select * from project;";
                sqldata = GetData(sql, 2); 
                comboBox1.Items.Clear();
                foreach (string[] x in sqldata)
                {
                    comboBox1.Items.Add(x[0] + "|" + x[1]);
                }
                sqldata.Clear();
            }
            else
            {
                MessageBox.Show("Please check the project name again. The name must have at least one letter.");
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                string[] txt = Acb.Text.Split('|');
                foreach (string s in Aclb1.CheckedItems)
                {
                    string[] txt1 = s.Split('|');
                    string sql = "insert into Involvment values (" + txt[0] + "," + txt1[0] + "," + inv + ");";
                        inv++;
                    Execute(sql);
                }
                MessageBox.Show("Successful.");
            }
            catch
            {
                MessageBox.Show("Make sure you chose both experts and a questionnaire.");
            }
        }

        private void Ecb_SelectedIndexChanged(object sender, EventArgs e)
        {
            FullUpdate();
        }

        private void Sb_Click(object sender, EventArgs e)
        {
            try  {
                other.Rows.Clear();other.Columns.Clear();
                Regex reg = new Regex(@"\-*[\d]+\.*[\d]*");
                bool o = true;
                for(int i = 0; i < Eagrid.Rows.Count - 1; i++)
                {
                    string s = Eagrid.Rows[i].Cells[0].Value.ToString().Replace(',', '.');
                    if (!reg.IsMatch(Eagrid.Rows[i].Cells[0].Value.ToString()))
                    {
                        o = false;
                    }
                    else
                    {
                        s = s.Replace('.', ',');string s1 = Eqgrid.Rows[i].Cells[2].Value.ToString().Replace('.', ',');string s2 = Eqgrid.Rows[i].Cells[3].Value.ToString().Replace('.', ',');
                        if (Convert.ToDouble(s) > Convert.ToDouble(s1) || Convert.ToDouble(s) < Convert.ToDouble(s2))
                            o = false;
                    }
                }
                if (o)
                {
                    Sb.Visible = false;
                    string qid;
                    string sql = "select max(smid) from selfmark;";
                    sm=GetNum(sql, sm);
                    sql = "select Max(iteration) from answer where (userid=" + Ubox.Text[0] + " and questionid=" + Eagrid.Rows[0].HeaderCell.Value + ");";
                    it=GetNum(sql, it);
                    if (it == 1)
                    {
                        qid = Ecb.Text.Trim(' ');
                        qid = qid[0].ToString();
                        sql = "insert into selfmark values(" + sm + "," + Ubox.Text[0] + "," + etb.Text + "," + qid + ");";
                        Execute(sql);
                    }
                    for (int i = 0; i < Eagrid.Rows.Count; i++)
                    {
                        sql = "select Max(answerid) from answer;";
                        an=GetNum(sql, an);
                        sql = "insert into answer values (" + an + "," + Convert.ToDouble(Eagrid.Rows[i].Cells[0].Value.ToString().Replace('.',',')) + "," + Eagrid.Rows[i].HeaderCell.Value + "," + Ubox.Text[0] + "," + it + ",\'" + Eagrid.Rows[i].Cells[1].Value + "\');";
                        Execute(sql);
                    }
                    El3.Visible = false;
                    etb.Visible = false;
                    MessageBox.Show("Successfull.");
                    sql = "delete from involvment where (uid=" + Ubox.Text[0] + " and qid=(select qid from question where questionid=" + Eagrid.Rows[0].HeaderCell.Value + "));";
                    Execute(sql);
                    string[] txt = Ubox.Text.Split('|');
                    sql = "SELECT QID from Involvment where UID=" + txt[0] + ";";
                    sqldata = GetData(sql, 1);
                    Ecb.Items.Clear();
                    foreach (string[] x in sqldata)
                    {
                        sql = "Select QID,Name from questionnaire where qid=" + x[0] + ";";
                        List<string[]> tem = GetData(sql, 2);
                        foreach (string[] y in tem)
                        {
                            Ecb.Items.Add(y[0] + "|" + y[1]);
                        }

                    }
                    sqldata.Clear();
                    etb.Text = "";
                    Eqgrid.Rows.Clear(); Eqgrid.Columns.Clear();
                    Eagrid.Rows.Clear(); Eagrid.Columns.Clear();
                    El3.Visible = false;
                    etb.Visible = false;
                }
                else
                {
                    MessageBox.Show("Please make sure all of your answers are correct.");
                }
            }

           catch
            {
                MessageBox.Show("Make sure you answered all the questions correctly and you wrote a selfmark for yourself.");
            }
        }

        private void itbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int c = 0;
            Aagrid.Rows.Clear(); Aagrid.Columns.Clear();
            Aagrid.Columns.Add("Answer", "Answer");
            Aagrid.Columns.Add("Explanation", "Explanation");
            Aagrid.Columns.Add("User", "User");
            string sql = "Select questionid from question where qid=" + Acb.Text[0] + ";";
            sqldata = GetData(sql, 1);
            foreach (string[] x in sqldata)
            {
                sql = "select answer,userid,explanation from answer where (questionid=" + x[0] + " and iteration=" + itbox.Text + ");";
                List<string[]> dat = new List<string[]>();
                try
                {
                    dat = GetData(sql, 3);
                }
                catch
                {

                }
                if (dat.Count > 0)
                {
                    foreach (string[] y in dat)
                    {
                        Aagrid.Rows.Add();
                        Aagrid.Rows[c].Cells[0].Value = y[0];
                        Aagrid.Rows[c].Cells[1].Value = y[2];
                        Aagrid.Rows[c].HeaderCell.Value = x[0];
                        sql = "select FullName from users where userid=" + y[1] + ";";
                        var cmd = new SqlCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = sql;
                        SqlDataReader reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            Aagrid.Rows[c].Cells[2].Value = reader[0].ToString();
                        }
                        reader.Close();
                        c++;
                    }
                }
            }
            sqldata.Clear();
        }

        private void Ab1_Click(object sender, EventArgs e)
        {
            try
            {
                double Median, SimpleMark, MidGroup; string Interval;
                string sql = "select value from selfmark where qid=" + Acb.Text[0] + ";";
                sqldata = GetData(sql, 1);
                MidGroup = 0;
                foreach(string[] x in sqldata)
                {
                    MidGroup = MidGroup + Convert.ToDouble(x[0]);
                }
                MidGroup = MidGroup / sqldata.Count;sqldata.Clear();
                SimpleMark = 0;
                double[] grades=new double[Aagrid.Rows.Count-1];
                double min = Double.MaxValue, max = Double.MinValue;
                for(int i = 0; i < Aagrid.Rows.Count - 1; i++)
                {
                    SimpleMark = SimpleMark + Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value);
                    grades[i] = Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value);
                    if (max < Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value))
                        max = Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value);
                    if (min > Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value))
                        min = Convert.ToDouble(Aagrid.Rows[i].Cells[0].Value);
                }
                SimpleMark = SimpleMark / (Aagrid.Rows.Count - 1);
                for (int i = 0; i < grades.Length; i++)
                {
                    for (int j = 0; j < grades.Length - 1; j++)
                    {
                        if (grades[j] > grades[j + 1])
                        {
                            double z = grades[j];
                            grades[j] = grades[j + 1];
                            grades[j + 1] = z;
                        }
                    }
                }
                if (grades.Length % 2 == 0)
                {
                    Median = (grades[grades.Length / 2-1] + grades[grades.Length / 2])/2;
                }
                else
                {
                    Median = (grades[Convert.ToInt32(grades.Length / 2)] * 2)/2;
                }
                double temp = (max - min) / 4;
                max = max - temp;min = min + temp;
                Interval="["+min+";"+max+"]";
                Art.Text = "Midgroup mark: " + MidGroup + "\n Simple mark: " + SimpleMark + "\n Median: " + Median + "\n Confidence interval: " + Interval;
                sql = "select count(resultid) from result where qid=" + Acb.Text[0] + " and iteration=" + itbox.Text[0]+";";
                int num1 = 0;
                sqldata = GetData(sql, 1);
                num1 = Convert.ToInt32(sqldata[0][0]);
                sqldata.Clear();
                sql = "select Max(resultid) from result;";
                res=GetNum(sql, res);
                if (num1 == 0)
                {
                    string median=Median.ToString(), midgroup=MidGroup.ToString(), simplemark=SimpleMark.ToString();
                    if (Median.ToString().Contains(',')) {
                        median = Median.ToString().Replace(',', '.');
                    }
                    if (MidGroup.ToString().Contains(',')) {
                        midgroup = MidGroup.ToString().Replace(',', '.');
                    }
                    if (SimpleMark.ToString().Contains(',')){
                        simplemark = SimpleMark.ToString().Replace(',', '.');
                            }
                    if (Interval.Contains(',')){
                        Interval = Interval.Replace(',', '.');
                    }
                    sql = "Insert into result values(" + res + "," + Acb.Text[0] + "," + median + "," +midgroup + "," + itbox.Text[0] + ",\'" + Interval + "\'," + simplemark + ");";
                    Execute(sql);
                }
            }
            catch
            {
                MessageBox.Show("Make sure you have chosen both questionnaire and an iteration.");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
        }

        private void A1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Press the \"Add\" button to open a field for new project creation. \n Press the \"Delete\" button to delete the chosen project from the memory.");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Press the \"Delete\" button to delete the chosen questionnaire from the memory.");

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("After selecting a questionnaire choose an existing iteration to see the answers. Press the \"Get results\" button to calculate all necessary parameters");
        }

        private void A3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("After choosing a questionnaire select all experts, who will participate in the process and press the \"Send\" button.");
        }

        private void A4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Choose a questionnaire to see it, and to open a form for writing answers.");
        }

        private void A5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The bigger field on the top shows all the questions in selected questionnaire. \n The bottom left field is used for writing the answers to questions. \n The bottom right field shows the answers of other experts, participating in working with selected questionnaire");
        }

        private void Eqgrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

    

