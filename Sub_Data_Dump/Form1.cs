using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Sub_Data_Dump
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string conn = "Provider=Microsoft.Jet.OLEDB.4.0;";
        #region My Functions
        private void listClear()
        {
            this.listBox2.Items.Clear();
            this.listBox4.Items.Clear();
            this.listBox6.Items.Clear();
            this.listBox7.Items.Clear();
            this.listBox8.Items.Clear();
            this.listBox9.Items.Clear();
            this.listBox10.Items.Clear();
            this.listBox11.Items.Clear();
        }
        public void get_tables()
        {
            string s1 = "Data Source=" + this.textBox1.Text;
            if (this.textBox1.Text.Contains(".accdb"))
            {
                conn = "Provider=Microsoft.ACE.OLEDB.12.0;";
            }
            else
            {
                conn = "Provider=Microsoft.Jet.OLEDB.4.0;";
            }

            OleDbConnection con = new OleDbConnection(conn + s1);
            try
            {
                con.Open();
                DataTable tables = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow dr in tables.Rows)
                {
                    string tablename = dr[2].ToString();
                    this.listBox1.Items.Add(tablename);
                }
            }
            catch
            {
                this.textBox2.Text = "Error in running Query or Invalid File Name \n";
            }
        }

        public void count_Rows()
        {
            string s1 = "Data Source=" + this.textBox1.Text;
            //string conn = "Provider=Microsoft.Jet.OLEDB.4.0;" + s1;
            //OleDbConnection con = new OleDbConnection(conn + s1);
            string qry = "select count(*) from [" + this.listBox1.SelectedItem.ToString() + "]";
            //this.richTextBox1.Text = qry;
            try
            {
                using (OleDbConnection conn1 = new OleDbConnection(conn + s1))
                using (OleDbCommand command = new OleDbCommand(qry, conn1))
                {
                    conn1.Open();
                    int count = (int)command.ExecuteScalar();
                    this.textBox4.Text = count.ToString();
                }
            }
            catch
            {
                this.textBox2.Text = "Error in counting records ";
            }

        }

        public static string NormalizeWhiteSpace(string S)
        {
            string s = S.Trim();
            bool iswhite = false;
            int sLength = s.Length;
            StringBuilder sb = new StringBuilder(sLength);
            foreach (char c in s.ToCharArray())
            {
                if (Char.IsWhiteSpace(c))
                {
                    if (iswhite)
                    {
                        //Continuing whitespace ignore it.
                        continue;
                    }
                    else
                    {
                        //New WhiteSpace

                        //Replace whitespace with a single space.
                        sb.Append(" ");
                        //Set iswhite to True and any following whitespace will be ignored
                        iswhite = true;
                    }
                }
                else
                {
                    sb.Append(c.ToString());
                    //reset iswhitespace to false
                    iswhite = false;
                }
            }
            return sb.ToString();
        }


        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            this.comboBox1.SelectedIndex = 0;
            this.listBox3.SelectedIndex = 0;
            this.button3.Enabled = false;
            this.button4.Enabled = false;
            this.textBox4.ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Database file";
            dlg.Filter = "Access 2007 (*.accdb)|*.accdb|Access 2003 (*.mdb)|*.mdb";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = dlg.FileName;
                get_tables();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listClear();
            string s1 = "Data Source=" + this.textBox1.Text;
            OleDbConnection con = new OleDbConnection(conn + s1);
            try
            {
                con.Open();
                string qry = "select top 5 * from " + "[" + this.listBox1.SelectedItem.ToString() + "]";
                using (var cmd = new OleDbCommand(qry, con))
                using (var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly))
                {
                    var table = reader.GetSchemaTable();
                    var nameCol = table.Columns["ColumnName"];
                    foreach (DataRow row in table.Rows)
                    {
                        string col = row[nameCol].ToString();
                        this.listBox2.Items.Add(col);
                    }
                }
                DataSet mydataset = new DataSet();
                OleDbCommand cmd1 = new OleDbCommand(qry, con);
                OleDbDataAdapter ada = new OleDbDataAdapter(cmd1);
                string[] restrictions1 = new string[4] { null, null, this.listBox1.SelectedItem.ToString(), null };
                DataTable dt = con.GetSchema("Columns", restrictions1);
                DataSet ds = new DataSet();
                DataTable t = new DataTable();
                //this.textBox2.Text = qry + '\n' + this.textBox2.Text;
                textBox2.AppendText(qry + "\n");
                this.label4.Text = "Default View of Table :" + this.listBox1.SelectedItem.ToString();
                ds.Tables.Add(t);
                ada.Fill(t);
                dataGridView1.DataSource = t.DefaultView;
                con.Close();
                count_Rows();
            }
            catch
            {
                this.textBox2.Text += "Failed to open the table\n";
            }
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // List Box Having Row Names
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox2.SelectedIndex != -1)
            {
                string s1 = this.listBox2.SelectedItem.ToString();
                int n1 = this.listBox3.SelectedIndex;
                switch (n1)
                {
                    case 0:
                        {
                            this.listBox4.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 1:
                        {
                            this.listBox5.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 2:
                        {
                            this.listBox6.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 3:
                        {
                            this.listBox7.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 4:
                        {
                            this.listBox8.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 5:
                        {
                            this.listBox9.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 6:
                        {
                            this.listBox10.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    case 7:
                        {
                            this.listBox11.Items.Add(s1);
                            this.listBox2.Items.Remove(s1);
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        #region listboxes
        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox4.SelectedIndex != -1)
            {
                string s1 = this.listBox4.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox4.Items.Remove(s1);
            }
        }

        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox5.SelectedIndex != -1)
            {
                string s1 = this.listBox5.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox5.Items.Remove(s1);
            }
        }

        private void listBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox6.SelectedIndex != -1)
            {
                string s1 = this.listBox6.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox6.Items.Remove(s1);
            }
        }

        private void listBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox7.SelectedIndex != -1)
            {
                string s1 = this.listBox7.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox7.Items.Remove(s1);
            }
        }

        private void listBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox8.SelectedIndex != -1)
            {
                string s1 = this.listBox8.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox8.Items.Remove(s1);
            }
        }

        private void listBox9_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (this.listBox9.SelectedIndex != -1)
            {
                string s1 = this.listBox9.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox9.Items.Remove(s1);
            }
        }

        private void listBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox10.SelectedIndex != -1)
            {
                string s1 = this.listBox10.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox10.Items.Remove(s1);
            }
        }

        private void listBox11_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (this.listBox11.SelectedIndex != -1)
            {
                string s1 = this.listBox11.SelectedItem.ToString();
                this.listBox2.Items.Add(s1);
                this.listBox11.Items.Remove(s1);
            }
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            string qry = "select  top 5";
            string add = " ";
            # region listbox check
            if (this.listBox4.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox4.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox4.Items[i].ToString() + "]";
                }
                add += " as [Mobile No],";
            }
            if (this.listBox5.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox5.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox5.Items[i].ToString() + "]";
                }
                add += " as [Name],";
            }
            if (this.listBox6.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox6.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox6.Items[i].ToString() + "]";
                }
                add += " as [Father's Name],";
            }
            if (this.listBox7.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox7.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox7.Items[i].ToString() + "]";
                }
                add += " as [Address],";
            }
            if (this.listBox8.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox8.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox8.Items[i].ToString() + "]";
                }
                add += " as [SIM Activation Date],";
            }
            if (this.listBox9.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox9.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox9.Items[i].ToString() + "]";
                }
                add += " as [POI Number],";
            }
            if (this.listBox10.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox10.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox10.Items[i].ToString() + "]";
                }
                add += " as [POA Number],";
            }
            if (this.listBox11.Items.Count > 0)
            {
                //add += " (";
                for (int i = 0; i < this.listBox11.Items.Count; i++)
                {
                    if (i > 0)
                    {
                        add += "&\"  \"&";
                    }
                    //add += "ISNULL(";
                    add += "[" + this.listBox11.Items[i].ToString() + "]";
                }
                add += " as [POS Code],";
            }
            #endregion

            string a1 = add.Substring(0, add.Length - 1);
            add = a1 + " from [" + this.listBox1.SelectedItem.ToString() + "];";
            qry += add;
            this.textBox3.Text = qry;
            this.button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string s1 = "Data Source=" + this.textBox1.Text;
            OleDbConnection con = new OleDbConnection(conn + s1);
            try
            {
                con.Open();
                string qry = this.textBox3.Text;
                DataSet mydataset = new DataSet();
                OleDbCommand cmd1 = new OleDbCommand(qry, con);
                OleDbDataAdapter ada = new OleDbDataAdapter(cmd1);

                DataSet ds = new DataSet();
                DataTable t = new DataTable();
                //this.textBox2.Text = qry + '\n' + this.textBox2.Text;               
                ds.Tables.Add(t);
                ada.Fill(t);
                dataGridView2.DataSource = t.DefaultView;
                con.Close();
                this.button4.Enabled = true;
            }
            catch
            {
                this.textBox2.Text += "Failed to Test Query!!!!!!!!!!!!!!\n";
            }
        }
        # region Bgworker_Dump
        public int ReadTheTable(string fileName,string tableName,string opFileName)
        {
            int numLines = 0;

            

            return numLines;
        }

        #endregion
        private void dump_query()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            Int64 x2 = Convert.ToInt64(this.textBox4.Text.ToString());
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox5.Text = saveFileDialog1.FileName;
                string save_file = saveFileDialog1.FileName;
                try
                {
                    StreamWriter sw = new StreamWriter(this.textBox5.Text);
                    string head = "Mobile_No|Name|Father's_Name|Address|Activation_Date|POI_No|POA_No|POS_Code|TSP";
                    string qry = this.textBox3.Text.ToString();


                    string s1 = "Data Source=" + this.textBox1.Text;


                    OleDbConnection con = new OleDbConnection(conn + s1);
                    try
                    {
                        con.Open();
                        sw.WriteLine(head);

                        DataSet mydataset = new DataSet();
                        OleDbCommand cmd = new OleDbCommand(qry, con);
                        OleDbDataAdapter ada = new OleDbDataAdapter(cmd);
                        OleDbDataReader reader = cmd.ExecuteReader();
                        Int64 x = 0;
                        double bar = 0;
                        double divder = Convert.ToDouble(x2);
                        while (reader.Read())
                        {
                            string data = "";
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (i > 0)
                                {
                                    data += "|";
                                }
                                data += reader.GetValue(i).ToString();
                            }
                            string tail = "|" + comboBox1.SelectedItem.ToString();
                            data = NormalizeWhiteSpace(data);
                            sw.WriteLine(data + tail);
                            x += 1;
                            progressBar1.Value = Convert.ToInt16(Convert.ToDouble(x) / divder * 100);

                        }
                        this.textBox2.Text = "Dumping Complete";
                        con.Close();

                    }
                    catch (Exception ex)
                    {
                        this.textBox2.Text = "Failed to Dump data :=( " + ex.ToString();
                    }
                    sw.Close();
                }
                catch
                {
                    this.textBox2.Text = "Something wrong";
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dump_query();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
