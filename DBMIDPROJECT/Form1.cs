using DBMIDPROJECT.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using System.Data.SqlClient;
using System.Runtime.Remoting.Lifetime;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Xml.Schema;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.Remoting.Contexts;
//using CrystalDecisions.CrystalReports.Engine;


namespace DBMIDPROJECT
{
    public partial class Form1 : Form
    {
        Panel current = new Panel();
        int count = 1;
        int id = 0;
        string frm = "";
        List<int> atten = new List<int>();
        bool mark = false;
        int reportcount = 0;
        DateTime datetime1 = new DateTime();
        bool view = false;
        public Form1()
        {
            InitializeComponent();
            hidepanels();
            current = panel47;
            current.Show();
            RetrieveCLO(0);
            Retrieve(0);
            RetrieveRubrics(0);
            RetrieveRubricsLevel(0);
            RetrieveAsseement(0);
            RetrieveAssessmentComponent(0);
            RetrieveInactive(0);
            RetrieveStudentForEvaluation();
            RetrieveStudentResult(0);
            button9.Text = "View";
            dataGridView1.AllowUserToAddRows= false;
            dataGridView1.AllowUserToDeleteRows= false;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;
            dataGridView3.AllowUserToAddRows = false;
            dataGridView3.AllowUserToDeleteRows = false;
            dataGridView4.AllowUserToAddRows = false;
            dataGridView4.AllowUserToDeleteRows = false;
            dataGridView5.AllowUserToAddRows = false;
            dataGridView5.AllowUserToDeleteRows = false;
            dataGridView6.AllowUserToAddRows = false;
            dataGridView6.AllowUserToDeleteRows = false;
            dataGridView7.AllowUserToAddRows = false;
            dataGridView7.AllowUserToDeleteRows = false;
            dataGridView8.AllowUserToAddRows = false;
            dataGridView8.AllowUserToDeleteRows = false;
            dataGridView9.AllowUserToAddRows = false;
            dataGridView9.AllowUserToDeleteRows = false;
            dataGridView10.AllowUserToAddRows = false;
            dataGridView10.AllowUserToDeleteRows = false;
            dataGridView11.AllowUserToAddRows = false;
            dataGridView11.AllowUserToDeleteRows = false;
            comboBox9.Text = "Rubric Id";
            Dashboard();
            comboBox1.Items.Remove("Id");
            comboBox2.Items.Remove("Id");
            //comboBox3.Items.Remove("Id");
            comboBox5.Items.Remove("Id");
            comboBox7.Items.Remove("Id");
            comboBox13.Items.Remove("Id");
            //comboBox8.Items.Remove("Id");
            comboBox16.Items.Add("Student Wise Clo Report");
            comboBox16.Items.Add("Student Wise Assessment Report");
        }
        private void Dashboard()
        {
            var con = Configuration.getInstance().getConnection();
            SqlCommand temp = new SqlCommand("SELECT count(*) FROM Student", con);
            Int32 count = (Int32)temp.ExecuteScalar();
            label29.Text = count.ToString();
            SqlCommand temp1 = new SqlCommand("SELECT count(*) FROM Student Where Status = 5", con);
            Int32 count1 = (Int32)temp1.ExecuteScalar();
            label40.Text = count1.ToString();
            SqlCommand temp2 = new SqlCommand("SELECT count(*) FROM Student Where Status = 6", con);
            Int32 count2 = (Int32)temp2.ExecuteScalar();
            label36.Text = count2.ToString();
            SqlCommand temp3 = new SqlCommand("SELECT count(*) FROM Clo", con);
            Int32 count3 = (Int32)temp3.ExecuteScalar();
            label39.Text = count3.ToString();
            SqlCommand temp4 = new SqlCommand("SELECT count(*) FROM Assessment", con);
            Int32 count4 = (Int32)temp4.ExecuteScalar();
            label37.Text = count4.ToString();
            SqlCommand temp5 = new SqlCommand("SELECT count(*) FROM Rubric", con);
            Int32 count5 = (Int32)temp5.ExecuteScalar();
            label42.Text = count5.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (count % 2 == 0) 
            {
                panel1.Width = 50;
            }
            else 
            {
                panel1.Width = 200;
            }
            count++;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            id = 0;
            pictureBox1.BackgroundImage = Resources.icons8_student_center_60;
            pictureBox2.Enabled = false;
            //hidepanels();
            Retrieve(1);
            current.Hide();
            current = panel2;
            current.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            pictureBox5.Enabled = false; 
            id = 0; 
            pictureBox1.BackgroundImage = Resources.icons8_course_assign_601;
            //hidepanels();
            RetrieveCLO(1);
            current.Hide();
            current = panel5;
            current.Show();
            //panel5.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            id = 0;
            pictureBox7.Enabled = false;
            pictureBox1.BackgroundImage = Resources.south_african_rand__1_;
            //hidepanels();
            additemsinCombo4();
            RetrieveRubrics(1);
            current.Hide();
            current = panel11;
            current.Show();
            //panel11.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            id = 0;
            pictureBox10.Enabled = false;
            pictureBox1.BackgroundImage = Resources.icons8_todo_list_48;
            //hidepanels();
            RetrieveAsseement(1);
            current.Hide();
            current = panel21;
            current.Show();
            //panel21.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            id = 0;
            pictureBox9.Enabled = false;
            pictureBox1.BackgroundImage = Resources.icons8_stairs_up_60;
            //hidepanels();
            rubricitemcombo(); 
            rubricitemcombo1();
            RetrieveRubricsLevel(1);
            //panel16.Show();
            current.Hide();
            current = panel16;
            current.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            id = 0; 
            pictureBox1.BackgroundImage = Resources.icons8_exam_60;
            //hidepanels();
            RetrieveStudentForEvaluation();
            RetrieveStudentResult(1);
            //panel38.Show();
            current.Hide();
            current = panel38;
            current.Show();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
        private void hidepanels() 
        {
            panel2.Hide();
            panel5.Hide();
            panel11.Hide();
            panel16.Hide();
            panel21.Hide();
            panel26.Hide();
            panel31.Hide();
            panel36.Hide();
            panel38.Hide();
            panel45.Hide();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox4.Text != "" && textBox5.Text != "")
            {
                if (!Regex.IsMatch(textBox1.Text, @"^[\p{L}\p{M}' \.\-]+$") || (textBox1.Text.ToString()).Count(c => c == ' ') == (textBox1.Text.ToString()).Length)
                {
                    MessageBox.Show("Invalid First Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.IsMatch(textBox2.Text, @"^[\p{L}\p{M}' \.\-]+$") && textBox2.Text!="")
                {
                    MessageBox.Show("Invalid Last Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.Match(textBox3.Text, "\\d{3}-\\d{3}-\\d{7}").Success && textBox3.Text != "")
                {
                    MessageBox.Show("Correct Format is 000-000-0000000", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!IsValidEmail(textBox4.Text))
                {
                    MessageBox.Show("Invalid Email", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.IsMatch(textBox5.Text, @"^(?=.{5,20}$)\d{4}-[^ ]{1,14}-\d{1,3}$"))
                {
                    MessageBox.Show("Registration Format is Session-Department-RollNo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp = new SqlCommand("SELECT count(*) FROM Student Where RegistrationNumber = @RegistrationNumber", con);
                        temp.Parameters.AddWithValue("@RegistrationNumber", textBox5.Text);
                        Int32 count = (Int32)temp.ExecuteScalar();
                        SqlCommand temp1 = new SqlCommand("SELECT id FROM Student Where Email = @RegistrationNumber", con);
                        temp1.Parameters.AddWithValue("@RegistrationNumber", textBox4.Text);
                        object temp2 = (object)temp1.ExecuteScalar();
                        if (temp2 == null) {
                            if (count == 0)
                            {
                                SqlCommand cmd = new SqlCommand("Insert into Student values (@FirstName, @LastName, @Contact,@Email,@RegistrationNumber,@Status)", con);
                                cmd.Parameters.AddWithValue("@FirstName", textBox1.Text);
                                cmd.Parameters.AddWithValue("@LastName", textBox2.Text);
                                cmd.Parameters.AddWithValue("@Contact", textBox3.Text);
                                cmd.Parameters.AddWithValue("@Email", textBox4.Text);
                                cmd.Parameters.AddWithValue("@RegistrationNumber", textBox5.Text);
                                cmd.Parameters.AddWithValue("@Status", 5);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Successfully saved");
                                Cleartextboxes();
                                Retrieve(1);
                                id = 0;
                                pictureBox2.Enabled = false;
                            }
                            else
                            {
                                MessageBox.Show("Student Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Email Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                string pattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";

                Regex regex = new Regex(pattern);
                Match match = regex.Match(email);
                return match.Success;
            }
            catch
            {
                return false;
            }
        }
        private void Retrieve(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("Select FirstName,LastName,Contact,Email,RegistrationNumber from Student Where Status = @Status", con);
                cmd.Parameters.AddWithValue("@Status", 5);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView1.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error");
            }
        }
        private void Cleartextboxes()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox4.Text != "" && textBox5.Text != "")
            {
                if (!Regex.IsMatch(textBox1.Text, @"^[\p{L}\p{M}' \.\-]+$") || (textBox1.Text.ToString()).Count(c => c == ' ') == (textBox1.Text.ToString()).Length)
                {
                    MessageBox.Show("Invalid First Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.IsMatch(textBox2.Text, @"^[\p{L}\p{M}' \.\-]+$") && textBox2.Text != "")
                {
                    MessageBox.Show("Invalid Last Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.Match(textBox3.Text, "\\d{3}-\\d{3}-\\d{7}").Success && textBox3.Text != "")
                {
                    MessageBox.Show("Correct Format is 000-000-0000000", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!IsValidEmail(textBox4.Text))
                {
                    MessageBox.Show("Invalid Email", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (!Regex.IsMatch(textBox5.Text, @"^(?=.{5,20}$)\d{4}-[^ ]{1,14}-\d{1,3}$"))
                {
                    MessageBox.Show("Registration Format is Session-Department-RollNo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp = new SqlCommand("SELECT count(*) FROM Student Where RegistrationNumber = @RegistrationNumber and Id != " + id, con);
                        temp.Parameters.AddWithValue("@RegistrationNumber", textBox5.Text);
                        Int32 count = (Int32)temp.ExecuteScalar();
                        SqlCommand temp1 = new SqlCommand("SELECT id FROM Student Where Email = @RegistrationNumber and Id != "+ id, con);
                        temp1.Parameters.AddWithValue("@RegistrationNumber", textBox4.Text);
                        object temp2 = (object)temp1.ExecuteScalar();
                        if (temp2 == null)
                        {
                            if(count == 0) {
                                if (id != 0)
                                {
                                    SqlCommand cmd = new SqlCommand("Update Student Set FirstName=@FirstName,LastName=@LastName,Contact=@Contact,Email=@Email,RegistrationNumber=@RegistrationNumber where Id in (SELECT Id from Student Join Lookup On Status = LookupId Where Name = 'Active' and Id = " + id + ");", con);
                                    cmd.Parameters.AddWithValue("@FirstName", textBox1.Text);
                                    cmd.Parameters.AddWithValue("@LastName", textBox2.Text);
                                    cmd.Parameters.AddWithValue("@Contact", textBox3.Text);
                                    cmd.Parameters.AddWithValue("@Email", textBox4.Text);
                                    cmd.Parameters.AddWithValue("@RegistrationNumber", textBox5.Text);
                                    cmd.Parameters.AddWithValue("@id", id);
                                    cmd.ExecuteNonQuery();
                                    MessageBox.Show("Successfully Updated");
                                    id = 0;
                                    Cleartextboxes();
                                    Retrieve(1);
                                    pictureBox2.Enabled = false;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Student Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Email Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                if (textBox6.Text == "")
                {
                    SqlCommand abc = new SqlCommand("Select FirstName,LastName,Contact,Email,RegistrationNumber from Student Where Status = @Status", con);
                    SqlDataAdapter da = new SqlDataAdapter(abc);
                    abc.Parameters.AddWithValue("@Status", 5);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
                else
                {
                    if (comboBox1.Text != "Select Attribute")
                    {
                        SqlCommand cmd = new SqlCommand("Select FirstName,LastName,Contact,Email,RegistrationNumber from Student Where " + comboBox1.Text + " like '%" + textBox6.Text + "%' and Status = 5", con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        cmd.Parameters.AddWithValue("@Combo", comboBox1.Text);
                        cmd.Parameters.AddWithValue("@Text", textBox6.Text);
                        cmd.Parameters.AddWithValue("@Status", 5);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;
                    }
                    else
                    {
                        MessageBox.Show("Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox6.Clear();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select Id From Student Where FirstName = '" + this.dataGridView1.CurrentRow.Cells[1].Value.ToString() + "' and LastName = '" + this.dataGridView1.CurrentRow.Cells[2].Value.ToString() + "' and Email = '" + this.dataGridView1.CurrentRow.Cells[4].Value.ToString() + "' and RegistrationNumber = '" + this.dataGridView1.CurrentRow.Cells[5].Value.ToString() + "'", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView1.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Are You Sure You Want to Delete", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand cmd = new SqlCommand("Update Student Set Status = 6 where Id in (SELECT Id from Student Join Lookup On Status = LookupId Where Name = 'Active' and Id = " + id + ");", con);
                        //MessageBox.Show("Update Student Set Status = 6 where Id in (SELECT Id from Student Join Lookup On Status = LookupId Where Name = 'Active' and Id = "+id+");");
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        Retrieve(1);
                        id = 0;
                        pictureBox2.Enabled = false ;
                    }
                    

                }
                    catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView1.Columns["Delete"].Index)
            {
                try
                {
                    textBox1.Text = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    textBox2.Text = this.dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    textBox3.Text = this.dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    textBox4.Text = this.dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    textBox5.Text = this.dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    pictureBox2.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {
            if (textBox11.Text != "")
            {
                if (textBox11.Text.All(x => char.IsLetterOrDigit(x) || char.IsWhiteSpace(x)) && (textBox11.Text.ToString()).Count(c => c == ' ') != (textBox11.Text.ToString()).Length)
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp = new SqlCommand("SELECT count(*) FROM Clo Where Name = @Name", con);
                        temp.Parameters.AddWithValue("@Name", textBox11.Text);
                        Int32 count = (Int32)temp.ExecuteScalar();
                        if (count == 0)
                        {
                            SqlCommand cmd = new SqlCommand("Insert into Clo values (@Name, @DateCreated,@DateUpdated)", con);
                            cmd.Parameters.AddWithValue("@Name", textBox11.Text);
                            cmd.Parameters.AddWithValue("@DateCreated", DateTime.Now);
                            cmd.Parameters.AddWithValue("@DateUpdated", DateTime.Now);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Successfully saved");
                            textBox11.Clear();
                            RetrieveCLO(1);
                            id = 0;
                            pictureBox5.Enabled = false;
                        }
                        else
                        {
                            MessageBox.Show("CLO Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void RetrieveCLO(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("Select Name,DateCreated,DateUpdated from Clo", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true; 
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView2.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select Id From Clo Where Name = '" + this.dataGridView2.CurrentRow.Cells[1].Value.ToString() + "'", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView2.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Corressponding Rubric,AssessmentComponent,RubricLevel,StudentResult Delete with the clo", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp3 = new SqlCommand("DELETE FROM StudentResult Where [RubricMeasurementId] IN (SELECT Id from RubricLevel where RubricId IN (Select Id from Rubric where CloId = " + id + "))", con);
                        temp3.ExecuteNonQuery();
                        SqlCommand temp1 = new SqlCommand("Delete from RubricLevel where RubricId IN (Select Id from Rubric where CloId = " + id + ")", con);
                        temp1.ExecuteNonQuery();
                        SqlCommand temp2 = new SqlCommand("Delete from AssessmentComponent where RubricId IN (Select Id from Rubric where CloId = " + id + ")", con);
                        temp2.ExecuteNonQuery();
                        SqlCommand temp = new SqlCommand("Delete from Rubric where CloId = " + id + ";", con);
                        temp.ExecuteNonQuery();
                        SqlCommand cmd = new SqlCommand("Delete from Clo where Id = " + id + ";", con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveRubrics(1);
                        RetrieveRubricsLevel(1);
                        RetrieveCLO(1);
                        RetrieveAssessmentComponent(1);
                        id = 0;
                        pictureBox5.Enabled = false;
                    }

                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView2.Columns["Delete"].Index)
            {
                try
                {
                    textBox11.Text = this.dataGridView2.CurrentRow.Cells[1].Value.ToString();
                    pictureBox5.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pictureBox5_Click_1(object sender, EventArgs e)
        {
            if (textBox11.Text != "")
            {
                if (textBox11.Text.All(x => char.IsLetterOrDigit(x) || char.IsWhiteSpace(x)) && (textBox11.Text.ToString()).Count(c => c == ' ') != (textBox11.Text.ToString()).Length)
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp = new SqlCommand("SELECT count(*) FROM Clo Where Name = @Name and Id != " + id + ";", con);
                        temp.Parameters.AddWithValue("@Name", textBox11.Text);
                        Int32 count = (Int32)temp.ExecuteScalar();
                        if (count == 0)
                        {                      
                            if (id != 0)
                            {
                                SqlCommand cmd = new SqlCommand("Update Clo Set Name=@Name,DateUpdated = @DateUpdated where Id = " + id + ";", con);
                                cmd.Parameters.AddWithValue("@Name", textBox11.Text);
                                cmd.Parameters.AddWithValue("@DateUpdated", DateTime.Now);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Successfully Updated");
                                textBox11.Clear();
                                id = 0;
                                RetrieveCLO(1);
                                pictureBox5.Enabled = false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("CLO already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                if (textBox7.Text == "")
                {
                    SqlCommand abc = new SqlCommand("Select Name,DateCreated,DateUpdated from Clo", con);
                    SqlDataAdapter da = new SqlDataAdapter(abc);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView2.DataSource = dt;
                }
                else
                {
                    if (comboBox2.Text != "Select Attribute")
                    {
                        SqlCommand cmd = new SqlCommand("Select Name,DateCreated,DateUpdated from Clo Where " + comboBox2.Text + " like '%" + textBox7.Text + "%'", con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        cmd.Parameters.AddWithValue("@Combo", comboBox2.Text);
                        cmd.Parameters.AddWithValue("@Text", textBox7.Text);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView2.DataSource = dt;
                    }
                    else
                    {
                        MessageBox.Show("Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox7.Clear();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled= true;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text != "" && comboBox4.Text != "Clo")
            {
                try
                {
                    if ((richTextBox1.Text.ToString()).Count(c => c == ' ') == (richTextBox1.Text.ToString()).Length)
                    {
                        MessageBox.Show("Invalid Details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp = new SqlCommand("SELECT Id FROM Clo Where Name = @Name", con);
                        temp.Parameters.AddWithValue("@Name", comboBox4.Text);
                        Int32 Cloid = (Int32)temp.ExecuteScalar();
                        SqlCommand temp3 = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details and CloId = " + Cloid, con);
                        temp3.Parameters.AddWithValue("@Details", richTextBox1.Text);
                        object abc = (object)temp3.ExecuteScalar();
                        if (abc == null)
                        {
                            SqlCommand temp1 = new SqlCommand("SELECT MAX(Id) FROM Rubric", con);
                            object temp2 = (object)temp1.ExecuteScalar();
                            int tempid;
                            if (temp2 == null)
                            {
                                tempid = 0;
                            }
                            else
                            {
                                try
                                {
                                    tempid = (int)temp2;
                                }
                                catch
                                {
                                    tempid = 0;
                                }
                            }
                            SqlCommand cmd = new SqlCommand("Insert into Rubric values (@Id,@Details, @CloId)", con);
                            cmd.Parameters.AddWithValue("@Id", tempid + 1);
                            cmd.Parameters.AddWithValue("@Details", richTextBox1.Text);
                            cmd.Parameters.AddWithValue("@CloId", Cloid);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Successfully saved");
                            richTextBox1.Clear();
                            comboBox4.ResetText();
                            RetrieveRubrics(1);
                            id = 0;
                            pictureBox7.Enabled = false;
                        }
                        else
                        {
                            MessageBox.Show("Rubric Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
    
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void RetrieveRubrics(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT Rubric.Id,Rubric.Details as Rubric,Clo.Name as Clo FROM Rubric JOIN Clo ON Rubric.CloId = Clo.Id", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView3.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView3.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void additemsinCombo4()
        {
            comboBox4.Items.Clear();
            var con = Configuration.getInstance().getConnection();
            SqlCommand cmd = new SqlCommand("Select Name from Clo", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox4.Items.Add(dr[0].ToString());
            }
        }

        private void comboBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled= true;    
        }

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select R.Id From Rubric R JOIN Clo C ON R.CloId = C.Id  Where R.Details = '" + this.dataGridView3.CurrentRow.Cells[2].Value.ToString() + "' and C.Name = '" + this.dataGridView3.CurrentRow.Cells[3].Value.ToString() + "'", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView3.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Corressponding AssessmentComponent,RubricLevel,StudentResult Delete with the Rubric", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp3 = new SqlCommand("DELETE FROM StudentResult Where [RubricMeasurementId] IN (SELECT Id from RubricLevel where RubricId = " + id + ")", con);
                        temp3.ExecuteNonQuery();
                        SqlCommand temp1 = new SqlCommand("Delete from RubricLevel where RubricId = " + id, con);
                        temp1.ExecuteNonQuery();
                        SqlCommand temp2 = new SqlCommand("Delete from AssessmentComponent where RubricId = " + id, con);
                        temp2.ExecuteNonQuery();
                        SqlCommand cmd = new SqlCommand("Delete from Rubric where Id = " + id + ";", con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveRubrics(1);
                        RetrieveAssessmentComponent(1);
                        RetrieveRubricsLevel(1);
                        id = 0;
                        pictureBox7.Enabled = false;
                    }
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView3.Columns["Delete"].Index)
            {
                try
                {
                    richTextBox1.Text = this.dataGridView3.CurrentRow.Cells[2].Value.ToString();
                    comboBox4.Text = this.dataGridView3.CurrentRow.Cells[3].Value.ToString();
                    pictureBox7.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text != "" && comboBox4.Text != "Clo")
            {
                try
                {
                    if ((richTextBox1.Text.ToString()).Count(c => c == ' ') == (richTextBox1.Text.ToString()).Length)
                    {
                        MessageBox.Show("Invalid Details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        if (id != 0)
                        {
                            var con = Configuration.getInstance().getConnection();
                            SqlCommand temp = new SqlCommand("SELECT Id FROM Clo Where Name = @Name", con);
                            temp.Parameters.AddWithValue("@Name", comboBox4.Text);
                            Int32 Cloid = (Int32)temp.ExecuteScalar();
                            SqlCommand temp3 = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details and CloId = " + Cloid +"and Id != "+id, con);
                            temp3.Parameters.AddWithValue("@Details", richTextBox1.Text);
                            object abc = (object)temp3.ExecuteScalar();
                            if (abc == null)
                            {
                                SqlCommand cmd = new SqlCommand("Update Rubric Set Details=@Details,CloId = @CloId where Id = " + id + ";", con);
                                cmd.Parameters.AddWithValue("@Id", id);
                                cmd.Parameters.AddWithValue("@Details", richTextBox1.Text);
                                cmd.Parameters.AddWithValue("@CloId", Cloid);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Successfully Updated");
                                richTextBox1.Clear();
                                comboBox4.ResetText();
                                RetrieveRubrics(1);
                                id = 0;
                                pictureBox7.Enabled = false;
                            }
                            else
                            {
                                MessageBox.Show("Rubric Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                if (textBox8.Text == "")
                {
                    RetrieveRubrics(1);
                }
                else
                {
                    if(comboBox3.Text != "Select Attribute") 
                    { 
                        SqlCommand cmd;
                        if (comboBox3.Text != "CloId") 
                        {
                            cmd = new SqlCommand("SELECT Rubric.Id,Rubric.Details as Rubric,Clo.Name as Clo FROM Rubric JOIN Clo ON Rubric.CloId = Clo.Id WHERE Rubric." + comboBox3.Text+" like '%"+textBox8.Text+"%'", con);
                        }
                        else 
                        {
                            cmd = new SqlCommand("SELECT Rubric.Id,Rubric.Details as Rubric,Clo.Name as Clo FROM Rubric JOIN Clo ON Rubric.CloId = Clo.Id WHERE Clo.Name like '%" + textBox8.Text+"%'", con);
                        }
                        cmd.Parameters.AddWithValue("@Combo", comboBox3.Text);
                        cmd.Parameters.AddWithValue("@Text", textBox8.Text);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);

                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView3.DataSource = dt;
                    }
                    else
                    {
                        MessageBox.Show("Please Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox8.Clear();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void rubricitemcombo()
        {
            comboBox6.Items.Clear();
            var con = Configuration.getInstance().getConnection();
            SqlCommand cmd = new SqlCommand("Select Id from Rubric", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox6.Items.Add(dr[0].ToString());
            }
        }
        private void rubricitemcombo1()
        {
            comboBox9.Items.Clear();
            var con = Configuration.getInstance().getConnection();
            SqlCommand cmd = new SqlCommand("Select Id from Rubric", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox9.Items.Add(dr[0].ToString());
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox13.Text, "[^0-9]"))
            {
                MessageBox.Show("Enter only Numbers.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox13.Text = textBox13.Text.Remove(textBox13.Text.Length - 1);
            }
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            if (richTextBox2.Text != "" && textBox13.Text !="" && comboBox6.Text != "Rubric")
            {
                try
                {
                    var con = Configuration.getInstance().getConnection();
                    //SqlCommand temp = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details", con);
                    //temp.Parameters.AddWithValue("@Details", comboBox6.Text);
                    //Int32 Rubricid = (Int32)temp.ExecuteScalar();
                    SqlCommand temp1 = new SqlCommand("SELECT Id FROM RubricLevel Where  MeasurementLevel= @MeasurementLevel", con);
                    temp1.Parameters.AddWithValue("@MeasurementLevel", textBox13.Text);
                    object level = (object)temp1.ExecuteScalar();
                    //SqlCommand temp1 = new SqlCommand("SELECT MAX(Id) FROM Rubric", con);
                    //Int32 tempid = (Int32)temp1.ExecuteScalar();
                    if (level == null)
                    {
                        if ((richTextBox2.Text.ToString()).Count(c => c == ' ') == (richTextBox2.Text.ToString()).Length)
                        {
                            MessageBox.Show("Invalid Details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand("Insert into RubricLevel values (@RubricId,@Details, @MeasurementLeval)", con);
                            cmd.Parameters.AddWithValue("@RubricId", comboBox6.Text);
                            cmd.Parameters.AddWithValue("@Details", richTextBox2.Text);
                            cmd.Parameters.AddWithValue("@MeasurementLeval", textBox13.Text);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Successfully saved");
                            richTextBox2.Clear();
                            comboBox6.ResetText();
                            textBox13.Clear();
                            RetrieveRubricsLevel(1);
                            id = 0;
                            pictureBox9.Enabled = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("This Measurement Level alreary exist for this rubric.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch
                {
                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void RetrieveRubricsLevel(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT Rubric.Details as Rubric,RubricLevel.Details  as Details,RubricLevel.MeasurementLevel FROM RubricLevel JOIN Rubric ON Rubric.Id = RubricLevel.RubricId", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView4.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView4.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView4_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Int32 Rubricid = 0;
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp1 = new SqlCommand("Select R.Id From RubricLevel R Where R.RubricId IN( SELECT Id FROM Rubric Where Rubric.Details = '"+ this.dataGridView4.CurrentRow.Cells[1].Value.ToString() + "') and R.Details = '" + this.dataGridView4.CurrentRow.Cells[2].Value.ToString() + "' and R.MeasurementLevel = '" + this.dataGridView4.CurrentRow.Cells[3].Value.ToString() + "'", con);
                id = (int)temp1.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView4.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Corressponding StudentResult also Delete with the RubricLevel", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp3 = new SqlCommand("DELETE FROM StudentResult Where [RubricMeasurementId] = " + id, con);
                        temp3.ExecuteNonQuery();
                        SqlCommand cmd = new SqlCommand("Delete from RubricLevel where Id = " + id + ";", con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveRubricsLevel(1);
                        id = 0;
                        pictureBox9.Enabled = false;
                    }
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView4.Columns["Delete"].Index)
            {
                try
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand temp = new SqlCommand("SELECT RubricId FROM RubricLevel Where Id = @Id", con);
                    temp.Parameters.AddWithValue("@Id", id);
                    Rubricid = (Int32)temp.ExecuteScalar();
                    comboBox6.Text = Rubricid.ToString();
                    richTextBox2.Text = this.dataGridView4.CurrentRow.Cells[2].Value.ToString();
                    textBox13.Text = this.dataGridView4.CurrentRow.Cells[3].Value.ToString();
                    pictureBox9.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            if (richTextBox2.Text != "" && textBox13.Text != "" && comboBox6.Text != "Rubric")
            {
                try
                {
                    if ((richTextBox2.Text.ToString()).Count(c => c == ' ') == (richTextBox2.Text.ToString()).Length)
                    {
                        MessageBox.Show("Invalid Details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp1 = new SqlCommand("SELECT Id FROM RubricLevel Where  MeasurementLevel= @MeasurementLevel and Id != " + id, con);
                        temp1.Parameters.AddWithValue("@MeasurementLevel", textBox13.Text);
                        object level = (object)temp1.ExecuteScalar();
                        if (level == null)
                        {
                            if (id != 0)
                            {
                                //MessageBox.Show(id.ToString());
                                //SqlCommand temp = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details", con);
                                //temp.Parameters.AddWithValue("@Details", comboBox6.Text);
                                //Int32 Rubricid = (Int32)temp.ExecuteScalar();
                                SqlCommand cmd = new SqlCommand("Update RubricLevel Set RubricId=@RubricId, Details=@Details,MeasurementLevel = @MeasurementLevel where Id = " + id + ";", con);
                                cmd.Parameters.AddWithValue("@Id", id);
                                cmd.Parameters.AddWithValue("@RubricId", comboBox6.Text);
                                cmd.Parameters.AddWithValue("@Details", richTextBox2.Text);
                                cmd.Parameters.AddWithValue("@MeasurementLevel", textBox13.Text);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Successfully Updated");
                                richTextBox2.Clear();
                                comboBox6.ResetText();
                                textBox13.Clear();
                                RetrieveRubricsLevel(1);
                                id = 0;
                                pictureBox9.Enabled = false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("This Measurement Level alreary exist for this rubric.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                    catch
                {
                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                if (textBox9.Text == "")
                {
                    RetrieveRubricsLevel(1);
                }
                else
                {
                    if (comboBox5.Text != "Select Attribute")
                    {
                        SqlCommand cmd;
                        if (comboBox5.Text != "Rubric")
                        {
                            cmd = new SqlCommand("SELECT Rubric.Details as Rubric,RubricLevel.Details as Details,RubricLevel.MeasurementLevel FROM RubricLevel JOIN Rubric ON Rubric.Id = RubricLevel.RubricId WHERE RubricLevel." + comboBox5.Text + " like '%" + textBox9.Text + "%'", con);
                        }
                        else
                        {
                            cmd = new SqlCommand("SELECT Rubric.Details as Rubric,RubricLevel.Details as Details,RubricLevel.MeasurementLevel FROM RubricLevel JOIN Rubric ON Rubric.Id = RubricLevel.RubricId WHERE Rubric.Details like '%" + textBox9.Text + "%'", con);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        cmd.Parameters.AddWithValue("@Combo", comboBox5.Text);
                        cmd.Parameters.AddWithValue("@Text", textBox9.Text);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView4.DataSource = dt;
                    }
                    else
                    {
                        MessageBox.Show("Please Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox9.Clear();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            id = 0;
            pictureBox12.Enabled = false;
            pictureBox1.BackgroundImage = Resources.icons8_quiz_60;
            //hidepanels();
            rubricitemcombo1();
            rubricitemcombo();
            Assessmentcombo();
            RetrieveAssessmentComponent(1);
            current.Hide();
            current = panel26;
            current.Show();
            //panel26.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            id = 0; 
            pictureBox1.BackgroundImage = Resources.report__1_;
            //hidepanels();
            //panel45.Show();
            current.Hide();
            current = panel45;
            current.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            id = 0; 
            pictureBox1.BackgroundImage = Resources.attendance__1_;
            //hidepanels();
            comboBox12.Hide();
            textBox17.Hide();
            //showattendance();
            //panel31.Show();
            current.Hide();
            current = panel31;
            current.Show();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox15.Text, "[^0-9]"))
            {
                MessageBox.Show("Enter only Numbers.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox15.Text = textBox15.Text.Remove(textBox15.Text.Length - 1);
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox12.Text, "[^0-9]"))
            {
                MessageBox.Show("Enter only Numbers.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox12.Text = textBox12.Text.Remove(textBox12.Text.Length - 1);
            }
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            if (textBox16.Text != "" && textBox15.Text != "" && textBox12.Text != "")
            {
                if (Regex.IsMatch(textBox16.Text.ToString(), @"^[A-Za-z0-9 ]+$"))
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp4 = new SqlCommand("Select Id FROM Assessment Where Title = '" + textBox16.Text + "'", con);
                        object exist = (object)temp4.ExecuteScalar();
                        if (exist != null)
                        {
                            MessageBox.Show("Assessment Already Exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if (int.Parse(textBox15.Text) == 0)
                        {
                            MessageBox.Show("Total Marks must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if (int.Parse(textBox12.Text) == 0)
                        {
                            MessageBox.Show("Total Weightage must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if((textBox16.Text.ToString()).Count(c => c == ' ') == (textBox16.Text.ToString()).Length)
                        {
                            MessageBox.Show("Invalid Title", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand("Insert into Assessment values (@Title,@DateCreated,@TotalMarks, @TotalWeightage)", con);
                            cmd.Parameters.AddWithValue("@Title", textBox16.Text);
                            cmd.Parameters.AddWithValue("@DateCreated", DateTime.Now);
                            cmd.Parameters.AddWithValue("@TotalMarks", textBox15.Text);
                            cmd.Parameters.AddWithValue("@TotalWeightage", textBox12.Text);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Successfully saved");
                            textBox12.Clear();
                            textBox15.Clear();
                            textBox16.Clear();
                            RetrieveAsseement(1);
                            id = 0;
                            pictureBox10.Enabled = false;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Title", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void RetrieveAsseement(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT Title,TotalMarks,TotalWeightage FROM Assessment", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView5.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView5.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView5_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select Id From Assessment Where Title = '" + this.dataGridView5.CurrentRow.Cells[1].Value.ToString() + "'", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView5.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Corressponding AssessmentComponent,StudentResult also Delete with the Assessment", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp3 = new SqlCommand("DELETE FROM StudentResult Where AssessmentComponentId IN (SELECT Id from AssessmentComponent where AssessmentId = " + id + ")", con);
                        temp3.ExecuteNonQuery();
                        SqlCommand temp2 = new SqlCommand("Delete from AssessmentComponent where AssessmentId = " + id + "", con);
                        temp2.ExecuteNonQuery();
                        SqlCommand cmd = new SqlCommand("Delete from Assessment where Id = " + id + ";", con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveAsseement(1);
                        RetrieveAssessmentComponent(1);
                        id = 0;
                        pictureBox10.Enabled = false;
                    }
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView5.Columns["Delete"].Index)
            {
                try
                {
                    textBox16.Text = this.dataGridView5.CurrentRow.Cells[1].Value.ToString();
                    textBox15.Text = this.dataGridView5.CurrentRow.Cells[2].Value.ToString();
                    textBox12.Text = this.dataGridView5.CurrentRow.Cells[3].Value.ToString();
                    pictureBox10.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error");
                }
            }
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            if (textBox15.Text != "" && textBox12.Text != "" && textBox16.Text != "")
            {
                if (Regex.IsMatch(textBox16.Text.ToString(), @"^[A-Za-z0-9 ]+$"))
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp4 = new SqlCommand("Select Id FROM Assessment Where Title = '" + textBox16.Text+"' and Id != "+ id, con);
                        object exist = (object)temp4.ExecuteScalar();
                        if(exist != null)
                        {
                            MessageBox.Show("Assessment Already Exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if (int.Parse(textBox15.Text) == 0)
                        {
                            MessageBox.Show("Total Marks must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if (int.Parse(textBox12.Text) == 0)
                        {
                            MessageBox.Show("Total Weightage must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if ((textBox16.Text.ToString()).Count(c => c == ' ') == (textBox16.Text.ToString()).Length)
                        {
                            MessageBox.Show("Invalid Title", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            if (id != 0)
                            {
                                SqlCommand cmd = new SqlCommand("Update Assessment Set Title=@Title, TotalMarks=@TotalMarks,TotalWeightage = @TotalWeightage where Id = " + id + ";", con);
                                cmd.Parameters.AddWithValue("@Title", textBox16.Text);
                                cmd.Parameters.AddWithValue("@TotalMarks", textBox15.Text);
                                cmd.Parameters.AddWithValue("@TotalWeightage", textBox12.Text);
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("Successfully Updated");
                                textBox16.Clear();
                                textBox15.Clear();
                                textBox12.Clear();
                                RetrieveAsseement(1);
                                id = 0;
                                pictureBox10.Enabled = false;
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Title", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if(comboBox7.Text != "Select Attribute")
                {
                    var con = Configuration.getInstance().getConnection();
                    if (textBox10.Text == "")
                    {
                        RetrieveAsseement(1);
                    }
                    else
                    {
                        SqlCommand cmd;
                        cmd = new SqlCommand("SELECT Title,TotalMarks,TotalWeightage FROM Assessment WHERE " + comboBox7.Text + " like '%" + textBox10.Text + "%'", con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView5.DataSource = dt;
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Attribute");
                    textBox10.Clear();
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Assessmentcombo()
        {
            comboBox10.Items.Clear();
            var con = Configuration.getInstance().getConnection();
            SqlCommand cmd = new SqlCommand("Select Title from Assessment", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox10.Items.Add(dr[0].ToString());
            }
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" && textBox19.Text != "" && comboBox9.Text != "Rubric Id" && comboBox10.Text != "Assessment")
            {
                if (!Regex.IsMatch(textBox20.Text.ToString(), @"^[A-Za-z0-9 ]+$"))
                {
                    MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp4 = new SqlCommand("Select Id FROM AssessmentComponent Where Name = '" + textBox20.Text + "'", con);
                        object exist = (object)temp4.ExecuteScalar();
                        if (exist != null)
                        {
                            MessageBox.Show("Assessment Already Exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else 
                        {
                            if (int.Parse(textBox19.Text) > 0)
                            {
                                SqlCommand temp1 = new SqlCommand("SELECT Id FROM Assessment Where Title = @Title", con);
                                temp1.Parameters.AddWithValue("@Title", comboBox10.Text);
                                Int32 Assessmentid = (Int32)temp1.ExecuteScalar();
                                SqlCommand temp2 = new SqlCommand("SELECT TotalMarks From Assessment Where Id = " + Assessmentid, con);
                                Int32 Assessmentmarks = (Int32)temp2.ExecuteScalar();
                                SqlCommand temp3 = new SqlCommand("Select SUM(TotalMarks) FROM AssessmentComponent WHERE Assessmentid = " + Assessmentid, con);
                                Int32 weightage = 0;
                                try
                                {
                                    weightage = (Int32)temp3.ExecuteScalar();
                                }
                                catch
                                {
                                }
                                if (Assessmentmarks - weightage >= int.Parse(textBox19.Text))
                                {
                                    //SqlCommand temp = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details", con);
                                    //temp.Parameters.AddWithValue("@Details", comboBox9.Text);
                                    //Int32 Rubricid = (Int32)temp.ExecuteScalar();
                                    if ((textBox20.Text.ToString()).Count(c => c == ' ') == (textBox20.Text.ToString()).Length)
                                    {
                                        MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    else
                                    {
                                        SqlCommand cmd = new SqlCommand("Insert into AssessmentComponent values (@Name,@RubricId, @TotalMarks,@DateCreted,@DateUpdated,@Assessmentid)", con);
                                        cmd.Parameters.AddWithValue("@Name", textBox20.Text);
                                        cmd.Parameters.AddWithValue("@RubricId", comboBox9.Text);
                                        cmd.Parameters.AddWithValue("@TotalMarks", textBox19.Text);
                                        cmd.Parameters.AddWithValue("@DateCreted", DateTime.Now);
                                        cmd.Parameters.AddWithValue("@DateUpdated", DateTime.Now);
                                        cmd.Parameters.AddWithValue("@Assessmentid", Assessmentid);
                                        cmd.ExecuteNonQuery();
                                        MessageBox.Show("Successfully saved");
                                        comboBox10.ResetText();
                                        comboBox9.ResetText();
                                        textBox19.Clear();
                                        textBox20.Clear();
                                        RetrieveAssessmentComponent(1);
                                        id = 0;
                                        pictureBox12.Enabled = false;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Total Marks must be less or equal than Corressponding Assessment", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Total Marks must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }   
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox19.Text, "[^0-9]"))
            {
                MessageBox.Show("Enter only Numbers.");
                textBox19.Text = textBox19.Text.Remove(textBox19.Text.Length - 1);
            }
        }
        private void RetrieveAssessmentComponent(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT Ac.Id,Ac.Name,R.Details  as [Rubric],Ac.TotalMarks,A.Title as Assessment FROM AssessmentComponent Ac JOIN Rubric R ON R.Id = Ac.RubricId JOIN Assessment A On A.Id = Ac.AssessmentId", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView6.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView6.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView6_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select Id From AssessmentComponent Where Name = '" + this.dataGridView6.CurrentRow.Cells[2].Value.ToString() + "' and RubricId IN ( SELECT Id FROM Rubric Where Details = '" + this.dataGridView6.CurrentRow.Cells[3].Value.ToString() + "') and TotalMarks = " + this.dataGridView6.CurrentRow.Cells[4].Value.ToString() +"", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            //MessageBox.Show(id.ToString());
            if (e.ColumnIndex == dataGridView6.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Corressponding StudentResult also Delete with the Assessment Component", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp3 = new SqlCommand("DELETE FROM StudentResult Where AssessmentComponentId = " + id + ";", con);
                        temp3.ExecuteNonQuery();
                        SqlCommand cmd = new SqlCommand("Delete from AssessmentComponent where Id = " + id + ";", con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveAssessmentComponent(1);
                        id = 0;
                        pictureBox12.Enabled = false;
                    }

                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (e.ColumnIndex != dataGridView6.Columns["Delete"].Index)
            {
                try
                {
                    textBox20.Text = this.dataGridView6.CurrentRow.Cells[2].Value.ToString();
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand temp1 = new SqlCommand("Select R.Id From AssessmentComponent A JOIN Rubric R ON R.Id = A.RubricId Where A.Name = '" + this.dataGridView6.CurrentRow.Cells[2].Value.ToString() + "' and R.Details = '" + this.dataGridView6.CurrentRow.Cells[3].Value.ToString() + "' and A.TotalMarks = " + this.dataGridView6.CurrentRow.Cells[4].Value + "", con);
                    comboBox9.Text = ((int)temp1.ExecuteScalar()).ToString();
                     
                    textBox19.Text = this.dataGridView6.CurrentRow.Cells[4].Value.ToString();
                    comboBox10.Text = this.dataGridView6.CurrentRow.Cells[5].Value.ToString();
                    comboBox9.DropDownStyle = ComboBoxStyle.DropDownList;
                    comboBox10.DropDownStyle = ComboBoxStyle.DropDownList;
                    pictureBox12.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" && textBox19.Text != "" && comboBox9.Text != "Rubric Id" && comboBox10.Text != "Assessment")
            {
                if (!Regex.IsMatch(textBox20.Text.ToString(), @"^[A-Za-z0-9 ]+$"))
                {
                    MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp4 = new SqlCommand("Select Id FROM AssessmentComponent Where Id != " + id + " and AssessmentId  IN (Select Id from Assessment Where Title = '"+comboBox10.Text+"') and Name = '"+textBox20.Text+"'", con);
                        object exist = (object)temp4.ExecuteScalar();
                        if (exist != null)
                        {
                            MessageBox.Show("Assessment Component Already Exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            if (int.Parse(textBox19.Text) > 0)
                            {
                                SqlCommand temp1 = new SqlCommand("SELECT Id FROM Assessment Where Title = @Title", con);
                                temp1.Parameters.AddWithValue("@Title", comboBox10.Text);
                                Int32 Assessmentid = (Int32)temp1.ExecuteScalar();
                                SqlCommand temp2 = new SqlCommand("SELECT TotalMarks From Assessment Where Id = " + Assessmentid, con);
                                Int32 Assessmentmarks = (Int32)temp2.ExecuteScalar();
                                SqlCommand temp3 = new SqlCommand("Select SUM(TotalMarks) FROM AssessmentComponent WHERE Assessmentid = " + Assessmentid + " and Id != " + id, con);
                                Int32 weightage = 0;
                                try
                                {
                                    weightage = (Int32)temp3.ExecuteScalar();
                                }
                                catch
                                {
                                }
                                if (Assessmentmarks - weightage >= int.Parse(textBox19.Text))
                                {
                                    if (id != 0)
                                    {
                                        //SqlCommand temp = new SqlCommand("SELECT Id FROM Rubric Where Details = @Details", con);
                                        //temp.Parameters.AddWithValue("@Details", comboBox9.Text);
                                        //Int32 Rubricid = (Int32)temp.ExecuteScalar();
                                        if ((textBox20.Text.ToString()).Count(c => c == ' ') == (textBox20.Text.ToString()).Length)
                                        {
                                            MessageBox.Show("Invalid Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        else
                                        {
                                            SqlCommand cmd = new SqlCommand("Update AssessmentComponent Set Name=@Name, RubricId=@RubricId,TotalMarks = @TotalMarks,DateUpdated = @DateUpdated,Assessmentid=@Assessmentid where Id = " + id + ";", con);
                                            cmd.Parameters.AddWithValue("@Name", textBox20.Text);
                                            cmd.Parameters.AddWithValue("@RubricId", comboBox9.Text);
                                            cmd.Parameters.AddWithValue("@TotalMarks", textBox19.Text);
                                            cmd.Parameters.AddWithValue("@DateUpdated", DateTime.Now);
                                            cmd.Parameters.AddWithValue("@Assessmentid", Assessmentid);
                                            cmd.ExecuteNonQuery();
                                            MessageBox.Show("Successfully Updated");
                                            comboBox10.ResetText();
                                            comboBox9.ResetText();
                                            textBox19.Clear();
                                            textBox20.Clear();
                                            RetrieveAssessmentComponent(1);
                                            rubricitemcombo1();
                                            id = 0;
                                            pictureBox12.Enabled = false;
                                        }
                                    }
                                    
                                }
                                else
                                {
                                    MessageBox.Show("Total Marks must be less or equal than Corressponding Assessment", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Total Marks must be greater than 0", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter the Required Input", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox8.Text != "Select Attribute")
                {
                    var con = Configuration.getInstance().getConnection();
                    if (textBox14.Text == "")
                    {
                        RetrieveAssessmentComponent(1);
                    }
                    else
                    {
                        SqlCommand cmd;
                        if (comboBox8.Text == "Rubric")
                        {
                            cmd = new SqlCommand("SELECT Ac.Id,Ac.Name,R.Details  as Rubric,Ac.TotalMarks,A.Title as Assessment FROM AssessmentComponent Ac JOIN Rubric R ON R.Id = Ac.RubricId JOIN Assessment A On A.Id = Ac.AssessmentId WHERE R.Details like '%" + textBox14.Text + "%'", con);
                        }
                        else if (comboBox8.Text == "Assessment")
                        {
                            cmd = new SqlCommand("SELECT Ac.Id,Ac.Name,R.Details  as Rubric,Ac.TotalMarks,A.Title as Assessment FROM AssessmentComponent Ac JOIN Rubric R ON R.Id = Ac.RubricId JOIN Assessment A On A.Id = Ac.AssessmentId WHERE A.Title like '%" + textBox14.Text + "%'", con);
                        }
                        else
                        {
                            cmd = new SqlCommand("SELECT Ac.Id,Ac.Name,R.Details  as Rubric,Ac.TotalMarks,A.Title as Assessment FROM AssessmentComponent Ac JOIN Rubric R ON R.Id = Ac.RubricId JOIN Assessment A On A.Id = Ac.AssessmentId WHERE Ac." + comboBox8.Text+" like '%" + textBox14.Text + "%'", con);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView6.DataSource = dt;
                        
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox14.Clear();
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {

            try
            {
                if (comboBox11.Text != "Category")
                {
                    removeattendancecolumn();
                    removeresetcolumn();
                    comboBox12.Hide();
                    textBox17.Hide();
                    if (Convert.ToDateTime(dateTimePicker1.Text) <= DateTime.Now)
                    {
                        view = true;
                        if (comboBox11.Text == "Mark")
                        {

                            var con = Configuration.getInstance().getConnection();
                            SqlCommand temp3 = new SqlCommand("Select Id From ClassAttendance Where AttendanceDate = @Date", con);
                            temp3.Parameters.AddWithValue("@Date", Convert.ToDateTime(dateTimePicker1.Text));
                            Object tempid = (Object)temp3.ExecuteScalar();
                            if (tempid == null)
                            {
                                SqlCommand cmd = new SqlCommand("SELECT RegistrationNumber,(FirstName+' '+LastName)as Name FROM Student  Where Status = 5", con);
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                DataTable dt = new DataTable();
                                da.Fill(dt);
                                dataGridView7.DataSource = dt;
                                if (frm == "")
                                {
                                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                                    checkBoxColumn.HeaderText = "Present";
                                    checkBoxColumn.Width = 30;
                                    checkBoxColumn.Name = "Present";
                                    dataGridView7.Columns.Insert(2, checkBoxColumn);
                                    for (int i = 0; i < dataGridView7.Rows.Count; i++)
                                    {
                                        dataGridView7.Rows[i].Cells[2].Value = true;
                                        dataGridView7.Rows[i].Cells[0].Style.BackColor = Color.Green;
                                    }
                                    DataGridViewCheckBoxColumn checkBoxColumn1 = new DataGridViewCheckBoxColumn();
                                    checkBoxColumn1.HeaderText = "Absent";
                                    checkBoxColumn1.Width = 30;
                                    checkBoxColumn1.Name = "Absent";
                                    dataGridView7.Columns.Insert(3, checkBoxColumn1);

                                    DataGridViewCheckBoxColumn checkBoxColumn2 = new DataGridViewCheckBoxColumn();
                                    checkBoxColumn2.HeaderText = "Leave";
                                    checkBoxColumn2.Width = 30;
                                    checkBoxColumn2.Name = "Leave";
                                    dataGridView7.Columns.Insert(4, checkBoxColumn2);

                                    DataGridViewCheckBoxColumn checkBoxColumn3 = new DataGridViewCheckBoxColumn();
                                    checkBoxColumn3.HeaderText = "Late";
                                    checkBoxColumn3.Width = 30;
                                    checkBoxColumn3.Name = "Late";
                                    dataGridView7.Columns.Insert(5, checkBoxColumn3);
                                    atten = new List<int>();
                                    for (int z = 0; z < dataGridView7.RowCount; z++)
                                    {
                                        atten.Add(1);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Attendance Already Marked", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        if (comboBox11.Text == "Update")
                        {
                            var con = Configuration.getInstance().getConnection();
                            SqlCommand cmd = new SqlCommand("SELECT RegistrationNumber,(FirstName+' '+LastName)as Name,AttendanceStatus FROM Student S JOIN StudentAttendance SA ON S.Id = SA.StudentId Join ClassAttendance C On C.Id = SA.AttendanceId Where C.AttendanceDate = @date  and S.Status = 5", con);
                            cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dataGridView7.DataSource = dt;
                            if (frm == "")
                            {
                                DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                                checkBoxColumn.HeaderText = "Present";
                                checkBoxColumn.Width = 30;
                                checkBoxColumn.Name = "Present";
                                dataGridView7.Columns.Insert(3, checkBoxColumn);


                                DataGridViewCheckBoxColumn checkBoxColumn1 = new DataGridViewCheckBoxColumn();
                                checkBoxColumn1.HeaderText = "Absent";
                                checkBoxColumn1.Width = 30;
                                checkBoxColumn1.Name = "Absent";
                                dataGridView7.Columns.Insert(4, checkBoxColumn1);

                                DataGridViewCheckBoxColumn checkBoxColumn2 = new DataGridViewCheckBoxColumn();
                                checkBoxColumn2.HeaderText = "Leave";
                                checkBoxColumn2.Width = 30;
                                checkBoxColumn2.Name = "Leave";
                                dataGridView7.Columns.Insert(5, checkBoxColumn2);

                                DataGridViewCheckBoxColumn checkBoxColumn3 = new DataGridViewCheckBoxColumn();
                                checkBoxColumn3.HeaderText = "Late";
                                checkBoxColumn3.Width = 30;
                                checkBoxColumn3.Name = "Late";
                                dataGridView7.Columns.Insert(6, checkBoxColumn3);
                                atten = new List<int>();
                                for (int z = 0; z < dataGridView7.RowCount; z++)
                                {
                                    if (dataGridView7[2, z].Value.ToString() == "1")
                                    {
                                        dataGridView7[3, z].Value = true;
                                        atten.Add(1);
                                        dataGridView7.Rows[z].Cells[0].Style.BackColor = Color.Green;
                                    }
                                    if (dataGridView7[2, z].Value.ToString() == "2")
                                    {
                                        dataGridView7[4, z].Value = true;
                                        atten.Add(2);
                                        dataGridView7.Rows[z].Cells[0].Style.BackColor = Color.Red;
                                    }
                                    if (dataGridView7[2, z].Value.ToString() == "3")
                                    {
                                        dataGridView7[5, z].Value = true;
                                        atten.Add(3);
                                        dataGridView7.Rows[z].Cells[0].Style.BackColor = Color.Yellow;
                                    }
                                    if (dataGridView7[2, z].Value.ToString() == "4")
                                    {
                                        dataGridView7[6, z].Value = true;
                                        atten.Add(4);
                                        dataGridView7.Rows[z].Cells[0].Style.BackColor = Color.Orange;
                                    }
                                }
                                dataGridView7.Columns.Remove("AttendanceStatus");
                            }
                        }
                        mark = true;
                    }
                    else
                    {
                        MessageBox.Show("Date must be less than today's date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Select Category", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
}

        private void comboBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (checkatten() == true)
            {
                if (dataGridView7.Rows.Count != 0)
                {
                    if (mark == true)
                    {
                        comboBox12.Hide();
                        textBox17.Hide();
                        if (comboBox11.Text != "Category")
                        {
                            if (comboBox11.Text == "Mark")
                            {
                                try
                                {
                                    var con = Configuration.getInstance().getConnection();

                                    SqlCommand temp1 = new SqlCommand("Insert into ClassAttendance values (@Date)", con);
                                    temp1.Parameters.AddWithValue("@Date", Convert.ToDateTime(dateTimePicker1.Text));
                                    temp1.ExecuteNonQuery();
                                    SqlCommand temp = new SqlCommand("Select Id From ClassAttendance Where AttendanceDate = @Date", con);
                                    temp.Parameters.AddWithValue("@Date", Convert.ToDateTime(dateTimePicker1.Text));
                                    Int32 Attendanceid = (Int32)temp.ExecuteScalar();
                                    int x = 0;
                                    foreach (DataGridViewRow row in dataGridView7.Rows)
                                    {
                                        if (row.Cells[0].Value == null)
                                        {
                                            break;
                                        }
                                        SqlCommand temp2 = new SqlCommand("Select Id From Student Where RegistrationNumber = @Regno", con);
                                        temp2.Parameters.AddWithValue("@Regno", row.Cells[0].Value);
                                        Int32 Studentid = (Int32)temp2.ExecuteScalar();
                                        SqlCommand cmd = new SqlCommand("Insert into StudentAttendance values (@AttendanceId,@StudentId,@Status)", con);
                                        cmd.Parameters.AddWithValue("@AttendanceId", Attendanceid);
                                        cmd.Parameters.AddWithValue("@StudentId", Studentid);
                                        cmd.Parameters.AddWithValue("@Status", atten[x]);
                                        cmd.ExecuteNonQuery();
                                        x++;
                                    }
                                    MessageBox.Show("Successfully Marked the Attendance");
                                }
                                catch
                                {
                                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (comboBox11.Text == "Update")
                            {
                                try
                                {
                                    var con = Configuration.getInstance().getConnection();
                                    int x = 0;
                                    SqlCommand temp = new SqlCommand("Select Id From ClassAttendance Where AttendanceDate = @Date", con);
                                    temp.Parameters.AddWithValue("@Date", Convert.ToDateTime(dateTimePicker1.Text));
                                    Int32 Attendanceid = (Int32)temp.ExecuteScalar();
                                    foreach (DataGridViewRow row in dataGridView7.Rows)
                                    {
                                        if (row.Cells[0].Value == null)
                                        {
                                            break;
                                        }
                                        SqlCommand temp2 = new SqlCommand("Select Id From Student Where RegistrationNumber = @Regno", con);
                                        temp2.Parameters.AddWithValue("@Regno", row.Cells[0].Value);
                                        Int32 Studentid = (Int32)temp2.ExecuteScalar();
                                        SqlCommand cmd = new SqlCommand("Update StudentAttendance Set AttendanceStatus=@Status WHERE AttendanceId = @Attendanceid and StudentId = @StudentId", con);
                                        cmd.Parameters.AddWithValue("@AttendanceId", Attendanceid);
                                        cmd.Parameters.AddWithValue("@StudentId", Studentid);
                                        cmd.Parameters.AddWithValue("@Status", atten[x]);
                                        cmd.ExecuteNonQuery();
                                        x++;
                                    }
                                    MessageBox.Show("Successfully Updated the Attendance");
                                }
                                catch
                                {
                                    MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Select Category", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    else
                    {
                        MessageBox.Show("Firstly Select Category and Insert On Grid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter all students Attendance", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool checkatten()
        {
            if (view != false)
            {
                foreach (DataGridViewRow row in dataGridView7.Rows)
                {
                    if ((row.Cells[2].Value == null || (bool)row.Cells[2].Value == false) && (row.Cells[3].Value == null || (bool)row.Cells[3].Value == false) && (row.Cells[4].Value == null || (bool)row.Cells[4].Value == false) && (row.Cells[5].Value == null || (bool)row.Cells[5].Value == false))
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }
        private void dataGridView7_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }


        private void dataGridView7_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == dataGridView7.Columns["Present"].Index)
                {
                    this.dataGridView7.CurrentRow.Cells[3].Value = false;
                    this.dataGridView7.CurrentRow.Cells[4].Value = false;
                    this.dataGridView7.CurrentRow.Cells[5].Value = false;
                    this.dataGridView7.CurrentRow.Cells[2].Value = true;
                    atten[dataGridView7.CurrentRow.Index] = 1;
                    dataGridView7.CurrentRow.Cells[0].Style.BackColor = Color.Green;
                }
                else if (e.ColumnIndex == dataGridView7.Columns["Absent"].Index)
                {
                    this.dataGridView7.CurrentRow.Cells[2].Value = false;
                    this.dataGridView7.CurrentRow.Cells[4].Value = false;
                    this.dataGridView7.CurrentRow.Cells[5].Value = false;
                    this.dataGridView7.CurrentRow.Cells[3].Value = true;
                    atten[dataGridView7.CurrentRow.Index] = 2;

                    dataGridView7.CurrentRow.Cells[0].Style.BackColor = Color.Red;
                }
                else if (e.ColumnIndex == dataGridView7.Columns["Leave"].Index)
                {
                    this.dataGridView7.CurrentRow.Cells[3].Value = false;
                    this.dataGridView7.CurrentRow.Cells[2].Value = false;
                    this.dataGridView7.CurrentRow.Cells[5].Value = false;
                    this.dataGridView7.CurrentRow.Cells[4].Value = true;
                    atten[dataGridView7.CurrentRow.Index] = 3;
                    dataGridView7.CurrentRow.Cells[0].Style.BackColor = Color.Yellow;
                }
                else
                {
                    this.dataGridView7.CurrentRow.Cells[3].Value = false;
                    this.dataGridView7.CurrentRow.Cells[4].Value = false;
                    this.dataGridView7.CurrentRow.Cells[2].Value = false;
                    this.dataGridView7.CurrentRow.Cells[5].Value = true;
                    atten[dataGridView7.CurrentRow.Index] = 4;
                    dataGridView7.CurrentRow.Cells[0].Style.BackColor = Color.Orange;
                }

            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId WHERE C.[AttendanceDate] = @date  and S.Status = 5", con);
                cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView7.DataSource = dt;
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            view = false;
            showattendance();
        }
        private void showattendance()
        {
            removeattendancecolumn();
            removeresetcolumn();
            comboBox12.Show();
            textBox17.Show();
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId WHERE C.[AttendanceDate] = @date  and St.Status = 5", con);
                cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView7.DataSource = dt;
                mark = false;

            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void removeattendancecolumn()
        {
            try
            {
                dataGridView7.Columns.Remove("Present");
                dataGridView7.Columns.Remove("Absent");
                dataGridView7.Columns.Remove("Late");
                dataGridView7.Columns.Remove("Leave");
            }
            catch
            {
            }
        }
        private void removeresetcolumn()
        {
            try
            {
                dataGridView7.Columns.Remove("RegistrationNumber");
                dataGridView7.Columns.Remove("Name");
                dataGridView7.Columns.Remove("AttendanceDate");
                dataGridView7.Columns.Remove("Status");
            }
            catch
            {
            }
        }
        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            if(textBox17.Text == "")
            {
                showattendance();
            }
            else
            {
                removeattendancecolumn();
                removeresetcolumn();
                if (comboBox12.Text == "RegistrationNumber")
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId Where St.RegistrationNumber like '%" +textBox17.Text+ "%' and C.AttendanceDate = @date and St.Status = 5;", con);
                    cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView7.DataSource = dt;
                }
                if (comboBox12.Text == "Name")
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId Where (St.FirstName+' '+St.LastName) like '%" + textBox17.Text + "%' and C.AttendanceDate = @date and St.Status = 5;", con);
                    cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView7.DataSource = dt;
                }
                if (comboBox12.Text == "AttendanceDate")
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId Where C.AttendanceDate like '%" + textBox17.Text + "%' and C.AttendanceDate = @date and St.Status = 5;", con);
                    cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView7.DataSource = dt;
                }
                if (comboBox12.Text == "Status")
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("SELECT St.RegistrationNumber,(St.FirstName+' '+St.LastName)as Name ,C.AttendanceDate ,L.Name as Status FROM StudentAttendance S JOIN ClassAttendance C ON C.Id = S.AttendanceId Join Lookup L on L.LookupId = S.AttendanceStatus Join Student St ON St.Id = S.StudentId Where L.Name like '%" + textBox17.Text + "%' and C.AttendanceDate = @date and St.Status = 5; ", con);
                    cmd.Parameters.AddWithValue("@date", Convert.ToDateTime(dateTimePicker1.Text));
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView7.DataSource = dt;
                }
                if (comboBox12.Text == "Attribute")
                {
                    MessageBox.Show("Select Attribute", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            id = 0;

            pictureBox1.BackgroundImage = Resources.icons8_student_male_60;
            //hidepanels();
            RetrieveInactive(1);
            current.Hide();
            current = panel36;
            current.Show();
            //panel36.Show();
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void RetrieveInactive(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("Select FirstName,LastName,Contact,Email,RegistrationNumber from Student Where Status = @Status", con);
                cmd.Parameters.AddWithValue("@Status", 6);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView8.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Active";
                        button.HeaderText = "Active";
                        button.Text = "Active";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView8.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                if (textBox18.Text == "")
                {
                    RetrieveInactive(1);
                }
                else
                {
                    if (comboBox13.Text != "Select Attribute")
                    {
                        SqlCommand cmd = new SqlCommand("Select FirstName,LastName,Contact,Email,RegistrationNumber from Student Where " + comboBox13.Text + " like '%" + textBox18.Text + "%' and Status = 6", con);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dataGridView8.DataSource = dt;
                    }
                    else
                    {
                        MessageBox.Show("Please Select Attribute");
                        textBox18.Clear();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView8_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand temp = new SqlCommand("Select Id From Student Where FirstName = '" + this.dataGridView8.CurrentRow.Cells[1].Value.ToString() + "' and LastName = '" + this.dataGridView8.CurrentRow.Cells[2].Value.ToString() + "' and Email = '" + this.dataGridView8.CurrentRow.Cells[4].Value.ToString() + "' and RegistrationNumber = '" + this.dataGridView8.CurrentRow.Cells[5].Value.ToString() + "'", con);
                id = (int)temp.ExecuteScalar();
            }
            catch { }
            if (e.ColumnIndex == dataGridView8.Columns["Active"].Index)
            {
                try
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Update Student Set Status = 5 where Id in (SELECT Id from Student Join Lookup On Status = LookupId Where Name = 'InActive' and Id = " + id + ");", con);
                    //MessageBox.Show("Update Student Set Status = 6 where Id in (SELECT Id from Student Join Lookup On Status = LookupId Where Name = 'Active' and Id = "+id+");");
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Successfully Active");
                    RetrieveInactive(1);
                    id = 0;
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void comboBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void comboBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void RetrieveStudentForEvaluation()
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("Select FirstName,LastName,RegistrationNumber from Student Where Status = @Status", con);
                cmd.Parameters.AddWithValue("@Status", 5);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView9.DataSource = dt;
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView9_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                try
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand temp = new SqlCommand("Select Id From Student Where FirstName = '" + this.dataGridView9.CurrentRow.Cells[0].Value.ToString() + "' and LastName = '" + this.dataGridView9.CurrentRow.Cells[1].Value.ToString() +"' and RegistrationNumber = '" + this.dataGridView9.CurrentRow.Cells[2].Value.ToString() + "'", con);
                    id = (int)temp.ExecuteScalar();
                }
                catch { }
                textBox21.Text = id.ToString();

            }
            catch
            {
                MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textBox21.Text != "Student ID")
                {
                    comboBox14.Items.Clear();
                    //comboBox14.Items.Add("Assessment Component");
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select Id from AssessmentComponent", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox14.Items.Add(dr[0].ToString());
                    }
                    //comboBox14.Text = "Assessment Component ID";
                }
            }
            catch
            {
                MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox14_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textBox21.Text != "Assessment Component ID")
                {
                    comboBox15.Items.Clear();
                    //comboBox14.Items.Add("Assessment Component");
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select RA.Details from AssessmentComponent A join Rubric R ON A.RubricId = R.Id join RubricLevel RA ON RA.RubricId = R.Id Where A.Id = @Id", con);
                cmd.Parameters.AddWithValue("@Id", int.Parse(comboBox14.Text));
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox15.Items.Add(dr[0].ToString());
                    }
                    comboBox15.Text = "Rubric Level";
                }
            }
            catch
            {
                MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
}

        private void button14_Click(object sender, EventArgs e)
        {
            if (textBox21.Text != "Student ID" && comboBox14.Text != "Assessment Component" && comboBox15.Text != "Rubric Level")
            {
                try
                {
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand temp = new SqlCommand("SELECT Id FROM RubricLevel Where Details = @Details", con);
                    temp.Parameters.AddWithValue("@Details", comboBox15.Text);
                    Int32 RubricLevelid = (Int32)temp.ExecuteScalar();
                    SqlCommand temp1 = new SqlCommand("SELECT RubricMeasurementId FROM StudentResult Where StudentId = @StudentId and AssessmentComponentId = @AssessmentComponentId", con);
                    temp1.Parameters.AddWithValue("@StudentId", textBox21.Text);
                    temp1.Parameters.AddWithValue("@AssessmentComponentId", comboBox14.Text);
                    object student = (object)temp1.ExecuteScalar();
                    if (student == null)
                    {
                        SqlCommand cmd = new SqlCommand("Insert into StudentResult values (@StudentId,@AssessmentComponentId, @RubricMeasurementId,@EvaluationDate)", con);
                        cmd.Parameters.AddWithValue("@StudentId", textBox21.Text);
                        cmd.Parameters.AddWithValue("@AssessmentComponentId", comboBox14.Text);
                        cmd.Parameters.AddWithValue("@RubricMeasurementId", RubricLevelid);
                        cmd.Parameters.AddWithValue("@EvaluationDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully saved");          
                    }
                    else if (student != null)
                    {
                        SqlCommand cmd = new SqlCommand("Update StudentResult SET RubricMeasurementId = @RubricMeasurementId Where StudentId = @StudentId and AssessmentComponentId = @AssessmentComponentId", con);
                        cmd.Parameters.AddWithValue("@StudentId", textBox21.Text);
                        cmd.Parameters.AddWithValue("@AssessmentComponentId", comboBox14.Text);
                        cmd.Parameters.AddWithValue("@RubricMeasurementId", RubricLevelid);
                        cmd.Parameters.AddWithValue("@EvaluationDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully updated");
                    }
                    textBox21.Text = "Student ID";
                    //comboBox14.Text = "Assessment Component ID";
                    comboBox15.Text = "Rubric Level";
                    RetrieveStudentResult(1);
                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter Required Inputs", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void RetrieveStudentResult(int count)
        {
            try
            {
                var con = Configuration.getInstance().getConnection();
                SqlCommand cmd = new SqlCommand("SELECT S.RegistrationNumber,(S.FirstName + ' ' + S.LastName) as Name, ass.Title,A.Name as Assessment,A.TotalMarks,R.MeasurementLevel as ObtainedMarks, R.Details FROM StudentResult SR JOIN Student S ON SR.StudentId = S.Id JOIN AssessmentComponent A ON SR.AssessmentComponentId = A.Id JOIN RubricLevel R ON R.Id = SR.RubricMeasurementId JOIN Assessment ass ON ass.Id = A.AssessmentId", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView10.DataSource = dt;
                if (count == 0)
                {
                    DataGridViewButtonColumn button = new DataGridViewButtonColumn();
                    {
                        button.Name = "Delete";
                        button.HeaderText = "Delete";
                        button.Text = "Delete";
                        button.FlatStyle = FlatStyle.Flat;
                        button.UseColumnTextForButtonValue = true;
                        button.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        this.dataGridView10.Columns.Add(button);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Unexpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView10_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == dataGridView6.Columns["Delete"].Index)
            {
                try
                {
                    var result = MessageBox.Show("Are You Sure You Want to Delete", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result.ToString() == "OK")
                    {
                        //Int32 RegistrationNumber = Int32.Parse(this.dataGridView10.CurrentRow.Cells[1].Value.ToString());
                        var con = Configuration.getInstance().getConnection();
                        SqlCommand temp1 = new SqlCommand("SELECT Id FROM Student WHERE RegistrationNumber = @RegistrationNumber", con);
                        temp1.Parameters.AddWithValue("@RegistrationNumber", this.dataGridView10.CurrentRow.Cells[1].Value.ToString());
                        Int32 Studentid = (Int32)temp1.ExecuteScalar();
                        SqlCommand temp2 = new SqlCommand("SELECT Id FROM AssessmentComponent WHERE Name = @Assessment and TotalMarks = @TotalMarks", con);
                        temp2.Parameters.AddWithValue("@Assessment", this.dataGridView10.CurrentRow.Cells[4].Value.ToString());
                        temp2.Parameters.AddWithValue("@TotalMarks", this.dataGridView10.CurrentRow.Cells[5].Value);
                        Int32 AssessmentComponentid = (Int32)temp2.ExecuteScalar();
                        SqlCommand temp3 = new SqlCommand("SELECT Id FROM RubricLevel WHERE MeasurementLevel = @ObtainedMarks and Details = @Details", con);
                        temp3.Parameters.AddWithValue("@ObtainedMarks", this.dataGridView10.CurrentRow.Cells[6].Value);
                        temp3.Parameters.AddWithValue("@Details", this.dataGridView10.CurrentRow.Cells[7].Value.ToString());
                        Int32 RubricLevelid = (Int32)temp3.ExecuteScalar();
                        SqlCommand cmd = new SqlCommand("DELETE FROM StudentResult WHERE StudentId = @StudentId and AssessmentComponentId = @AssessmentComponentId and [RubricMeasurementId] = @RubricLevelId", con);
                        cmd.Parameters.AddWithValue("@StudentId", Studentid);
                        cmd.Parameters.AddWithValue("@AssessmentComponentId", AssessmentComponentid);
                        cmd.Parameters.AddWithValue("@RubricLevelId", RubricLevelid);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Delete");
                        RetrieveStudentResult(1);
                    }

                }
                catch
                {
                    MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void comboBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled= true;
        }

        private void button15_Click(object sender, EventArgs e)
        {

            try
            {
                if (reportcount == 0 && comboBox16.Text == "Student Attendance Report")
                {
                    button15.Text = "PDF";
                    datetime1 = Convert.ToDateTime(dateTimePicker2.Text);
                    reportcount++;
                }
                else if (comboBox16.Text == "Student Attendance Report" && Convert.ToDateTime(dateTimePicker2.Text) > DateTime.Now)
                {
                    MessageBox.Show("Date must be less than today's date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    string temp = "";
                    var con = Configuration.getInstance().getConnection();
                    if (comboBox16.Text == "Assessment Wise Report")
                    {
                        temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM Assessment A JOIN AssessmentComponent AC ON AC.AssessmentId = A.Id JOIN Rubric R ON R.Id = AC.RubricId JOIN RubricLevel RL ON RL.RubricId = R.Id WHERE A.Title = '" + comboBox17.Text + "' GROUP BY RL.RubricId)\r\nSELECT (S.FirstName+' '+S.LastName) as Name,S.RegistrationNumber,A.TotalMarks,ROUND(SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks),4) as ObtainedMarks,A.TotalWeightage,ROUND((CONVERT(float,SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks))/A.TotalMarks)*A.TotalWeightage,4) as ObtainedWeightage\r\nFROM StudentResult SR\r\nJOIN Student S ON S.Id = SR.StudentId\r\nJOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId\r\nJOIN Assessment A ON AC.AssessmentId = A.Id\r\nJOIN Rubric R ON R.Id = AC.RubricId\r\nJOIN RubricLevel RL ON RL.RubricId = R.Id\r\nJOIN MAXI ON MAXI.RubricId = R.Id\r\nWHERE A.Title = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id\r\nGROUP BY S.FirstName,S.LastName,S.RegistrationNumber,A.TotalMarks,A.TotalWeightage ORDER BY S.RegistrationNumber";
                    }
                    if (comboBox16.Text == "Clo Wise Report")
                    {
                        temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM AssessmentComponent AC JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id WHERE C.Name = '" + comboBox17.Text + "' GROUP BY RL.RubricId) SELECT(S.FirstName + ' ' + S.LastName) as Name,S.RegistrationNumber, ROUND((Convert(float,SUM(ROUND(convert(float, RL.MeasurementLevel) / MAXI.maximum, 4) * AC.TotalMarks))/Sum(AC.TotalMarks))*100,4) as [Percentage(%)] FROM StudentResult SR JOIN Student S ON S.Id = SR.StudentId JOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id JOIN MAXI ON MAXI.RubricId = R.Id WHERE C.Name = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id GROUP BY S.FirstName,S.LastName,S.RegistrationNumber ORDER BY S.RegistrationNumber";
                    }
                    if (comboBox16.Text == "Analytical Attendence Report")
                    {
                        temp = "SELECT S.RegistrationNumber,(S.FirstName+' '+S.LastName) as Name,(Convert(float,P.Present)/Total.A)*100 as [Present(%)],(Convert(float,A.Absent)/Total.A)*100 as [Absent(%)],(Convert(float,L.Leave)/Total.A)*100 as [Leave(%)],(Convert(float,LA.Late)/Total.A)*100 as [Late(%)]\r\nFROM Student S\r\nLeft JOIN ( SELECT S.RegistrationNumber  as Regno,count(*) as Present\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Present'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)P\r\nON P.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Absent\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Absent'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)A\r\nON A.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Leave\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Leave'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)L\r\nON L.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Late\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Late'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)LA\r\nON LA.Regno = S.RegistrationNumber\r\nJOIN (\r\nSELECT COUNT (DISTINCT C.AttendanceDate) A\r\nFROM ClassAttendance C\r\n)Total\r\nON Total.A >=P.Present\r\nGROUP BY S.RegistrationNumber,S.FirstName,S.LastName,Total.A,P.Present,A.Absent,L.Leave,LA.Late\r\n\r\n\r\n\r\n\r\n\r\n";
                    }
                    if (comboBox16.Text == "Student Attendance Report")
                    {
                        temp = "SELECT CA.AttendanceDate,L.Name as Status\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN ClassAttendance CA\r\nON CA.Id = SA.AttendanceId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE S.RegistrationNumber = '" + comboBox17.Text + "' and CA.AttendanceDate >= '" + datetime1 + "' and CA.AttendanceDate <= '" + Convert.ToDateTime(dateTimePicker2.Text) + "' and S.Status = 5\r\nORDER BY CA.AttendanceDate";
                    }
                    if (comboBox16.Text == "Student Wise Clo Report")
                    {
                        temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM AssessmentComponent AC JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id GROUP BY RL.RubricId) SELECT(S.FirstName + ' ' + S.LastName) as Name, C.Name as CLO, ROUND((Convert(float,SUM(ROUND(convert(float, RL.MeasurementLevel) / MAXI.maximum, 4) * AC.TotalMarks))/Sum(AC.TotalMarks))*100,4) as [Percentage(%)] FROM StudentResult SR JOIN Student S ON S.Id = SR.StudentId JOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id JOIN MAXI ON MAXI.RubricId = R.Id WHERE SR.RubricMeasurementId = RL.Id and RegistrationNumber = '" + comboBox17.Text + "' GROUP BY S.FirstName,S.LastName, C.Name\r\n";
                    }
                    if (comboBox16.Text == "Student Wise Assessment Report")
                    {
                        temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM Assessment A JOIN AssessmentComponent AC ON AC.AssessmentId = A.Id JOIN Rubric R ON R.Id = AC.RubricId JOIN RubricLevel RL ON RL.RubricId = R.Id  GROUP BY RL.RubricId)\r\nSELECT A.Title,A.TotalMarks,ROUND(SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks),4) as ObtainedMarks,A.TotalWeightage,ROUND((CONVERT(float,SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks))/A.TotalMarks)*A.TotalWeightage,4) as ObtainedWeightage\r\nFROM StudentResult SR\r\nJOIN Student S ON S.Id = SR.StudentId\r\nJOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId\r\nJOIN Assessment A ON AC.AssessmentId = A.Id\r\nJOIN Rubric R ON R.Id = AC.RubricId\r\nJOIN RubricLevel RL ON RL.RubricId = R.Id\r\nJOIN MAXI ON MAXI.RubricId = R.Id\r\nWHERE S.RegistrationNumber = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id\r\nGROUP BY A.Title,A.TotalMarks,A.TotalWeightage";
                    }
                    SqlCommand cmd = new SqlCommand(temp, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (comboBox16.Text == "Student Attendance Report")
                    {
                        dataGridView11.DataSource = dt;
                    }
                    if (comboBox16.Text == "Analytical Attendence Report")
                    {
                        CreatePDF(dt, comboBox16.Text);
                    }
                    else
                    {
                        CreatePDF(dt, comboBox16.Text + "(" + comboBox17.Text + ")");
                    }
                    dateTimePicker2.Hide();
                    button15.Hide();
                    comboBox17.Hide();
                    comboBox16.Text = "Report";
                    reportcount = 0;
                }
            }
            catch
            {
                MessageBox.Show("UnExpected Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox16_TextChanged(object sender, EventArgs e)
        {
            if(comboBox16.Text != "Reports")
            {
                if(comboBox16.Text == "Clo Wise Report")
                {
                    comboBox17.Show();
                    comboBox17.Items.Clear();
                    button15.Hide();
                    dateTimePicker2.Hide();
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select Name from Clo", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comboBox17.Text = "Select Clo";
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox17.Items.Add(dr[0].ToString());
                    }

                }
                if (comboBox16.Text == "Assessment Wise Report")
                {
                    comboBox17.Show();
                    comboBox17.Items.Clear();
                    button15.Hide();
                    dateTimePicker2.Hide();
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select Title from Assessment", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comboBox17.Text="Select Assessment";
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox17.Items.Add(dr[0].ToString());
                    }
                }
                if (comboBox16.Text == "Analytical Attendence Report")
                {
                    button15.Show();
                    button15.Text = "PDF";
                    comboBox17.Hide();
                    dateTimePicker2.Hide();
                    var con = Configuration.getInstance().getConnection();
                    string temp = "SELECT S.RegistrationNumber,(S.FirstName+' '+S.LastName) as Name,(Convert(float,P.Present)/Total.A)*100 as [Present(%)],(Convert(float,A.Absent)/Total.A)*100 as [Absent(%)],(Convert(float,L.Leave)/Total.A)*100 as [Leave(%)],(Convert(float,LA.Late)/Total.A)*100 as [Late(%)]\r\nFROM Student S\r\nLeft JOIN ( SELECT S.RegistrationNumber  as Regno,count(*) as Present\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Present'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)P\r\nON P.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Absent\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Absent'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)A\r\nON A.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Leave\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Leave'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)L\r\nON L.Regno = S.RegistrationNumber\r\nLeft JOIN (\r\nSELECT S.RegistrationNumber  as Regno,count(*) as Late\r\nFROM Student S\r\nJOIN StudentAttendance SA\r\nON S.Id = SA.StudentId\r\nJOIN Lookup L\r\nON L.LookupId = SA.AttendanceStatus\r\nWHERE L.Name = 'Late'\r\nGROUP BY SA.StudentId,S.RegistrationNumber\r\n)LA\r\nON LA.Regno = S.RegistrationNumber\r\nJOIN (\r\nSELECT COUNT (DISTINCT C.AttendanceDate) A\r\nFROM ClassAttendance C\r\n)Total\r\nON Total.A >=P.Present\r\nGROUP BY S.RegistrationNumber,S.FirstName,S.LastName,Total.A,P.Present,A.Absent,L.Leave,LA.Late\r\n\r\n\r\n\r\n\r\n\r\n";
                    SqlCommand cmd = new SqlCommand(temp, con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView11.DataSource = dt;
                }
                if(comboBox16.Text == "Student Attendance Report")
                {
                    comboBox17.Show();
                    dataGridView11.DataSource = null;
                    //dataGridView11.DataBind();
                    comboBox17.Items.Clear();                   
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select RegistrationNumber from Student Where Status = 5", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comboBox17.Text = "Select Student";
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox17.Items.Add(dr[0].ToString());
                    }
                }
                if (comboBox16.Text == "Student Wise Clo Report" || comboBox16.Text == "Student Wise Assessment Report")
                {
                    comboBox17.Show();
                    comboBox17.Items.Clear();
                    button15.Hide();
                    dateTimePicker2.Hide();
                    var con = Configuration.getInstance().getConnection();
                    SqlCommand cmd = new SqlCommand("Select RegistrationNumber from Student", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    comboBox17.Text = "Select Registration Number";
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox17.Items.Add(dr[0].ToString());
                    }

                }
            }
        }

        private void comboBox17_TextChanged(object sender, EventArgs e)
        {
            string temp = "";
            var con = Configuration.getInstance().getConnection();
            if (comboBox17.Text != "Select Assessment" && comboBox17.Text != "Select Clo" && comboBox17.Text != "Select Registration Number")
            {
                button15.Show();
                button15.Text = "PDF";
            }
            else
            {
                button15.Hide();
            }
            if (comboBox16.Text == "Assessment Wise Report")
            {
                temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM Assessment A JOIN AssessmentComponent AC ON AC.AssessmentId = A.Id JOIN Rubric R ON R.Id = AC.RubricId JOIN RubricLevel RL ON RL.RubricId = R.Id WHERE A.Title = '" + comboBox17.Text + "' GROUP BY RL.RubricId)\r\nSELECT (S.FirstName+' '+S.LastName) as Name,S.RegistrationNumber,A.TotalMarks,ROUND(SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks),4) as ObtainedMarks,A.TotalWeightage,ROUND((CONVERT(float,SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks))/A.TotalMarks)*A.TotalWeightage,4) as ObtainedWeightage\r\nFROM StudentResult SR\r\nJOIN Student S ON S.Id = SR.StudentId\r\nJOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId\r\nJOIN Assessment A ON AC.AssessmentId = A.Id\r\nJOIN Rubric R ON R.Id = AC.RubricId\r\nJOIN RubricLevel RL ON RL.RubricId = R.Id\r\nJOIN MAXI ON MAXI.RubricId = R.Id\r\nWHERE A.Title = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id\r\nGROUP BY S.FirstName,S.LastName,S.RegistrationNumber,A.TotalMarks,A.TotalWeightage ORDER BY S.RegistrationNumber";
            }
            if (comboBox16.Text == "Clo Wise Report")
            {
                temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM AssessmentComponent AC JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id WHERE C.Name = '" + comboBox17.Text + "' GROUP BY RL.RubricId) SELECT(S.FirstName + ' ' + S.LastName) as Name,S.RegistrationNumber, ROUND((Convert(float,SUM(ROUND(convert(float, RL.MeasurementLevel) / MAXI.maximum, 4) * AC.TotalMarks))/Sum(AC.TotalMarks))*100,4) as [Percentage(%)] FROM StudentResult SR JOIN Student S ON S.Id = SR.StudentId JOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id JOIN MAXI ON MAXI.RubricId = R.Id WHERE C.Name = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id GROUP BY S.FirstName,S.LastName,S.RegistrationNumber ORDER BY S.RegistrationNumber";
            }
            if(comboBox16.Text == "Student Wise Clo Report")
            {
                temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM AssessmentComponent AC JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id GROUP BY RL.RubricId) SELECT(S.FirstName + ' ' + S.LastName) as Name, C.Name as CLO, ROUND((Convert(float,SUM(ROUND(convert(float, RL.MeasurementLevel) / MAXI.maximum, 4) * AC.TotalMarks))/Sum(AC.TotalMarks))*100,4) as [Percentage(%)] FROM StudentResult SR JOIN Student S ON S.Id = SR.StudentId JOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId JOIN Rubric R ON R.Id = AC.RubricId JOIN Clo C ON C.Id = R.CloId JOIN RubricLevel RL ON RL.RubricId = R.Id JOIN MAXI ON MAXI.RubricId = R.Id WHERE SR.RubricMeasurementId = RL.Id and RegistrationNumber = '" + comboBox17.Text + "' GROUP BY S.FirstName,S.LastName, C.Name\r\n";
            }
            if (comboBox16.Text == "Student Wise Assessment Report")
            {
                temp = "WITH MAXI as (SELECT RL.RubricId,Max(RL.MeasurementLevel) as maximum FROM Assessment A JOIN AssessmentComponent AC ON AC.AssessmentId = A.Id JOIN Rubric R ON R.Id = AC.RubricId JOIN RubricLevel RL ON RL.RubricId = R.Id  GROUP BY RL.RubricId)\r\nSELECT A.Title,A.TotalMarks,ROUND(SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks),4) as ObtainedMarks,A.TotalWeightage,ROUND((CONVERT(float,SUM(ROUND(convert(float,RL.MeasurementLevel)/MAXI.maximum,4)*AC.TotalMarks))/A.TotalMarks)*A.TotalWeightage,4) as ObtainedWeightage\r\nFROM StudentResult SR\r\nJOIN Student S ON S.Id = SR.StudentId\r\nJOIN AssessmentComponent AC ON AC.Id = SR.AssessmentComponentId\r\nJOIN Assessment A ON AC.AssessmentId = A.Id\r\nJOIN Rubric R ON R.Id = AC.RubricId\r\nJOIN RubricLevel RL ON RL.RubricId = R.Id\r\nJOIN MAXI ON MAXI.RubricId = R.Id\r\nWHERE S.RegistrationNumber = '" + comboBox17.Text + "' and SR.RubricMeasurementId = RL.Id\r\nGROUP BY A.Title,A.TotalMarks,A.TotalWeightage";
            }
            if (comboBox16.Text == "Student Attendance Report" && comboBox17.Text != "Select Student")
            {
                dateTimePicker2.Show();
                button15.Text = "To";
                button15.Show();
            }
            if (comboBox16.Text != "Student Attendance Report")
            {
                SqlCommand cmd = new SqlCommand(temp, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView11.DataSource = dt;
            }
            
        }
        public void CreatePDF(DataTable dataTable ,string name)
        {
            Document document = new Document();

            try
            {
                string path = Directory.GetCurrentDirectory();

                string fileName = name+".pdf";
                string filePath = Path.Combine(path, fileName);

                //Create a new PDF writer instance and bind it with the document and FileStream
                PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));

                document.Open();

                Paragraph title1 = new Paragraph();
                title1.Alignment = Element.ALIGN_CENTER;
                title1.Font = FontFactory.GetFont("Tahoma", 24);
                title1.Add("\nUniversity Of Engineering And Technology\n");
                document.Add(title1);

                Paragraph title = new Paragraph();
                title.Alignment = Element.ALIGN_CENTER;
                title.Font = FontFactory.GetFont("Tahoma", 22);
                title.Add("\n"+name+"\n\n");
                document.Add(title);

                PdfPTable pdfTable = new PdfPTable(dataTable.Columns.Count);


                foreach (DataColumn column in dataTable.Columns)
                {
                    pdfTable.AddCell(new Phrase(column.ColumnName));
                }

                foreach (DataRow row in dataTable.Rows)
                {
                    foreach (object item in row.ItemArray)
                    {
                        pdfTable.AddCell(new Phrase(item.ToString()));
                    }
                }

                document.Add(pdfTable);
                Paragraph title2 = new Paragraph();
                title2.Alignment = Element.ALIGN_CENTER;
                title2.Font = FontFactory.GetFont("Tahoma", 12);
                title2.Add("\nDated:" + DateTime.Now.ToString()+"\n\n");
                document.Add(title2);

                document.Close();

                MessageBox.Show("PDF file created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void comboBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void dataGridView7_CellValueChanged_1(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void tableLayoutPanel50_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            current.Hide();
            current = panel47;
            Dashboard();
            pictureBox1.BackgroundImage = Resources.icons8_dashboard_layout_48;
            current.Show();
        }

        private void comboBox7_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
    }
}
