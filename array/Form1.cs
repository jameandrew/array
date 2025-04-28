using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public partial class Form1 : Form
    {
        Form2 form2 = new Form2();

        public string Name { get; set; }

        string[] Student = new string[5];
        int i = 0;
        string data = "";
        string gender = "";
        string hobbies = "";
        string favorite = "";
        string saying = "";
        string Username = "";
        string Password = "";
        string Course = "";
        public Form1(string Name)
        {
            InitializeComponent();
            this.Name = Name;
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"D:\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];
            data = "";
            gender = "";
            hobbies = "";
            favorite = "";
            saying = "";
            Username = "";
            Password = "";
            Course = "";

            if (!string.IsNullOrEmpty(txtName.Text) )
            {
               data += txtName.Text ;
            }  
            if (rdbFemale.Checked)
            {
                gender += rdbFemale.Text ;
            }
            else if(rdbMale.Checked)
            {
                gender += rdbMale.Text ;
            }
            if(chkBadminton.Checked)
            {
                hobbies += chkBadminton.Text + ", ";
            } 
            if(chkBasketball.Checked)
            {
                hobbies += chkBasketball.Text + ", ";
            }
            if(chkVolleyball.Checked)
            {
                hobbies += chkVolleyball.Text + ", ";
            }   
            if (cmbColor.SelectedIndex == 0)
            {
                favorite += cmbColor.Text;
            }
            if(cmbColor.SelectedIndex == 1)
            {
                favorite += cmbColor.Text;
            }
            if(cmbColor.SelectedIndex == 2)
            {
                favorite += cmbColor.Text;
            }
            if(!string.IsNullOrEmpty(txtSaying.Text))
            {
                saying += txtSaying.Text;
            }
            if(!string.IsNullOrEmpty(txtUname.Text))
            {
                Username = txtUname.Text;
            }
            if(!string.IsNullOrEmpty(txtPword.Text))
            {
                Password = txtPword.Text;
            }
            if (cmbCourse.SelectedIndex == 1)
            {
                Course += cmbCourse.Text;
            }
            else if (cmbCourse.SelectedIndex == 2)
            {
                Course += cmbCourse.Text;
            }
            else if (cmbCourse.SelectedIndex == 3)
            {
                Course += cmbCourse.Text;
            }


            Student[i] = data + gender + hobbies + favorite + saying;
            MessageBox.Show("You successfully stored your information", "Successfully", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            //form2.Insert(++i, txtName.Text, gender, hobbies, favorite, saying);

            int row = sheet.Rows.Length + 1;

            sheet.Range[row,1].Value = txtName.Text;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = favorite;
            sheet.Range[row, 5].Value = saying;
            sheet.Range[row, 6].Value = txtUname.Text;
            sheet.Range[row, 7].Value = txtPword.Text;
            sheet.Range[row, 8].Value = cmbCourse.Text;
            book.SaveToFile(@"D:\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);
            sheet.Range[row, 9].Value = "1";
            DataTable dt = sheet.ExportDataTable();
            form2.dgvData.DataSource = dt;

            txtName.Clear();
            rdbFemale.Checked = false;
            rdbMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
        }

        private void btnDisplayAll_Click(object sender, EventArgs e)
        {
            form2.Show();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtName.Clear();
            rdbFemale.Checked = false;
            rdbMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            Form2 form = (Form2)Application.OpenForms["Form2"];
            string hobby = "";
            string genders = "";
            string favcolor = "";

            Workbook book = new Workbook();
            book.LoadFromFile(@"D:\Downloads\EVEDRI.xlsx");
            Worksheet sheet = book.Worksheets[0];

            hobby = "";
            genders = "";
            favcolor = "";

            foreach(DataGridViewRow update in form.dgvData.SelectedRows)
            {
                update.Cells[0].Value = txtName.Text;

                if(rdbMale.Checked)
                {
                    update.Cells[1].Value = genders =  rdbMale.Text;
                }
                else if(rdbFemale.Checked)
                {
                    update.Cells[1].Value = genders = rdbFemale.Text;
                }

                if(chkVolleyball.Checked)
                {
                    update.Cells[2].Value = hobby += chkVolleyball.Text += ", ";
                }
                if (chkBasketball.Checked)
                {
                    update.Cells[2].Value = hobby += chkBasketball.Text += ", ";
                }
                if (chkBadminton.Checked)
                {
                    update.Cells[2].Value = hobby += chkBadminton.Text += ", ";
                }

                if (cmbColor.SelectedIndex == 0)
                {
                    update.Cells[3].Value = favcolor += cmbColor.Text;
                }
                if (cmbColor.SelectedIndex == 1)
                {
                    update.Cells[3].Value = favcolor += cmbColor.Text;
                }
                if (cmbColor.SelectedIndex == 2)
                {
                    update.Cells[3].Value = favcolor += cmbColor.Text;
                }

                update.Cells[4].Value = txtSaying.Text;
                update.Cells[5].Value = txtUname.Text;
                update.Cells[6].Value = txtPword.Text;

                int row = (Convert.ToInt32(lblID.Text)) + 2;

                sheet.Range[row, 1].Value = txtName.Text;
                sheet.Range[row, 2].Value = genders;
                sheet.Range[row, 3].Value = hobby;
                sheet.Range[row, 4].Value = favcolor;
                sheet.Range[row, 5].Value = txtSaying.Text;
                sheet.Range[row, 6].Value = txtUname.Text;
                sheet.Range[row, 7].Value = txtPword.Text;
                book.SaveToFile(@"D:\Downloads\EVEDRI.xlsx", ExcelVersion.Version2016);

                DataTable dt = sheet.ExportDataTable();

                form2.dgvData.DataSource = dt;

                txtName.Clear();
                rdbFemale.Checked = false;
                rdbMale.Checked = false;
                chkBadminton.Checked = false;
                chkBasketball.Checked = false;
                chkVolleyball.Checked = false;
                cmbColor.SelectedIndex = -1;
                txtSaying.Clear();
                txtUname.Clear();
                txtPword.Clear();

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmDashboard dashboard = new frmDashboard(Name);

            dashboard.Show();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
