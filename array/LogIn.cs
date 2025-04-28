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
    public partial class LogIn : Form
    {
        public LogIn()
        {
            InitializeComponent();
        }

        public void btnLOgin_Click(object sender, EventArgs e)
        {   
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"D:\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;

            bool log = true;
            string username = "";
           
            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i,6].Value == txtUsername.Text && sheet.Range[i,7].Value == txtPassword.Text)
                {
                    username = sheet.Range[i, 1].Value;
                    log = true; break;
                }
                else
                {
                    log = false;
                }
            }

            if (log == true)
            {
                MessageBox.Show("Log-In successfully", "Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                Form1 frm1 = new Form1(username);
                frm1.ShowDialog();
            }
            else
            {
                MessageBox.Show("Log-In unsuccessful", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (txtPassword.UseSystemPasswordChar)
            {
                pictureBox1.Image = Image.FromFile("C:\\Users\\HF\\Downloads\\hide.png");
                txtPassword.UseSystemPasswordChar = false;
                pictureBox1.Text = "Show";
            }
            else
            {
                pictureBox1.Image = Image.FromFile("C:\\Users\\HF\\Downloads\\visibility.png");
                txtPassword.UseSystemPasswordChar = true;
                pictureBox1.Text = "Hide";
            }

            pictureBox1.Refresh();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
