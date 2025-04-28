using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace array
{
    public partial class frmDashboard : Form
    {
        Workbook workbook = new Workbook();

        public string Name { get; set; }
        public frmDashboard(string name)
        {
            InitializeComponent();
            this.Name = name;
            ShowName();
            lblTotalAct.Text = ShowTotal(9, "1").ToString();
            lblTotalInact.Text = ShowTotal(9, "0").ToString();
            lblTotalMale.Text = ShowTotal(2, "Male").ToString();
            lblTotalFemale.Text = ShowTotal(2, "Female").ToString();
            lblRed.Text = ShowTotal(4, "Red").ToString();
            lblBlue.Text = ShowTotal(4, "Blue").ToString();
            lblPink.Text = ShowTotal(4, "Pink").ToString();
            lblBasket.Text = ShowTotal(3, "Basketball").ToString();
            lblVolley.Text = ShowTotal(3, "Volleyball").ToString();
            lblBadminton.Text = ShowTotal(3, "Badminton").ToString();
            lblBSIT.Text = ShowTotal(8, "BSIT").ToString();
            lblBSCS.Text = ShowTotal(8, "BSCS").ToString();
            lblBSFM.Text = ShowTotal(8, "BSFM").ToString();
            Name = name;
        }

        public int ShowTotal(int count, string values)
        {
            workbook.LoadFromFile(@"D:\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;

            int total = 0;

            for(int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, count].Value == values)
                {
                    total++;
                }
            }
            return total;
        }

        public void ShowName()
        {
            workbook.LoadFromFile(@"D:\Downloads\EVEDRI.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;
            LogIn log = new LogIn();

            string username = log.txtUsername.Text;
            string password = log.txtPassword.Text;

            for (int i = 2; i <= row; i++)
            {
                // Check if the current row's username and password match the login credentials
                if (sheet.Range[i, 6].Value == username && sheet.Range[i, 7].Value == password)
                {
                    // If a match is found, set the button text to the student's name from the sheet
                    btnStuName.Text = sheet.Range[i, 1].Value;
                    break;  // Exit the loop once a match is found
                }
            }


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void guna2CustomGradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblTotalAct_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e) 
        {
            btnTIme.Text = DateTime.Now.ToString("hh: mm: ss tt");
            btnDate.Text = DateTime.Now.ToString("MM/ dd/ yyyy");
        }
       
        private void frmDashboard_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
