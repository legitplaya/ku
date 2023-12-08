using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;
using MaterialSkin;


namespace kursRab
{
    public partial class Form1 : MaterialForm
    {
        static public string User;
        public Form1()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Grey800, Primary.Grey900, Primary.BlueGrey900, Accent.Red700, TextShade.WHITE);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
 
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            
                if (materialTextBox21.Text == "admin" && materialTextBox22.Text == "admin")
                {
                    User = "admin";
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else if (materialTextBox21.Text == "user" && materialTextBox22.Text == "user")
                {
                    User = "user";
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неправильный логин или пароль", "Отказано в доступе", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }


        private void materialTextBox22_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (materialTextBox21.Text == "admin" && materialTextBox22.Text == "admin")
                {
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else if (materialTextBox21.Text == "user" && materialTextBox22.Text == "user")
                {
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неправильный логин или пароль", "Отказано в доступе", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        private void materialTextBox21_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (materialTextBox21.Text == "admin" && materialTextBox22.Text == "admin")
                {
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else if (materialTextBox21.Text == "user" && materialTextBox22.Text == "user")
                {
                    Form2 f2 = new Form2();
                    f2.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неправильный логин или пароль", "Отказано в доступе", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
    }
}
