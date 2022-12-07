using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelComparer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void excelFileB1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file1 = new OpenFileDialog
            {
                Title = "Choose Excel File",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                InitialDirectory = @"C:\Documents",
                RestoreDirectory = false,
                Multiselect = false
            };

            if (file1.ShowDialog() == DialogResult.OK)
            {
                excelFile1TextBox.Text = file1.FileName;
            }
        }

        private void excelFileB2_Click(object sender, EventArgs e)
        {
            OpenFileDialog file2 = new OpenFileDialog
            {
                Title = "Choose Excel File",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                InitialDirectory = @"C:\Documents",
                RestoreDirectory = false,
                Multiselect = false
            };


            if (file2.ShowDialog() == DialogResult.OK)
            {
                excelFile2TextBox.Text = file2.FileName;
            }
        }

        private bool FilePathValidator(TextBox textbox)
        {
            bool status = File.Exists(textbox.Text) ? true : false;

            if (!status)
            {
                errorProvider.SetIconPadding(textbox, -20);
                errorProvider.SetError(textbox, "Invalid file directory");
            }
            else
            {
                this.errorProvider.Clear();
            }


            return status;
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult userResponse = MessageBox.Show("Are you sure you want to exit?", "Close Application?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (userResponse == DialogResult.Yes)
                {
                    Application.Exit();
                }
                else
                {
                    e.Cancel = true;
                }

            }
        }

        private void compareExcel1Button_Click_1(object sender, EventArgs e)
        {
            bool filePath1 = FilePathValidator(excelFile1TextBox);

            bool filePath2 = FilePathValidator(excelFile2TextBox);

            if (filePath1 && filePath2)
            {
                if (Application.OpenForms.Count < 3)
                {
                    Form2 excelForm = new Form2
                    {
                        ExcelSheet1 = excelFile1TextBox.Text,
                        ExcelSheet2 = excelFile2TextBox.Text
                    };
                    excelForm.FormTitle = Path.GetFileNameWithoutExtension(excelFile1TextBox.Text);

                    //Application.UseWaitCursor = true;
                    this.WindowState = FormWindowState.Minimized;
                    excelForm.Show();
                    excelForm.BringToFront();
                    excelForm.Activate();
                }
                else
                {
                    MessageBox.Show("Maximum amount of windows open", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //this.Hide();


            }
        }

        private void compareExcel2Button_Click(object sender, EventArgs e)
        {
            bool filePath1 = FilePathValidator(excelFile1TextBox);

            bool filePath2 = FilePathValidator(excelFile2TextBox);

            if (filePath1 && filePath2)
            {
                if (Application.OpenForms.Count < 3)
                {
                    Form2 excelForm = new Form2
                    {
                        ExcelSheet1 = excelFile2TextBox.Text,
                        ExcelSheet2 = excelFile1TextBox.Text
                    };
                    excelForm.FormTitle = Path.GetFileNameWithoutExtension(excelFile2TextBox.Text);

                    //Application.UseWaitCursor = true;
                    this.WindowState = FormWindowState.Minimized;
                    excelForm.Show();
                    excelForm.BringToFront();
                    excelForm.Activate();
                }
                else
                {
                    MessageBox.Show("Maximum amount of windows open", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //this.Hide();


            }
        }
    }
}
