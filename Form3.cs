using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelComparer
{
    public partial class Form3 : Form
    {
        public string SelectedFilter { get { return filterComboBox.SelectedItem.ToString(); } }


        public Form3(List<string> columnHeaders)
        {
            InitializeComponent();

            foreach(var item in columnHeaders)
            {
                filterComboBox.Items.Add(item);
            }

            filterComboBox.SelectedIndex = 0;

        }

        private void FilterButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;

        }
    }
}
