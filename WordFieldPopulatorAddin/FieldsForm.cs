using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordFieldPopulatorAddin
{
    public partial class FieldsForm : Form
    {
        internal Dictionary<string, string> Values = new Dictionary<string, string>();

        public FieldsForm(IDictionary<String, String> values)
        {
            InitializeComponent();
            tableLayoutPanel1.Controls.Clear();
            foreach (KeyValuePair<String, String> pair in values)
            {
                Label l = new Label();
                l.Text = pair.Key;
                l.Margin = new Padding(l.Margin.Left, 5, l.Margin.Right, l.Margin.Bottom);
                l.AutoSize = true;
                tableLayoutPanel1.Controls.Add(l);

                TextBox t = new TextBox();
                t.Text = pair.Value;
                t.Width = 300;
                t.Tag = pair.Key;
                tableLayoutPanel1.Controls.Add(t);
            }
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            foreach (var item in tableLayoutPanel1.Controls)
            {
                if (item is TextBox)
                {
                    var field = (TextBox)item;
                    Values.Add(field.Tag.ToString(), field.Text);
                }
            }
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
