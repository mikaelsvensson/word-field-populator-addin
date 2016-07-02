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
                t.TextChanged += FieldValue_TextChanged;
                tableLayoutPanel1.Controls.Add(t);
            }
        }

        private void FieldValue_TextChanged(object sender, EventArgs e)
        {
            var field = (TextBox)sender;
            Values.Add(field.Tag.ToString(), field.Text);
        }

        private void FieldsForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
