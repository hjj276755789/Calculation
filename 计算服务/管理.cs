using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculation
{
    public partial class 管理 : Form
    {
        public 管理()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.OK)
            {
                string file = op.FileName;
                MBGL_DataProvider.ADD_MB(textBox1.Text.Trim(),
                    textBox2.Text.Trim(),
                    textBox3.Text.Trim(),
                    textBox4.Text.Trim(),
                    textBox5.Text.Trim(),
                    textBox6.Text.Trim(),
                    textBox6.Text.Trim()
                    );


            }
           
        }
    }
}
