using Calculation.Base;
using Calculation.JS;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace Calculation
{
    public partial class 周报 : Form
    {
        public 周报()
        {
            InitializeComponent();
        }

        private void 周报_Load(object sender, EventArgs e)
        {
            int i = 0;
            foreach (DataRow item in Dal.CJGL_DataProvider.GET_CJLB().Rows )
            {
                CheckBox ck = new CheckBox();
                ck.Text = item["cjmc"].ToString();
                ck.Tag = item["cjbh"].ToString();
                ck.Checked = item["sfxz"].ToString() =="1";
                ck.Location = new System.Drawing.Point(10, 10 + i * 25);
                groupBox1.Controls.Add(ck);
                i++;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (Control item in groupBox1.Controls)
            {
                if(item is CheckBox)
                {
                    if (((CheckBox)item).Checked) { 
                        list.Add(((CheckBox)item).Tag.ToString());
                    }
                }
            }
            Dal.CJGL_DataProvider.SET_BBCJ(list);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            TemplateManage m = new TemplateManage();
            m.Create_zb(1,Int32.Parse( this.nian.Text.Trim()), Int32.Parse(this.zhou.Text.Trim()));
            MessageBox.Show("生成完毕");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Base_date.init_zb(Int32.Parse(this.nian.Text.Trim()), Int32.Parse(this.zhou.Text.Trim()));
            this.button3.Text = Base_date.bzwz;
        }
    }
}
