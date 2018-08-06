using Aspose.Cells;

using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;

namespace Calculation
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //ModifyInMemory.ActivateMemoryPatching();
            button1.Enabled = false;
            button2.Enabled = true;

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            if (opf.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(opf.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt=  cs.ExportDataTable(1,0,cs.MaxDataRow,cs.MaxDataColumn+1);
                timer1.Start();
                Thread th = new Thread(new ParameterizedThreadStart(thread1));
                
                th.Start(dt);
            }
           
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            if (opf.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(opf.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
                timer2.Start();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            if (opf.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(opf.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
                timer2.Start();
                Thread th = new Thread(new ParameterizedThreadStart(thread3));

                th.Start(dt);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {


        }

        private static void thread1(Object dt)
        {
           // var t=  Dal.ZB_Data_CJBA_DataProvider.Insert((DataTable)dt);     
        }


    
     
        private static void thread3(Object dt)
        {
            var t = Dal.ZB_Data_TDCJ_DataProvider.Insert((DataTable)dt);
        }
        private static void thread4()
        {

        }
        private int a = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {

        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            button4.Text = Dal.ZB_Data_CJBA_DataProvider.index.ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
              Thread InvokeThread = new Thread(new ThreadStart(InvokeMethod));  
               InvokeThread.SetApartmentState(ApartmentState.STA);  
               InvokeThread.Start();  
               InvokeThread.Join();  

        }

        private void InvokeMethod()
         {  
             OpenFileDialog InvokeDialog = new OpenFileDialog();  

             if (InvokeDialog.ShowDialog() == DialogResult.OK)  
             {
                    Workbook workbook = new Workbook(InvokeDialog.FileName);
                    Cells cs = workbook.Worksheets[0].Cells;
                    DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);

                    Dal.Jsjg_yb_DataProvider.ADD_SCGXFX(dt);

                }  

         }
        private void InvokeMethodpsb()
        {
            OpenFileDialog InvokeDialog = new OpenFileDialog();

            if (InvokeDialog.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(InvokeDialog.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);

                Dal.Jsjg_yb_DataProvider.ADD_SCGXFX_PSB(dt);

            }

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            Thread InvokeThread = new Thread(new ThreadStart(InvokeMethodpsb));
            InvokeThread.SetApartmentState(ApartmentState.STA);
            InvokeThread.Start();
            InvokeThread.Join();
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Thread InvokeThread = new Thread(new ThreadStart(scjgfx));
            InvokeThread.SetApartmentState(ApartmentState.STA);
            InvokeThread.Start();
            InvokeThread.Join();
        }

        private void scjgfx()
        {
            OpenFileDialog InvokeDialog = new OpenFileDialog();

            if (InvokeDialog.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(InvokeDialog.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);

                Dal.Jsjg_yb_DataProvider.ADD_SCJGFX(dt);

            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog InvokeDialog = new OpenFileDialog();

            if (InvokeDialog.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook(InvokeDialog.FileName);
                Cells cs = workbook.Worksheets[0].Cells;
                DataTable dt = cs.ExportDataTable(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);

                Dal.Jsjg_zb_DataProvider.ADD_XKPZS(dt);

            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
           
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Service s = new Service();
            s.Show();
        }
    }
}
