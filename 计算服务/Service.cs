using Calculation.Base;
using Calculation.JS;
using Calculation.Models.Enums;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Calculation
{
    public partial class Service : Form
    {
        public Service()
        {
            InitializeComponent();
        }
        private static bool islistening = false;
        private static HttpListener listerner;
        private static int yxsc = 0;
        private static int jsqq = 0;


        void init()
        {
            if (listerner == null)
                listerner = new HttpListener();
            islistening = true;
            this.timer1.Start();
            Thread th = new Thread(new ThreadStart(thread2));
            th.Start();
            this.button1.Enabled = false;
            this.button2.Enabled = true;
        }


        private  void thread2()
        {
            while (islistening)
            {
                try
                {
                    listerner.AuthenticationSchemes = AuthenticationSchemes.Anonymous;//指定身份验证 Anonymous匿名访问
                    listerner.Prefixes.Add("http://127.0.0.1:8000/ss/");
                    listerner.Start();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("服务启动失败...");
                    break;
                }
                Console.WriteLine("服务器启动成功.......");

                //线程池
                int minThreadNum;
                int portThreadNum;
                int maxThreadNum;
                ThreadPool.GetMaxThreads(out maxThreadNum, out portThreadNum);
                ThreadPool.GetMinThreads(out minThreadNum, out portThreadNum);
                Console.WriteLine("最大线程数：{0}", maxThreadNum);
                Console.WriteLine("最小空闲线程数：{0}", minThreadNum);
                //ThreadPool.QueueUserWorkItem(new WaitCallback(TaskProc1), x);

                Console.WriteLine("\n\n等待客户连接中。。。。");
                while (true)
                {
                    try
                    {
                        //等待请求连接
                        //没有请求则GetContext处于阻塞状态
                        HttpListenerContext ctx = listerner.GetContext();

                        ThreadPool.QueueUserWorkItem(new WaitCallback(TaskProc), ctx);
                    }
                    catch (Exception e)
                    {

                       // new Base_Log().Log("线程问题：" + e.Message);
                    }

                }
                //listerner.Stop();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            init();
        }

        private   void TaskProc(object o)
        {
            HttpListenerContext ctx = (HttpListenerContext)o;

            ctx.Response.StatusCode = 200;//设置返回给客服端http状态代码
            jsqq++;
            //接收Get参数
            try
            {
                int nf = Int32.Parse(ctx.Request.QueryString["nf"]);
                int zc = Int32.Parse(ctx.Request.QueryString["zc"]);
                int mbid = Int32.Parse(ctx.Request.QueryString["mbid"]);

                //进行处理
                dateTask dt = new dateTask(mbid, nf, null, zc);
                //接收POST参数
                Stream stream = ctx.Request.InputStream;
                System.IO.StreamReader reader = new System.IO.StreamReader(stream, Encoding.UTF8);
                String body = reader.ReadToEnd();
                Console.WriteLine("收到POST数据:" + HttpUtility.UrlDecode(body));


                //使用Writer输出http响应代码,UTF8格式
                using (StreamWriter writer = new StreamWriter(ctx.Response.OutputStream, Encoding.UTF8))
                {
                    writer.Write(SResult.Success);

                    Thread th = new Thread(new ParameterizedThreadStart(tt));
                    th.Start(dt);

                    writer.Write("任务已经启动！");
                    writer.Close();
                    ctx.Response.Close();

                }
            }
            catch (Exception e)
            {
                Base_Log.Log("准备阶段" + e.Message);
            }
           
        }

        

        public void tt(object zc)
        {
            try
            {
                dateTask dt = zc as dateTask;
                TemplateManage m = new TemplateManage();
                string xzdz= m.Create_zb(dt.mbid,dt.nf, dt.zc);
                new Dal.RWGL_DataProvider().SET_RWZT(dt.mbid, dt.nf, dt.zc, RW_ZT.完成可下载, xzdz);
            }
            catch (Exception e)
            {
                //new Base_Log().Log("生成失败："+e.Message);
            }
  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            islistening = false;
            if (listerner.IsListening) { 
                listerner.Stop();
            }
            this.timer1.Stop();
            this.Dispose();
            this.button1.Enabled = true;
            this.button2.Enabled = false;
            yxsc = 0;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            yxsc++;
            this.label1.Text = "运行时长："+ yxsc+"秒";
            this.label2.Text = "接受请求：" + jsqq +"次";
        }

        
        ~Service()
　    { 
　    // 为了保持代码的可读性性和可维护性,千万不要在这里写释放非托管资源的代码 
　    // 必须以Dispose(false)方式调用,以false告诉Dispose(bool disposing)函数是从垃圾回收器在调用 析构函数 时调用的 
   　    Dispose(false); 
　    }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(ConfigurationManager.AppSettings["sqlconn"]);
        }
    }
}
