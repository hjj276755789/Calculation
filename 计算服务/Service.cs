using Calculation.Base;
using Calculation.JS;
using Calculation.Models.Enums;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
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
            this.FormClosed += Service_FormClosed;
            InitializeComponent();
            init();
        }
        private static bool islistening = false;
        private static HttpListener listerner;
        private static int jsqq = 0;
        private static Thread th;
        private static TemplateManage tm;


        void init()
        {
            tm = TemplateManage.ini();
            if (listerner == null)
                listerner = new HttpListener();
            islistening = true;
            this.timer1.Start();
            if (th == null) { 
                th = new Thread(new ThreadStart(thread2));
                th.IsBackground = true;
            }
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
                    Console.WriteLine("服务启动失败..."+ex.Message);
                    break;
                }
                Console.WriteLine("服务器启动成功.......");

                //线程池
                int minThreadNum;
                int portThreadNum;
                int maxThreadNum;
                ThreadPool.GetMaxThreads(out maxThreadNum, out portThreadNum);
                ThreadPool.GetMinThreads(out minThreadNum, out portThreadNum);
                Base_Log.Log(string.Format("最大线程数：{0}", maxThreadNum));
                Base_Log.Log(string.Format("最小空闲线程数：{0}", minThreadNum));
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

                        Base_Log.Log("线程问题：" + e.Message);
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

                    Thread th = new Thread(new ParameterizedThreadStart(tt));
                    th.Start(dt);
                    writer.Write("{ isSucessfull: true,Msg=任务已经启动}");
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
                TemplateManage.add_rw(dt.mbid, dt.nf, dt.zc);
                Base_Log.Log("任务已经加入队列");
                //string xzdz= tm.Create_zb(dt.mbid,dt.nf, dt.zc);
                //if (!string.IsNullOrEmpty(xzdz))
                //{
                //    if (!new Dal.RWGL_DataProvider().SET_RWZT(dt.mbid, dt.nf, dt.zc, RW_ZT.完成可下载, xzdz))
                //        Base_Log.Log("创建文件成功，插入数据失败");
                //    }
                //else
                //{
                //    Base_Log.Log("生成失败：下载地址并未生成并返回！");
                //    return;
                //}
                //Base_Log.Log("生成成功\n");
            }
            catch (Exception e)
            {
                Base_Log.Log("生成失败：" + e.Message);
                //关闭线程
                Process.GetCurrentProcess().Kill();
                
            }
  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ini();
        }

        public void ini()
        {
            n = 0;
            r = 0;
            s = 0;
            f = 0;
            m = 0;
            islistening = false;
            if (listerner == null)
                return;
            if (listerner.IsListening)
            {
                listerner.Stop();
            }
            this.button1.Enabled = true;
            this.button2.Enabled = false;
            timer1.Stop();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.label1.Text = "运行时长：" + sj();
            this.label2.Text = "接受请求：" + jsqq + "次";
            tm.execute_rw();
            if(TemplateManage.rwlb!=null&&TemplateManage.rwlb.Count>0)
            { 
                if(TemplateManage.dqrw!=null)
                    this.label3.Text = "任务队列：" + TemplateManage.rwlb.Count + "****当前任务："+TemplateManage.dqrw.mbid+"***任务状态："+TemplateManage.dqrw.zt;
            }
            else
            {
                this.label3.Text = "任务队列：0";
            }
            //tm.tt();
        }

        #region 时间显示

     
        /// <summary>
        /// 年
        /// </summary>
        private static int n = 0;
        /// <summary>
        /// 日
        /// </summary>
        private static int r = 0;
        /// <summary>
        /// 时
        /// </summary>
        private static int s = 0;
        /// <summary>
        /// 分
        /// </summary>
        private static int f = 0;
        /// <summary>
        /// 秒
        /// </summary>
        private static int m = 0;

        public static string sj()
        {
            m++;
            if(m>=60)
            {
                f++;
                m = 0;
                if (f >= 60)
                { 
                    f = 0;
                    s++;
                    if(s>=24)
                    {
                        s = 0;
                        r++;
                        if (r >= 360)
                        { 
                            r = 0;
                            n++;
                        }

                    }
                }

            }
            return string.Format("{0}年{1}天{2}时{3}分{4}秒", n, r, s, f, m);
        }






        private void Service_FormClosed(object sender, FormClosedEventArgs e)
        {
           
            if(th!=null&&th.IsAlive)
            {
                th.DisableComObjectEagerCleanup();
            }
        }
        #endregion

    }
}
