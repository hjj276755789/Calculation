using Calculation.Base;
using Calculation.JS;
using System;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;



namespace Calculation
{
    public partial class Service : Form
    {


        public Service()
        {
            this.FormClosed += Service_FormClosed;
            InitializeComponent();
            sm = new SocketStoreManager();
            tm = TemplateManage.ini();
            init();
        }


        #region 时间及队列显示
        private void timer1_Tick(object sender, EventArgs e)
        {
            sm.Initialization();
            tm.execute_rw();

            if (TemplateManage.keys != null && TemplateManage.keys.Count > 0)
            { 
                foreach (var item in TemplateManage.keys)
                {
                   var ss= SocketStoreManager.Find(item);
                    ss.sc.Send(PackData("3"));
                }
                TemplateManage.keys.Clear();
            }
            if (TemplateManage.rwlb != null && TemplateManage.rwlb.Count > 0)
            {
                if (TemplateManage.dqrw != null)
                    this.label3.Text = "任务队列：" + TemplateManage.rwlb.Count + "****当前任务：" + TemplateManage.dqrw.mbid + "***任务状态：" + TemplateManage.dqrw.zt;
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
            if (m >= 60)
            {
                f++;
                m = 0;
                if (f >= 60)
                {
                    f = 0;
                    s++;
                    if (s >= 24)
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

            if (th != null && th.IsAlive)
            {
                th.DisableComObjectEagerCleanup();
            }
        }
        #endregion

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        #endregion

        #region websocket


        private static int jsqq = 0;
        SocketStoreManager sm;
        Socket sc = null; //当前socket实体
        private static Thread th;    //端口监听线程            
        private static bool islistening = false;
        private static int port = 12345; //监听端口
        private Socket listener1 = new Socket(new IPEndPoint(IPAddress.Any, port).Address.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
        private static TemplateManage tm;

        void init()
        {

            this.timer1.Start();
            islistening = true;
            if (th == null)
            {
                th = new Thread(new ThreadStart(thread));
                th.IsBackground = true;
                th.Start();
            }
        }
        private void thread()
        {
            while (islistening)
            {
                byte[] buffer = new byte[1024];
                timer1.Start();
                try
                {
                    listener1.Bind(new IPEndPoint(IPAddress.Any, port));
                    listener1.Listen(2000);
                    Console.WriteLine("等待客户端连接....");
                    while (true)
                    {
                        Socket sc = listener1.Accept();//接受一个连接
                        SocketStore ss = new SocketStore();
                        ss.key = sc.RemoteEndPoint.ToString();
                        ss.sc = sc;
                        ss.time = DateTime.Now;
                        SocketStoreManager.Add(ss);
                        this.sc = sc;
                        Console.WriteLine("接受到了客户端：" + sc.RemoteEndPoint.ToString() + "连接....");

                        int length = sc.Receive(buffer);//接受客户端握手信息
                        sc.Send(PackHandShakeData(GetSecKeyAccetp(buffer, length)));
                        sc.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, new AsyncCallback(Recieve), sc);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }

            }
        }


        private static byte[] PackHandShakeData(string secKeyAccept)
        {
            var responseBuilder = new StringBuilder();
            responseBuilder.Append("HTTP/1.1 101 Switching Protocols" + Environment.NewLine);
            responseBuilder.Append("Upgrade: websocket" + Environment.NewLine);
            responseBuilder.Append("Connection: Upgrade" + Environment.NewLine);
            responseBuilder.Append("Sec-WebSocket-Accept: " + secKeyAccept + Environment.NewLine + Environment.NewLine);
            //如果把上一行换成下面两行，才是thewebsocketprotocol-17协议，但居然握手不成功，目前仍没弄明白！
            //responseBuilder.Append("Sec-WebSocket-Accept: " + secKeyAccept + Environment.NewLine);
            //responseBuilder.Append("Sec-WebSocket-Protocol: chat" + Environment.NewLine);

            return Encoding.UTF8.GetBytes(responseBuilder.ToString());
        }

        /// <summary>
        /// 生成Sec-WebSocket-Accept
        /// </summary>
        /// <param name="handShakeText">客户端握手信息</param>
        /// <returns>Sec-WebSocket-Accept</returns>
        private static string GetSecKeyAccetp(byte[] handShakeBytes, int bytesLength)
        {
            string handShakeText = Encoding.UTF8.GetString(handShakeBytes, 0, bytesLength);
            string key = string.Empty;
            Regex r = new Regex(@"Sec\-WebSocket\-Key:(.*?)\r\n");
            Match m = r.Match(handShakeText);
            if (m.Groups.Count != 0)
            {
                key = Regex.Replace(m.Value, @"Sec\-WebSocket\-Key:(.*?)\r\n", "$1").Trim();
            }
            byte[] encryptionString = SHA1.Create().ComputeHash(Encoding.ASCII.GetBytes(key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"));
            return Convert.ToBase64String(encryptionString);
        }

        /// <summary>
        /// 解析客户端数据包
        /// </summary>
        /// <param name="recBytes">服务器接收的数据包</param>
        /// <param name="recByteLength">有效数据长度</param>
        /// <returns></returns>
        private static string AnalyticData(byte[] recBytes, int recByteLength)
        {
            if (recByteLength < 2) { return string.Empty; }

            bool fin = (recBytes[0] & 0x80) == 0x80; // 1bit，1表示最后一帧  
            if (!fin)
            {
                return string.Empty;// 超过一帧暂不处理 
            }

            bool mask_flag = (recBytes[1] & 0x80) == 0x80; // 是否包含掩码  
            if (!mask_flag)
            {
                return string.Empty;// 不包含掩码的暂不处理
            }

            int payload_len = recBytes[1] & 0x7F; // 数据长度  

            byte[] masks = new byte[4];
            byte[] payload_data;

            if (payload_len == 126)
            {
                Array.Copy(recBytes, 4, masks, 0, 4);
                payload_len = (UInt16)(recBytes[2] << 8 | recBytes[3]);
                payload_data = new byte[payload_len];
                Array.Copy(recBytes, 8, payload_data, 0, payload_len);
            }
            else if (payload_len == 127)
            {
                Array.Copy(recBytes, 10, masks, 0, 4);
                byte[] uInt64Bytes = new byte[8];
                for (int i = 0; i < 8; i++)
                {
                    uInt64Bytes[i] = recBytes[9 - i];
                }
                UInt64 len = BitConverter.ToUInt64(uInt64Bytes, 0);

                payload_data = new byte[len];
                for (UInt64 i = 0; i < len; i++)
                {
                    payload_data[i] = recBytes[i + 14];
                }
            }
            else
            {
                Array.Copy(recBytes, 2, masks, 0, 4);
                payload_data = new byte[payload_len];
                Array.Copy(recBytes, 6, payload_data, 0, payload_len);

            }

            for (var i = 0; i < payload_len; i++)
            {
                payload_data[i] = (byte)(payload_data[i] ^ masks[i % 4]);
            }

            return Encoding.UTF8.GetString(payload_data);
        }


        /// <summary>
        /// 打包服务器数据
        /// </summary>
        /// <param name="message">数据</param>
        /// <returns>数据包</returns>
        private static byte[] PackData(string message)
        {
            byte[] contentBytes = null;

            byte[] temp = Encoding.UTF8.GetBytes(message);

            if (temp.Length < 126)
            {

                contentBytes = new byte[temp.Length + 2];

                contentBytes[0] = 0x81;

                contentBytes[1] = (byte)temp.Length;

                Array.Copy(temp, 0, contentBytes, 2, temp.Length);

            }
            else if (temp.Length < 0xFFFF)
            {

                contentBytes = new byte[temp.Length + 4];

                contentBytes[0] = 0x81;

                contentBytes[1] = 126;

                contentBytes[2] = (byte)(temp.Length >> 8);

                contentBytes[3] = (byte)(temp.Length & 0xFF);

                Array.Copy(temp, 0, contentBytes, 4, temp.Length);

            }
            else
            {

                contentBytes = new byte[temp.Length + 10];

                contentBytes[0] = 0x81;

                contentBytes[1] = 127;

                contentBytes[2] = 0;

                contentBytes[3] = 0;

                contentBytes[4] = 0;

                contentBytes[5] = 0;

                contentBytes[6] = (byte)(temp.Length >> 24);

                contentBytes[7] = (byte)(temp.Length >> 16);

                contentBytes[8] = (byte)(temp.Length >> 8);

                contentBytes[9] = (byte)(temp.Length & 0xFF);

                Array.Copy(temp, 0, contentBytes, 10, temp.Length);

            }
            return contentBytes;
        }

        /// <summary>
        /// 处理客户端发送的消息，接收成功后加入到msgPool，等待广播
        /// </summary>
        /// <param name="result">Result.</param>
        private void Recieve(IAsyncResult result)
        {

            Socket client = result.AsyncState as Socket;
            byte[] buffer = new byte[1024];
            try
            {
                client.Receive(buffer);
            }
            catch (Exception e)
            {
                Console.WriteLine("获取消息出错！" + e.Message);
            }
            string str = AnalyticData(buffer, buffer.Length);
            Console.WriteLine("************" + str + "*************8");
            lock (SocketStoreManager.s_lock)
            {
                var sc = SocketStoreManager.pool.FirstOrDefault(m => m.key.Contains(client.RemoteEndPoint.ToString()));
                if (sc == null)
                {
                    Console.WriteLine("未找到匹配socket连接");
                    return;
                }
                Console.WriteLine("找到匹配项" + sc.key);
                sc.time = DateTime.Now;
                try
                {
                    if (par(str, sc.key))
                    {
                        client.Send(PackData("2"));
                    }
                    else client.Send(PackData("1"));
                    Console.WriteLine("找到socket连接并修改时间，同时发送返回信息");
                }
                catch (Exception e)
                {
                    Console.WriteLine("发送回执出错！" + e.Message);
                }
                try
                {
                    client.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, new AsyncCallback(Recieve), client);
                }
                catch (Exception)
                {

                    Console.WriteLine("连接关闭");
                }
            }


        }

        #endregion

        public bool  par(string str,string key)
        {
            var par = str.Split(',');
            if(par.Length==3)
            {
                TemplateManage.add_rw(str,key);
                Base_Log.Log("任务已经加入队列");
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
