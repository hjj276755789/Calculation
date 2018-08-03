using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Base_Log
    {
        public static void Log(string message)
        {
            string path = @"E:\log.txt";
            
            FileStream fs = new FileStream(path, FileMode.Append);//文本加入不覆盖

            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);//转码

            sw.WriteLine(message);

            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();


        }
    }
}
