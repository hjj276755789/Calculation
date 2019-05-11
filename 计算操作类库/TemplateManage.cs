using Aspose.Slides;
using Calculation.Base;
using Calculation.Models;
using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculation.JS
{

    public class TemplateManage
    {
        private static TemplateManage uniqueInstance;

        public static List<Rw_Zxrw_ITEM> rwlb { get; set; }

        public static Rw_Zxrw_ITEM dqrw { get; set; }
        
        public static List<string> keys { get; set; }

        public static void add_rw(int mbid, int year, int zc)
        {
            Rw_Zxrw_ITEM rw = new Rw_Zxrw_ITEM();
            rw.mbid = mbid;
            rw.nf = year;
            rw.zc = zc;
            rw.zt = Models.Enums.ZX_ZT.未开始;
            try
            {
                if (rwlb == null)
                    rwlb = new List<Rw_Zxrw_ITEM>();
                rwlb.Add(rw);
            }
            catch (Exception)
            {

                throw;
            }
           
        }
        public static void add_rw(string str,string key)
        {
            Rw_Zxrw_ITEM rw = new Rw_Zxrw_ITEM();
            var par = str.Split(',');
            if (par.Length == 3)
            {
                
                rw.mbid = Int32.Parse(par[0]);
                rw.nf = Int32.Parse(par[1]);
                rw.zc = Int32.Parse(par[2]);
                rw.zt = Models.Enums.ZX_ZT.未开始;
                rw.key = key;
            }
            else
            {

            }
           
          
            try
            {
                if (rwlb == null)
                    rwlb = new List<Rw_Zxrw_ITEM>();
                rwlb.Add(rw);
            }
            catch (Exception)
            {

                throw;
            }

        }
        public static TemplateManage ini()
        {

            if (uniqueInstance == null)
            {
                uniqueInstance = new TemplateManage();
            }
            return uniqueInstance;
        }
        public string Create_zb(int mbid, int year, int zc)
        {
            Aspose_Crack.SlideCrack();
            DataTable dt = Dal.CJGL_DataProvider.GET_CJLB_BB(mbid);
            Presentation p1 = SlideFactory.GetInstance().ppt;
            p1.Slides.RemoveAt(0);
            Base_date.init_zb(year, zc);
            Cache_param_zb.ini_zb(mbid, year, zc);
            Base_Log.Log("开始任务");
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    Base_Log.Log("第" + row["cjbh"].ints().ToString() + "号插件");
                    Type type = Type.GetType(row["cjclass"].ToString());      // 
                    var obj = System.Activator.CreateInstance(type);       // 创建实例
                    MethodInfo method = type.GetMethod(row["cjmethod"].ToString(), new Type[] { typeof(string), typeof(int) });      // 获取方法信息
                    object[] parameters = new object[] { row["cjdz"], row["cjbh"].ints() };
                    var slide = method.Invoke(obj, parameters);
                    if (slide != null && ((SlideCollection)slide).Count > 0)
                    {
                        foreach (var item in (SlideCollection)slide)
                        {
                            if (item != null)
                                p1.Slides.AddClone(item);
                        }                        // 调用方法，参数为空
                    }
                }
                string path = "E:\\zb\\" + mbid + "\\" + year + "\\" + zc + "\\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filename = path + Base.Base_date.bz + ".pptx";
                p1.Save(filename, Aspose.Slides.Export.SaveFormat.Pptx);
                return filename;
            }
            catch (Exception e)
            {
                Base_Log.Log("插件生成报错:" + e.Message);
            }
            return null;

        }
        public void Create_zb1()
        {
            if (dqrw != null)
            {
                Aspose_Crack.SlideCrack();
                DataTable dt = Dal.CJGL_DataProvider.GET_CJLB_BB(dqrw.mbid);
                Presentation p1 = SlideFactory.GetInstance().ppt;
                p1.Slides.RemoveAt(0);
                Base_date.init_zb(dqrw.nf, dqrw.zc);
                Cache_param_zb.ini_zb(dqrw.mbid, dqrw.nf, dqrw.zc);
                Base_Log.Log("开始任务");
                try
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        Base_Log.Log("第" + row["cjbh"].ints().ToString() + "号插件");
                        Type type = Type.GetType(row["cjclass"].ToString());      // 
                        var obj = System.Activator.CreateInstance(type);       // 创建实例
                        MethodInfo method = type.GetMethod(row["cjmethod"].ToString(), new Type[] { typeof(string), typeof(int) });      // 获取方法信息
                        object[] parameters = new object[] { row["cjdz"], row["cjbh"].ints() };
                        var slide = method.Invoke(obj, parameters);
                        if (slide != null && ((SlideCollection)slide).Count > 0)
                        {
                            foreach (var item in (SlideCollection)slide)
                            {
                                if (item != null)
                                    p1.Slides.AddClone(item);
                            }                        // 调用方法，参数为空
                        }
                    }
                    string path = "E:\\zb\\" + dqrw.mbid + "\\" + dqrw.nf + "\\" + dqrw.zc + "\\";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    string filename = path + Base.Base_date.bz + ".pptx";
                    p1.Save(filename, Aspose.Slides.Export.SaveFormat.Pptx);
                    if (!new Dal.RWGL_DataProvider().SET_RWZT(dqrw.mbid, dqrw.nf, dqrw.zc, RW_ZT.完成可下载, filename))
                    {
                        Base_Log.Log("创建文件成功，插入数据失败");
                    }
                        dqrw.zt = Models.Enums.ZX_ZT.生成完毕;
                }
                catch (Exception e)
                {
                    Base_Log.Log("插件生成报错:" + e.Message);
                    dqrw.zt = Models.Enums.ZX_ZT.生成完毕;
                }
            }
            else
            {
                Base_Log.Log("没有当前任务");
            }
        }
        public void execute_rw()
        {
            try
            {
                if (rwlb != null && rwlb.Count > 0)
                {
                    if (dqrw == null|| dqrw.zt == Models.Enums.ZX_ZT.生成完毕)
                    { 
                        dqrw = rwlb.FirstOrDefault();
                        lock (rwlb) {
                            rwlb.Remove(dqrw);
                        }
                    }
                }
                if (dqrw != null)
                {
                    if (dqrw.zt == Models.Enums.ZX_ZT.未开始)
                    {
                        try
                        {
                            Base_Log.Log("执行任务开始：");
                            dqrw.zt = Models.Enums.ZX_ZT.生成中;
                            Thread th = new Thread(new ThreadStart(Create_zb1));
                            th.Start();
                        }
                        catch (Exception)
                        {
                            if (keys == null)
                                keys = new List<string>();
                            keys.Add(dqrw.key);
                            dqrw.zt = Models.Enums.ZX_ZT.生成完毕;
                        }

                    }
                    else if(dqrw.zt== Models.Enums.ZX_ZT.生成完毕)
                    {
                        if (keys == null)
                            keys = new List<string>();
                        keys.Add(dqrw.key);
                        dqrw = null;
                    }

                }
            }
            catch (Exception e)
            {

                throw;
            }


        }

    }




}
