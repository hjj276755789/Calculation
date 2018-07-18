using Aspose.Slides;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculation.JS
{
 
    public class TemplateManage
    {
       
        public string Create_zb(int mbid,int year,int zc)
        {
            Aspose_Crack.SlideCrack();
            DataTable dt=  Dal.CJGL_DataProvider.GET_CJLB_BB(mbid);
            Presentation p1 = SlideFactory.GetInstance().ppt;
            Base_date.init_zb(year, zc);
            Cache_param_zb.ini_zb(mbid, year, zc);
            Base_Log.Log("开始任务");
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    Base_Log.Log("第"+row["cjbh"].ints().ToString()+"号插件");
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
                string path = "d:\\zb\\" + mbid + "\\" + year + "\\" + zc + "\\";
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

    }




}
