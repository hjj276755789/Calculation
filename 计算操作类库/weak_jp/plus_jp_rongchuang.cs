using Aspose.Slides;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.JS
{
    public class plus_jp_rongchuang: plus_jp_base
    {

        public ISlideCollection _plus_jp_rongchuang_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);


                #region P2

                foreach (var item in param)
                {

                    

                    var tp = new Presentation(str);
                    var temp = tp.Slides;
                    var page = temp[1];
                    IAutoShape text1 = (IAutoShape)page.Shapes[4];
                    text1.TextFrame.Text = string.Format(text1.TextFrame.Text, item.bamc, item.ytcs[0]);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("jzgj");
                    dt.Columns.Add("lpmc");
                    dt.Columns.Add("yt");
                    dt.Columns.Add("xkts");
                    dt.Columns.Add("xkxsts");
                    dt.Columns.Add("xktnjj");

                    dt.Columns.Add("szcjts"); //上周成交数据
                    dt.Columns.Add("szcjtnjj"); //上周成交套内均价
                    dt.Columns.Add("szcjjmjj"); //上周成交建面均价

                    dt.Columns.Add("szrgts");   //上周认购套数
                    dt.Columns.Add("szrgtnjj"); //上周认购套内均价
                    dt.Columns.Add("szrgjmjj"); //上周认购建面均价


                    dt.Columns.Add("bzcjts");   //本周成交套数
                    dt.Columns.Add("bzcjtnjj"); //本周成交套内均价
                    dt.Columns.Add("szcjjmjj"); //本周成交建面均价

                    dt.Columns.Add("bzrgts"); //本周认购数据
                    dt.Columns.Add("bzrgtnjj"); //本周认购套内均价
                    dt.Columns.Add("bzrgjmjj"); //本周建面均价

                    dt.Columns.Add("tshb");  //认购环比
                    dt.Columns.Add("jghb");  //价格环比
                    dt.Columns.Add("bhyy");  //变化原因
                    dt.Columns.Add("bz");    //下周加推预计

                    //if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                    //{
                    //    dt = GET_JPXM_ROW(dt, item.jpxmlb);
                    //}

                    //Office_Tables.SetJP_FD_Table(page, dt, 2, null, null);
                    //t.AddClone(page);
                }




                #endregion
                #region P3

                foreach (var item in _plus_jp_dyt_tgtp(cjbh))
                {
                    if (item != null)
                        t.AddClone(item);
                }

                #endregion
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
        }

    }
}
