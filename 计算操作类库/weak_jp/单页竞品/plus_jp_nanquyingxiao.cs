using Aspose.Slides;
using Calculation.Base;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Calculation.JS
{
    public class plus_jp_nanquyingxiao : plus_jp_base
    {
        public ISlideCollection _plus_jp_nanquyingxiao_1(string str, int cjbh)
        {
            try
            {
                var param = Cache_param_zb._param_jp.Where(m => m.cjid == cjbh);
                var p = new Presentation();
                var t = p.Slides;
                t.RemoveAt(0);

                var tp1 = new Presentation(str);
                var temp1 = tp1.Slides;
                #region 竞品首页
                var page1 = temp1[0];
                IAutoShape text1 = (IAutoShape)page1.Shapes[2];
                text1.TextFrame.Text = string.Format(text1.TextFrame.Text, Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));
                #endregion
                t.AddClone(page1);
                #region 竞品分布
                var page2 = temp1[1];
                #endregion
                t.AddClone(page2);
                foreach (var item in param)
                {

                    var tp = new Presentation(str);
                    var temp = tp.Slides;

                    t.AddClone(page2);



                    
                    var page3 = temp[3];
                    DataTable dt = new DataTable();
                    #region 格局统计
                    if (item.ytcs[0] == "商务" || item.ytcs[0] == "商铺") {
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面体量);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建筑面积环比);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购套数环比);

                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价);
                        dt.Columns.Add(Base_Config_Rgsj.本周_认购建面均价环比);

                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[2];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);
                     
                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_BiGuiYuan_JPBX_Table(page3, dt, 5, null, null);
                            t.AddClone(page2);
                        }
                    }
                    #endregion
                    else
                    {
                        dt.Columns.Add(Base_Config_Jzgj.项目名称);

                        dt.Columns.Add(Base_Config_Cjba.本周_建筑面积);
                        dt.Columns.Add(Base_Config_Cjba.本周_建筑面积环比);

                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数);
                        dt.Columns.Add(Base_Config_Cjba.本周_备案套数环比);

                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价);
                        dt.Columns.Add(Base_Config_Cjba.本周_建面均价环比);

                        dt.Columns.Add(Base_Config_Rgsj.本周_变化原因);

                        IAutoShape text2 = (IAutoShape)page2.Shapes[2];
                        text2.TextFrame.Text = string.Format(text2.TextFrame.Text, item.bamc);

                        if (item.jpxmlb != null && item.jpxmlb.Count > 0)
                        {
                            dt = GET_JPXM_BX(dt, item.jpxmlb);
                            Office_Tables.SetJP_BiGuiYuan_JPBX_Table(page3, dt, 5, null, null);
                            t.AddClone(page2);
                        }
                    }


                }
                return t;
            }
            catch (Exception e)
            {
                Base_Log.Log(e.Message);
                return null;
            }
        }
        public string[] zzyt = { "高层", "小高层", "洋房", "洋楼", "别墅" };
        public System.Data.DataTable GET_JPXM_BX(System.Data.DataTable dt, List<JP_JPXM_INFO> jpxm)
        {
            foreach (var item in jpxm)
            {
                if (item.ytcs[0] == "商务"||item.ytcs[0]=="商铺")
                {

                        DataRow dr1 = dt.NewRow();

                        #region 数据准备
                        //竞品业态
                        var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && zzyt.Contains(m["yt"].ToString()) );

                        var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && zzyt.Contains(m["yt"].ToString()));
                        var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && zzyt.Contains(m["yt"].ToString()));
                        //本周本案认购数据
                        var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                        #endregion

                        dt.Rows.Add(GET_ROW(null, dr1, dt, temp_ba_bz,  temp_cjba_bz, temp_cjba_sz, item));
                }
                else
                {
                    DataRow dr1 = dt.NewRow();

                    #region 数据准备
                    //竞品业态
                    var temp_rgsj_bz = Cache_data_rgsj.bz.AsEnumerable().Where(m => m["xm"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_bz = Cache_data_cjjl.bz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    var temp_cjba_sz = Cache_data_cjjl.sz.AsEnumerable().Where(m => m["lpmc"].ToString() == item.lpcs[0] && m["yt"].ToString() == item.ytcs[0]);
                    //本周本案认购数据
                    var temp_ba_bz = temp_rgsj_bz.FirstOrDefault();
                    #endregion

                    dt.Rows.Add(GET_ROW(item.ytcs[0], dr1, dt, temp_ba_bz, temp_cjba_bz, temp_cjba_sz, item));
                }


            }


            return dt;
        }
    }
}
