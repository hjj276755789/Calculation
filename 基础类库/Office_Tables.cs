using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public class Office_Tables
    {
        public static void SetChart(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable chart = (ITable)sld.Shapes[index];

            for (int i =0; i < chart.Columns.Count; i++)
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int j = 0; j < (xsts.HasValue ? xsts : dt.Rows.Count); j++)
                    {
                        chart.Columns[i][j + 1].TextFrame.Text = dt.Rows[j][i].ToString();
                    }
                }
            }
        }

        public static void SetTable(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(1, false);
        }

        /// <summary>
        /// 周度排名
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        /// <param name="style"></param>
        /// <param name="xsts"></param>
        public static void SetJP_Base_ZDPM_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_FD_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz );
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3,false);
        }
        public static void SetJP_LUNENG_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_RUIAN_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_RUIAN_JQHD_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
         
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(1, false);
        }

        /// <summary>
        /// 设置周度业态排名
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        /// <param name="style"></param>
        /// <param name="xsts"></param>
        public static void SetJP_BASE_ZDYTPM_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];

            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(1, false);
        }

        public static void SetJP_YG100XMLY_ZDYTPM_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            

            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(1, false);
        }
        public static void SetJP_JUNFENG_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_HuaQiaoCheng_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
            //for (int i = 0; i < chart.Columns.Count; i++)
            //{
            //    if (dt != null && dt.Rows.Count > 0)
            //    {
            //        for (int j = 0; j < (xsts.HasValue ? xsts : dt.Rows.Count); j++)
            //        {
            //            chart.Columns[i][j + 3].TextFrame.Text = dt.Rows[j][i].ToString();
            //            //chart.Rows[i]
            //        }
            //    }
            //}
        }

        public static void SetJP_JiaZhaoYe_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][13].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_DongYuanDiChan_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_JiangBeiZuiZhiYe_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_ShouChuang_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-2);
            table.Rows[1][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[1][13].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_ShouChuang_JPBX_swsp_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_ZhongJiao_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_Langshi_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_BiGuiYuan_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_LiFanFeiCuiFu_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_WanHua_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_ZeKe_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts,string yt)
        {
            ITable table = (ITable)sld.Shapes[index];
            if (yt == "商铺") {
                table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
                table.Rows[0][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
                foreach (System.Data.DataRow item in dt.Rows)
                {
                    IRow row = table.Rows[2];
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        row[i].TextFrame.Text = item[i].ToString();
                    }
                    table.Rows.AddClone(row, false);
                }
                table.Rows.RemoveAt(2, false);
            }
            else {
                table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
                table.Rows[0][12].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
                foreach (System.Data.DataRow item in dt.Rows)
                {
                    IRow row = table.Rows[3];
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        row[i].TextFrame.Text = item[i].ToString();
                    }
                    table.Rows.AddClone(row, false);
                }
                table.Rows.RemoveAt(3, false);
            }
            
        }

        public static void SetJP_ZhongTieTouZi_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_XiangGangZhiDi_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][12].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }


        public static void SetJP_GongYuanDaDao_JPBX_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz );
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        /// <summary>
        /// 旭辉城--持续销售项目
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        /// <param name="style"></param>
        /// <param name="xsts"></param>
        public static void SetJP_XUHUICHENG_CHIXUXIAOSHOUXIANGMU_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-2);
            table.Rows[1][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        /// <summary>
        /// 旭辉-- 重点企业销售额
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        /// <param name="style"></param>
        /// <param name="xsts"></param>
        public static void SetJP_XUHUICHENG_XIAOSHOUE_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_LVDI_PUTONG_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[1][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_LVDI_SHANGWU_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-3);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-2);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[0][14].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_BEIDAZIYUAN_PT_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_JINGDIDICHAN_PT_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            table.Rows[1][12].TextFrame.Text = string.Format("{0}月备案（{1}）",Base_date.sy_First.Month, Base_date.GET_NFYFMC(Base_date.bn,Base_date.sy_First.Month));
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }


        public static void SetJP_BEIMENGZHIDI_JINGZHENGXIANGMU_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_XIANGYU_JINGPINGBEIAN_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = string.Format("备案套数（{0}）",Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));
            table.Rows[1][10].TextFrame.Text = string.Format("备案套数（{0}）", Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1));
            table.Rows[1][14].TextFrame.Text = string.Format("累计套数（{0}）", Base_date.GET_NFYFMC(Base_date.by_First, Base_date.bz_Last));
            table.Rows[1][18].TextFrame.Text = string.Format("累计套数（{0}）", "1.1-" + Base_date.bz_Last.ToString("M.d"));
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_XIANGYU_JINGPINGBEIAN_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = string.Format("备案套数（{0}）", Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));
            table.Rows[1][10].TextFrame.Text = string.Format("备案套数（{0}）", Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1));
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_XIANGYU_biaoxian_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][6].TextFrame.Text =  Base_date.GET_ZCMC(Base_date.bn, Base_date.bz -1);
            table.Rows[0][12].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz );
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_XIANGYU_ZHOUBAO_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[2][1].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[2][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz -2);
            table.Rows[2][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz -1);
            table.Rows[2][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            table.Rows[2][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-3);
            table.Rows[2][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-2);
            table.Rows[2][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[2][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_CHONGQINGSHICHANGYINGXIAOBU_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            table.Rows[1][11].TextFrame.Text = string.Format("{0}月备案", Base_date.sy_First.Month);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();
                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }


        public static void SetJP_CHONGQINGGONGSIYINGXIAOBU_XIAOSHOUE_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[1][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[1][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[1][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_CHONGQING18TI_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-1);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }

        public static void SetJP_ZHONGJIAOMANSHAN_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][1].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_ZHONGJIAOMANSHAN_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_ZHONGJIAOMANSHAN_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[1];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(1, false);
        }
        public static void SetJP_YANGGUANGCHENG_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz );
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }


        public static void SetJP_JINHUIJIANBAO_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[1][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-3);
            table.Rows[1][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz-2);
            table.Rows[1][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz -1);
            table.Rows[1][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz );
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_JINHUIJIANBAO_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][5].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
         
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }


        public static void SetJP_QIDIXIEXIN_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text =string.Format("{0}(认购)", Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1));
            table.Rows[0][6].TextFrame.Text = string.Format("{0}(认购)", Base_date.GET_ZCMC(Base_date.bn, Base_date.bz));


            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_QIDIXIEXIN_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_baoyi_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[0][9].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[0][11].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][13].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_baoyi_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 3);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 2);
            table.Rows[0][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_JINKESHICHANGZHOUBAO_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][3].TextFrame.Text = string.Format("{0}年累计成交建筑面积（万方）", Base_date.bn);
            table.Rows[0][4].TextFrame.Text = string.Format("{0}年累计成交金额（万元）", Base_date.bn);
            table.Rows[0][5].TextFrame.Text = string.Format("{0}年累计成交套内均价（元/㎡）", Base_date.bn);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            table.Rows[0][10].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_JINKESHICHANGZHOUBAO_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {

            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][3].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            table.Rows[0][7].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }

        public static void SetJP_RONGCHUANG_1_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][4].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][8].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[3];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(3, false);
        }
        public static void SetJP_RONGCHUANG_2_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
        public static void SetJP_RONGCHUANG_3_Table(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style, int? xsts)
        {
            ITable table = (ITable)sld.Shapes[index];
            table.Rows[0][2].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz - 1);
            table.Rows[0][6].TextFrame.Text = Base_date.GET_ZCMC(Base_date.bn, Base_date.bz);
            foreach (System.Data.DataRow item in dt.Rows)
            {
                IRow row = table.Rows[2];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i].TextFrame.Text = item[i].ToString();

                }
                table.Rows.AddClone(row, false);
            }
            table.Rows.RemoveAt(2, false);
        }
    }

}
