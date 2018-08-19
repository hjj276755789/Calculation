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
    }
}
