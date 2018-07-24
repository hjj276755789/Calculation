
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Calculation.Base;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
namespace Calculation.Base
{
    public class Office_Charts
    {
        /// <summary>
        /// 生成单列图表
        /// </summary>
        /// <param name="sld">当前ppt页面</param>
        /// <param name="dt">数据</param>
        /// <param name="index">图表所属表格排序（当前slide）</param>
        public static void SingleAxexchart(ISlide sld, System.Data.DataTable dt, int index, Office_ChartStyle style)
        {
            IChart chart = (IChart)sld.Shapes[index];
            
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count+1);
            //实例化图表数据表
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;
            chart.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);

            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);
            chart.ChartData.WriteWorkbookStream(mem);

            //设置数据区域
            chart.ChartData.SetRange(range);
                //交换横纵坐标
                if (style.坐标方向 == Base_Config.坐标方向.横向)
                {
                    chart.ChartData.SwitchRowColumn();
                }

                IChartSeries series = chart.ChartData.Series[0];

                series.Labels.DefaultDataLabelFormat.ShowValue = style.是否显示文字;
                series.Labels.DefaultDataLabelFormat.Position  = style.文字位置;
            series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = style.文字旋转方向;
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sld">当前ppt页面</param>
        /// <param name="dt">数据</param>
        /// <param name="fc">第一坐标列</param>
        /// <param name="sc">第二坐标列</param>
        /// <param name="index">图表所属表格排序（当前slide）</param>
        public static void DoubleAxexchart(ISlide sld, System.Data.DataTable dt, int index,int fc,int sc)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);
            
            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;
            
            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);
           
            t1.ChartData.SetRange(range);

            IChartSeries series = t1.ChartData.Series[fc];

            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;

            ////设置第二个系列表
            IChartSeries series1 = t1.ChartData.Series[sc];
            series1.PlotOnSecondAxis = true;
            series1.Type = ChartType.StackedLineWithMarkers;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
            series1.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sld">当前ppt页面</param>
        /// <param name="dt">数据</param>
        /// <param name="fc">第一坐标列</param>
        /// <param name="sc">第二坐标列</param>
        /// <param name="index">图表所属表格排序（当前slide）</param>
        public static void ThreeWchart(ISlide sld, System.Data.DataTable dt, int index)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);

            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);

            IChartSeries series = t1.ChartData.Series[0];
            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
            //series.Labels.DefaultDataLabelFormat.Format.Fill.FillType = FillType.Solid;
            //series.Labels.DefaultDataLabelFormat.Format.Fill.SolidFillColor.Color = System.Drawing.Color.White;
            //series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;
            IChartSeries series1 = t1.ChartData.Series[1];
            series1.Type = ChartType.ClusteredColumn;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
            //series1.Labels.DefaultDataLabelFormat.TextFormat.
            //series1.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;
            ////设置第二个系列表
            IChartSeries series2 = t1.ChartData.Series[2];
            series2.PlotOnSecondAxis = true;
            series2.Type = ChartType.LineWithMarkers;
            series2.Labels.DefaultDataLabelFormat.ShowValue = true;
            series2.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
            series2.Labels.DefaultDataLabelFormat.Format.Fill.FillType = FillType.Solid;
            series2.Labels.DefaultDataLabelFormat.Format.Fill.SolidFillColor.Color = System.Drawing.Color.White;
            //series2.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;

            IChartSeries series3 = t1.ChartData.Series[3];
            series3.PlotOnSecondAxis = true;
            series3.Type = ChartType.Line;
            //series3.Labels.DefaultDataLabelFormat.ShowValue = false;
            series3.Labels[9].DataLabelFormat.ShowValue= true;
            series3.Labels.DefaultDataLabelFormat.Format.Fill.FillType = FillType.Solid;
            series3.Labels.DefaultDataLabelFormat.Format.Fill.SolidFillColor.Color = System.Drawing.Color.DarkRed;
            series3.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
            series3.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;


        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sld">当前ppt页面</param>
        /// <param name="dt">数据</param>
        /// <param name="index">图表所属表格排序（当前slide）</param>
        public static void DoubleAxexchart(ISlide sld ,System.Data.DataTable dt, int index,ChartType type)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count + 1) + "$" + (dt.Rows.Count + 1);
            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);
            IChartSeries series = t1.ChartData.Series[0];
            series.Type = t1.Chart.Type;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;

            ////设置第二个系列表
            IChartSeries series1 = t1.ChartData.Series[2];
            series1.PlotOnSecondAxis = true;
            series1.Type = ChartType.StackedLineWithMarkers;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
            series1.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.TextVerticalType = TextVerticalType.Vertical270;




        }


        /// <summary>
        /// 供需分析图表
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        /// <param name="fc"></param>
        /// <param name="sc"></param>
        public static void Chart_gxfx(ISlide sld, System.Data.DataTable dt,int index)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);

            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);

            ///第一列
            IChartSeries series = t1.ChartData.Series[0];
            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
            ///第二列
            IChartSeries series1 = t1.ChartData.Series[1];
            series1.Type = ChartType.ClusteredColumn;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
            ////设置第二个系列表
            IChartSeries series2 = t1.ChartData.Series[2];
            series2.PlotOnSecondAxis = true;
            series2.Type = ChartType.StackedLineWithMarkers;
            series2.Labels.DefaultDataLabelFormat.ShowValue = true;
            series2.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
        }


        public static void Chart_gxzs(ISlide sld, System.Data.DataTable dt, int index)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);

            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);
            t1.ChartData.SwitchRowColumn();

            ///第一列
            IChartSeries series = t1.ChartData.Series[0];
            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            ///第二列
            IChartSeries series1 = t1.ChartData.Series[1];
            series1.Type = ChartType.ClusteredColumn;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            ////设置第二个系列表
            IChartSeries series2 = t1.ChartData.Series[2];
            series2.PlotOnSecondAxis = true;
            series2.Type = ChartType.StackedLineWithMarkers;
            series2.Labels.DefaultDataLabelFormat.ShowValue = true;
            series2.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
        }


        public static void Chart_cjqs(ISlide sld, System.Data.DataTable dt, int index)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);

            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);
            t1.ChartData.SwitchRowColumn();

            ///第一列
            IChartSeries series = t1.ChartData.Series[0];
            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            ///第二列
            IChartSeries series1 = t1.ChartData.Series[1];
            series1.Type = ChartType.StackedLineWithMarkers;
            series1.PlotOnSecondAxis = true;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Center;
            ////设置第二个系列表
            //IChartSeries series2 = t1.ChartData.Series[2];
            //series2.PlotOnSecondAxis = true;
            //series2.Type = ChartType.StackedLineWithMarkers;
            //series2.Labels.DefaultDataLabelFormat.ShowValue = true;
            //series2.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
        }

        /// <summary>
        /// 竞品复地--第一个图表
        /// </summary>
        /// <param name="sld"></param>
        /// <param name="dt"></param>
        /// <param name="index"></param>
        public static void Chart_jp_fudi_chart1(ISlide sld, System.Data.DataTable dt, int index)
        {
            IChart t1 = (IChart)sld.Shapes[index];
            string range = "Sheet1!$A$1:$" + Base_ColumnsHelper.GET_INDEX(dt.Columns.Count) + "$" + (dt.Rows.Count + 1);

            //实例化图表数据表
            IChartDataWorkbook fact = t1.ChartData.ChartDataWorkbook;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMajorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMaxValue = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinorUnit = true;
            //t1.Axes.SecondaryVerticalAxis.IsAutomaticMinValue = true;

            t1.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = Office_TableToWork.GetWorkBooxFromDataTable(dt);
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            t1.ChartData.WriteWorkbookStream(mem);

            t1.ChartData.SetRange(range);

            ///第一列
            IChartSeries series = t1.ChartData.Series[0];
            series.Type = ChartType.ClusteredColumn;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
            ///第二列
            IChartSeries series1 = t1.ChartData.Series[1];
            series1.PlotOnSecondAxis = true;
            series1.Type = ChartType.StackedLineWithMarkers;
            series1.Labels.DefaultDataLabelFormat.ShowValue = true;
            series1.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;

        }
    }
}
