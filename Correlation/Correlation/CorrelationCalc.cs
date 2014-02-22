using MathNet.Numerics.Statistics;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationLib
{
    public class CorrelationCalc
    {
        private object MisValue = System.Reflection.Missing.Value;

        private string DependentTitle { get; set; }

        private double[] Dependents { get; set; }

        private List<KeyValuePair<string, double[]>> Independents = new List<KeyValuePair<string, double[]>>();

        public void SetDependents(double[] dependents, string title = "Dependent")
        {
            this.Dependents = dependents;
            this.DependentTitle = title;
        }

        public void AddIndependents(double[] independents, string title = "Independent")
        {
            this.Independents.Add(new KeyValuePair<string,double[]>(title, independents));
        }

        private void CheckDependents()
        {
            if (this.Dependents == null)
            {
                throw new Exception();
            }
        }

        #region Pearson
        public double[] GetPearsons()
        {
            CheckDependents();
            return this.Independents.Select(i => GetPearson(i.Value)).ToArray();
        }

        private double GetPearson(double[] independent)
        {
            return Correlation.Pearson(this.Dependents, independent);
        }
        #endregion

        #region Spearman
        public double[] GetSpearmans()
        {
            CheckDependents();
            return this.Independents.Select(i => GetSpearman(i.Value)).ToArray();
        }

        private double GetSpearman(double[] independent)
        {
            return Correlation.Spearman(this.Dependents, independent);
        }
        #endregion

        #region Graphs
        public void MakeGraphs(string dir, bool includePearson, bool includeSpearman)
        {
            Application xlApp = new ApplicationClass();
            Workbook xlWorkBook = xlApp.Workbooks.Add(MisValue);
            List<Worksheet> workSheets = PopulateWorkSheets(xlWorkBook, includePearson, includeSpearman);
            dir = SaveFile(dir, xlWorkBook);
            ReleaseResources(xlApp, xlWorkBook, workSheets);
        }

        private List<Worksheet> PopulateWorkSheets(Workbook xlWorkBook, bool includePearson, bool includeSpearman)
        {
            for (int i = 0; i < this.Independents.Count + 1; i++)
            {
                xlWorkBook.Worksheets.Add();
            }

            List<Worksheet> workSheets = CreateSummaryCharts(xlWorkBook);

            for (int i = 0; i < this.Independents.Count; i++)
            {
                Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(i + 3);
                string title = this.Independents[i].Key;
                double[] values = this.Independents[i].Value;

                double? pearson = includePearson ? GetPearson(values) as double? : null;
                double? spearman = includeSpearman ? GetSpearman(values) as double? : null;

                xlWorkSheet.Name = title;
                PopulateWorksheet(xlWorkSheet, values, title);
                PopulateChart(xlWorkSheet, title, pearson, spearman);
                workSheets.Add(xlWorkSheet);
            }

            return workSheets;
        }

        private List<Worksheet> CreateSummaryCharts(Workbook xlWorkBook)
        {
            List<Worksheet> workSheets = new List<Worksheet>();
            string[] names = this.Independents.Select(i => i.Key).ToArray();

            Worksheet w1 = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            CreateSummaryWorkSheet(w1, "Pearson Coefficients", names, this.GetPearsons());
            PopulateSummaryChart(w1, "Pearson Coefficients");
            workSheets.Add(w1);

            Worksheet w2 = (Worksheet)xlWorkBook.Worksheets.get_Item(2);
            CreateSummaryWorkSheet(w2, "Spearman Coefficients", names, this.GetSpearmans());
            PopulateSummaryChart(w2, "Spearman Coefficients");
            workSheets.Add(w2);

            return workSheets;
        }

        private void CreateSummaryWorkSheet(Worksheet xlWorkSheet, string title, string[] names, double[] values)
        {
            xlWorkSheet.Name = title;
            List<KeyValuePair<string, double>> points = new List<KeyValuePair<string, double>>();
            for (int i = 0; i < names.Length; i++)
            {
                points.Add(new KeyValuePair<string,double>(names[i], values[i]));
            }
            points = points.OrderBy(i => i.Value).ToList();

            names = points.Select(i => i.Key).ToArray();
            values = points.Select(i => i.Value).ToArray();

            xlWorkSheet.Cells[1, 1] = "Attribute";
            for (int i = 0; i < names.Length; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = names[i];
            }

            xlWorkSheet.Cells[1, 2] = title;
            for (int i = 0; i < values.Length; i++)
            {
                xlWorkSheet.Cells[i + 2, 2] = values[i];
            }
        }

        private void PopulateSummaryChart(Worksheet xlWorkSheet, string title)
        {
            ChartObjects xlCharts = (ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(200, 50, 500, 450);
            Chart chartPage = myChart.Chart;
            chartPage.ChartType = XlChartType.xlColumnClustered;
            chartPage.HasLegend = false;
            chartPage.HasTitle = true;
            chartPage.ChartTitle.Text = String.Format("The dependence of {0} as {1}", this.DependentTitle, title);

            Axis xAxis = (Axis)chartPage.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "Attributes";

            Axis yAxis = (Axis)chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = this.DependentTitle;

            Range chartRange = xlWorkSheet.get_Range("A2", String.Format("B{0}", this.Independents.Count + 1));
            chartPage.SetSourceData(chartRange, XlRowCol.xlColumns);

            Series series = (Series)chartPage.SeriesCollection(1);
            series.InvertIfNegative = true;
            series.Format.Fill.ForeColor.RGB = Microsoft.VisualBasic.Information.RGB(0, 0, 250);
            series.InvertColor = Microsoft.VisualBasic.Information.RGB(250, 0, 0);
        }

        private void PopulateWorksheet(Worksheet xlWorkSheet, double[] independents, string independentTitle)
        {
            xlWorkSheet.Cells[1, 1] = independentTitle;
            for (int i = 0; i < independents.Length; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = independents[i];
            }  

            xlWorkSheet.Cells[1, 2] = this.DependentTitle;
            for (int i = 0; i < this.Dependents.Length; i++)
            {
                xlWorkSheet.Cells[i + 2, 2] = this.Dependents[i];
            }
        }

        private void PopulateChart(Worksheet xlWorkSheet, string independentTitle, double? pearson, double? spearman)
        {
            Chart chartPage = CreateChartHeader(xlWorkSheet, independentTitle, pearson, spearman);
            CreateChartBody(xlWorkSheet, chartPage);
        }

        private Chart CreateChartHeader(Worksheet xlWorkSheet, string independentTitle, double? pearson, double? spearman)
        {
            ChartObjects xlCharts = (ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(200, 50, 500, 450);
            Chart chartPage = myChart.Chart;
            chartPage.ChartType = XlChartType.xlXYScatter;
            chartPage.HasLegend = false;
            chartPage.HasTitle = true;
            chartPage.ChartTitle.Text = String.Format("The dependence of {0} on {1}", this.DependentTitle, independentTitle);

            if (pearson.HasValue)
            {
                chartPage.ChartTitle.Text += "\nPearson: " + pearson.Value.ToString("0.000");
            }

            if (spearman.HasValue)
            {
                chartPage.ChartTitle.Text += "\nSpearman: " + spearman.Value.ToString("0.000");
            }

            Axis xAxis = (Axis)chartPage.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = independentTitle;

            Axis yAxis = (Axis)chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = this.DependentTitle;

            return chartPage;
        }

        private void CreateChartBody(Worksheet xlWorkSheet, Chart chartPage)
        {
            Range chartRange = xlWorkSheet.get_Range("A2", String.Format("B{0}", this.Dependents.Length + 1));
            chartPage.SetSourceData(chartRange, XlRowCol.xlColumns);

            Series series = (Series)chartPage.SeriesCollection(1);
            series.XValues = xlWorkSheet.get_Range("A2", String.Format("A{0}", this.Dependents.Length + 1));
            series.Values = xlWorkSheet.get_Range("B2", String.Format("B{0}", this.Dependents.Length + 1));

            Trendlines trendlines = (Trendlines)series.Trendlines(System.Type.Missing);
            Trendline trendLine = trendlines.Add(XlTrendlineType.xlLinear);
        }

        private static string SaveFile(string dir, Workbook xlWorkBook)
        {
            if (dir == null || dir == String.Empty)
            {
                dir = Directory.GetCurrentDirectory();
            }
            xlWorkBook.SaveAs(String.Format("{0}\\data.xls", dir), XlFileFormat.xlWorkbookNormal);
            return dir;
        }

        private void ReleaseResources(Application xlApp, Workbook xlWorkBook, List<Worksheet> workSheets)
        {
            xlWorkBook.Close(true);
            xlApp.Quit();
            workSheets.ForEach(i => ReleaseObject(i));
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch { obj = null; }
            finally { GC.Collect(); }
        }
        #endregion
    }
}