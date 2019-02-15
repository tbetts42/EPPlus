using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelChartTwoSeriesTest : TestBase
    {
        [TestInitialize]
        public void Initialize()
        {
            var xmlDoc = new XmlDocument();
            var xmlNsm = new XmlNamespaceManager(new NameTable());
            xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
        }

        [TestMethod]
        public void Add_Two_Charts_Many_Series_Sets_Correct_Idx()
        {
            var worksheet = _pck.Workbook.Worksheets.Add("BigChart");
            AddTestData(worksheet);

            var primaryChart = worksheet.Drawings.AddChart("bigChart", eChartType.ColumnClustered) as ExcelChart;
            primaryChart.SetPosition(50, 50);
            primaryChart.SetSize(800, 300);

            var secondaryChart = primaryChart.PlotArea.ChartTypes.Add(eChartType.Line);
            secondaryChart.UseSecondaryAxis = true;

            var rowCount = 3; // 21-23
            var colCount = 4;
            var headerRow = 20;
            var categoryColumn = 3; // Column C is category labels
            var primarySeriesMeasure = worksheet.Cells["D19"].Value.ToString();
            var secondarySeriesMeasure = worksheet.Cells["H19"].Value.ToString();
            var firstTableFirstColumn = 4; // Column D is where the first data table starts
            int firstTableLastColumn = firstTableFirstColumn + colCount - 1;
            var secondTableFirstColumn = 8;
            int secondTableLastColumn = secondTableFirstColumn + colCount - 1;

            var primarySeriesCategoryLabels = 
                worksheet.Cells[headerRow, firstTableFirstColumn, headerRow, firstTableLastColumn];

            for (int rowNum = headerRow + 1; rowNum <= headerRow + rowCount; rowNum++)
            {
                var category = worksheet.Cells[rowNum, categoryColumn].Value.ToString();
                var primarySeriesValues =
                    worksheet.Cells[rowNum, firstTableFirstColumn, rowNum, firstTableLastColumn];
                var primarySeries = primaryChart.Series.Add(primarySeriesValues, primarySeriesCategoryLabels);
                primarySeries.Header = $"{category} - {primarySeriesMeasure}";

                var secondarySeriesValues =
                    worksheet.Cells[rowNum, secondTableFirstColumn, rowNum, secondTableLastColumn];

                // adding to the secondaryChart can create duplicate IDs
                var secondarySeries = secondaryChart.Series.Add(secondarySeriesValues, primarySeriesCategoryLabels);
                // Adding to the primaryChart allows the test to pass, but that's not the test
                //var secondarySeries = primaryChart.Series.Add(secondarySeriesValues, primarySeriesCategoryLabels);
                secondarySeries.Header = $"{category} - {secondarySeriesMeasure}";
            }

            // Check XML for unique Idx and SortOrder
            var uniqueIds = new List<int>();
            foreach (ExcelChartSerie series in secondaryChart.Series)
            {
                var idx = series.GetXmlNodeInt("c:idx/@val");
                Assert.IsFalse(uniqueIds.Contains(idx), $"Two chart series have identical ID={idx}");
                uniqueIds.Add(idx);
            }

            SaveWorksheet("TwoSeries.xlsx");
        }

        /// <summary>
        /// Creates two related tables, basically a de-normalized pivot table.
        /// Series labels are in Row 20, from column D to column G
        /// First table starts at column D (4).
        /// Second table starts at column H (8).
        /// </summary>
        /// <param name="ws"></param>
        private static void AddTestData(ExcelWorksheet ws)
        {
            ws.Cells["D19"].Value = "Average Price";
            ws.Cells["D19:G19"].Merge = true;
            ws.Cells["D19:G19"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["D20"].Value = "2015";
            ws.Cells["E20"].Value = "2016";
            ws.Cells["F20"].Value = "2017";
            ws.Cells["G20"].Value = "2018";

            ws.Cells["H19"].Value = "Revenue";
            ws.Cells["H19:K19"].Merge = true;
            ws.Cells["H19:K19"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["H20"].Value = "2015";
            ws.Cells["I20"].Value = "2016";
            ws.Cells["J20"].Value = "2017";
            ws.Cells["K20"].Value = "2018";

            // Categories
            ws.Cells["C20"].Value = "Category";
            ws.Cells["C21"].Value = "Toys";
            ws.Cells["C22"].Value = "Games";
            ws.Cells["C23"].Value = "Bikes";

            // Toys
            ws.Cells["D21"].Value = 100;
            ws.Cells["E21"].Value = 110;
            ws.Cells["F21"].Value = 115;
            ws.Cells["G21"].Value = 110;

            ws.Cells["H21"].Value = 5;
            ws.Cells["I21"].Value = 7;
            ws.Cells["J21"].Value = 6;
            ws.Cells["K21"].Value = 8;

            // Games
            ws.Cells["D22"].Value = 120;
            ws.Cells["E22"].Value = 130;
            ws.Cells["F22"].Value = 110;
            ws.Cells["G22"].Value = 150;

            ws.Cells["H22"].Value = 11;
            ws.Cells["I22"].Value = 7;
            ws.Cells["J22"].Value = 9;
            ws.Cells["K22"].Value = 12;

            // Bikes
            ws.Cells["D23"].Value = 230;
            ws.Cells["E23"].Value = 235;
            ws.Cells["F23"].Value = 200;
            ws.Cells["G23"].Value = 230;

            ws.Cells["H23"].Value = 15;
            ws.Cells["I23"].Value = 12;
            ws.Cells["J23"].Value = 14;
            ws.Cells["K23"].Value = 17;

        }
    }
}
