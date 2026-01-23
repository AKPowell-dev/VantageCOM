using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Gantt
{
	private struct XD
	{
		public float A;

		public float B;

		public bool A;

		public float C;

		public int A;

		public int B;
	}

	private static readonly string m_A = VH.A(75353);

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		checked
		{
			XlCalculation calc = default(XlCalculation);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				DateTime dateTime = DateTime.Parse(DateTime.Now.ToShortDateString());
				if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					Range range = (Range)application.Selection;
					XD xD = new XD
					{
						A = 6f,
						B = 4f,
						C = 60f
					};
					int num = QuickCharts.InputColor();
					QuickCharts.PrepareExcel(application, ref calc);
					Axis axis2;
					Axis axis;
					Series series;
					Range range2;
					Range range4;
					Range range6;
					Range range8;
					try
					{
						Worksheet worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(range.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						ChartObject chartObject = QuickCharts.AddChart(worksheet, xD.A, xD.B);
						chartObject.Placement = XlPlacement.xlFreeFloating;
						int num2 = chartObject.BottomRightCell.Row + 1;
						Chart chart = chartObject.Chart;
						chart.ChartType = XlChartType.xlLine;
						QuickCharts.RequireAxes(chart);
						axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
						axis2 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
						range2 = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[num2, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[num2 + 9, 11]));
						int num3 = 0;
						Range range3 = range2;
						((Range)range3.Cells[1, 1]).Value2 = VH.A(75055);
						((Range)range3.Cells[1, 2]).Value2 = VH.A(57265);
						((Range)range3.Cells[1, 3]).Value2 = VH.A(75064);
						((Range)range3.Cells[1, 4]).Value2 = VH.A(75071);
						((Range)range3.Cells[1, 5]).Value2 = VH.A(75092);
						((Range)range3.Cells[1, 6]).Value2 = VH.A(75109);
						((Range)range3.Cells[1, 7]).Value2 = VH.A(75118);
						((Range)range3.Cells[1, 8]).Value2 = VH.A(75135);
						((Range)range3.Cells[1, 9]).Value2 = VH.A(75152);
						((Range)range3.Cells[1, 10]).Value2 = VH.A(75173);
						((Range)range3.Cells[1, 11]).Value2 = VH.A(75192);
						range4 = (Range)range3.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)];
						Range range5 = range4;
						((Range)range5.Cells[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.ToOADate();
						((Range)range5.Cells[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(7.0).ToOADate();
						((Range)range5.Cells[4, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(20.0).ToOADate();
						((Range)range5.Cells[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(20.0).ToOADate();
						((Range)range5.Cells[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(40.0).ToOADate();
						((Range)range5.Cells[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(43.0).ToOADate();
						((Range)range5.Cells[8, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(43.0).ToOADate();
						((Range)range5.Cells[9, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(50.0).ToOADate();
						((Range)range5.Cells[10, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(60.0).ToOADate();
						_ = null;
						range4.get_Offset((object)1, (object)0).get_Resize(Operators.SubtractObject(range4.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Color = num;
						range6 = (Range)range3.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)];
						Range range7 = range6;
						((Range)range7.Cells[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(14.0).ToOADate();
						((Range)range7.Cells[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(28.0).ToOADate();
						((Range)range7.Cells[4, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(34.0).ToOADate();
						((Range)range7.Cells[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(48.0).ToOADate();
						((Range)range7.Cells[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(54.0).ToOADate();
						((Range)range7.Cells[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(71.0).ToOADate();
						((Range)range7.Cells[8, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(77.0).ToOADate();
						((Range)range7.Cells[9, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(84.0).ToOADate();
						((Range)range7.Cells[10, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = dateTime.AddDays(84.0).ToOADate();
						_ = null;
						range6.get_Offset((object)1, (object)0).get_Resize(Operators.SubtractObject(range6.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Color = num;
						range8 = (Range)range3.Columns[4, RuntimeHelpers.GetObjectValue(Missing.Value)];
						Range range9 = range8;
						((Range)range9.Cells[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 1;
						((Range)range9.Cells[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 1;
						((Range)range9.Cells[4, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.75;
						((Range)range9.Cells[5, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.75;
						((Range)range9.Cells[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.5;
						((Range)range9.Cells[7, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.25;
						((Range)range9.Cells[8, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.35;
						((Range)range9.Cells[9, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0.25;
						((Range)range9.Cells[10, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0;
						_ = null;
						Range range10 = range8.get_Offset((object)1, (object)0).get_Resize(Operators.SubtractObject(range8.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value));
						range10.Font.Color = num;
						range10.NumberFormat = VH.A(75211);
						_ = null;
						int num4 = Conversions.ToInteger(range3.Rows.CountLarge);
						for (int i = 2; i <= num4; i++)
						{
							num3++;
							object instance = range3.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)];
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 1 }, null, null, null)).Value2 = VH.A(75220) + num3;
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 5 }, null, null, null)).Formula = VH.A(48936) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 3 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(13778) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 2 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 6 }, null, null, null)).Formula = VH.A(48936) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 5 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75231) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 4 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 7 }, null, null, null)).Formula = VH.A(48936) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 5 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(13778) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 6 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 8 }, null, null, null)).Formula = VH.A(57636) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 4 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75234) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 11 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75241);
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 9 }, null, null, null)).Formula = VH.A(57636) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 4 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75254) + ((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 11 }, null, null, null)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75241);
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 10 }, null, null, null)).Formula = VH.A(75261);
							((Range)NewLateBinding.LateGet(instance, null, VH.A(62391), new object[1] { 11 }, null, null, null)).Formula = VH.A(75272) + ((Range)range2.Cells[1, 11]).get_Address((object)1, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75295);
							instance = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							Range range11 = ((Range)range3.Rows[RuntimeHelpers.GetObjectValue(range3.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0);
							((Range)range11.Cells[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).Formula = VH.A(75306);
							((Range)range11.Cells[8, RuntimeHelpers.GetObjectValue(Missing.Value)]).Formula = VH.A(75261);
							((Range)range11.Cells[9, RuntimeHelpers.GetObjectValue(Missing.Value)]).Formula = VH.A(75261);
							((Range)range11.Cells[10, RuntimeHelpers.GetObjectValue(Missing.Value)]).Value2 = 0;
							_ = null;
							range3 = null;
							string shortDatePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
							range4.get_Offset((object)1, (object)0).get_Resize(Operators.SubtractObject(range4.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)).NumberFormat = shortDatePattern;
							range6.get_Offset((object)1, (object)0).NumberFormat = shortDatePattern;
							string xValues = VH.A(48936) + ((Range)range2.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2.Columns[8, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57333), new object[2]
							{
								Operators.AddObject(range2.Rows.CountLarge, 1),
								Missing.Value
							}, null, null, null)), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.XValues = xValues;
							series.HasDataLabels = false;
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2.Columns[9, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57333), new object[2]
							{
								Operators.AddObject(range2.Rows.CountLarge, 1),
								Missing.Value
							}, null, null, null)), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.XValues = xValues;
							series.HasDataLabels = false;
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2.Columns[10, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57333), new object[2]
							{
								Operators.AddObject(range2.Rows.CountLarge, 1),
								Missing.Value
							}, null, null, null)), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							Series series2 = series;
							series2.XValues = xValues;
							series2.Format.Line.Visible = MsoTriState.msoFalse;
							series2.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
							series2.HasDataLabels = true;
							Microsoft.Office.Interop.Excel.DataLabels obj = (Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
							obj.Position = XlDataLabelPosition.xlLabelPositionAbove;
							obj.ShowCategoryName = true;
							obj.ShowValue = false;
							_ = null;
							series2.HasErrorBars = true;
							series2.ErrorBars.EndStyle = XlEndStyleCap.xlNoCap;
							series2.ErrorBar(XlErrorBarDirection.xlY, XlErrorBarInclude.xlErrorBarIncludePlusValues, XlErrorBarType.xlErrorBarTypeFixedValue, Operators.SubtractObject(range2.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value));
							series2.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludeNone, XlErrorBarType.xlErrorBarTypeFixedValue, 0, RuntimeHelpers.GetObjectValue(Missing.Value));
							_ = null;
							Axis axis3 = axis;
							axis3.HasMajorGridlines = true;
							axis3.MinimumScale = new DateTime(dateTime.Year, dateTime.Month, 1).ToOADate();
							axis3.MaximumScale = new DateTime(dateTime.AddDays(84.0).Year, dateTime.AddDays(84.0).Month + 1, 1).ToOADate();
							axis3.MajorUnit = 14.0;
							axis3.MinorUnit = 7.0;
							axis3.TickLabels.NumberFormat = VH.A(75323);
							_ = null;
							Axis axis4 = axis2;
							axis4.HasMajorGridlines = false;
							axis4.Crosses = XlAxisCrosses.xlAxisCrossesMinimum;
							_ = null;
							xValues = VH.A(48936) + ((Range)range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize(Operators.SubtractObject(range2.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(range2.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.ChartType = XlChartType.xlBarStacked;
							series.XValues = xValues;
							series.HasDataLabels = false;
							try
							{
								series.Format.Fill.Visible = MsoTriState.msoFalse;
								series.Format.Line.Visible = MsoTriState.msoFalse;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(range2.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.ChartType = XlChartType.xlBarStacked;
							series.XValues = xValues;
							series.HasDataLabels = false;
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(range2.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.ChartType = XlChartType.xlBarStacked;
							series.XValues = xValues;
							series.HasDataLabels = false;
							((_Chart)chart).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary, (object)true);
							axis.Crosses = XlAxisCrosses.xlAxisCrossesMaximum;
							((Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary)).Crosses = XlAxisCrosses.xlAxisCrossesAutomatic;
							Axis obj2 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, XlAxisGroup.xlSecondary);
							obj2.ReversePlotOrder = true;
							obj2.TickLabelSpacing = 1;
							obj2.MajorTickMark = XlTickMark.xlTickMarkNone;
							obj2.MinorTickMark = XlTickMark.xlTickMarkNone;
							_ = null;
							Axis axis5 = axis2;
							axis5.ReversePlotOrder = true;
							axis5.Crosses = XlAxisCrosses.xlAxisCrossesMaximum;
							axis5.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone;
							axis5.MaximumScale = Conversions.ToDouble(Operators.SubtractObject(range2.Rows.CountLarge, 1));
							axis5.MinimumScale = 0.0;
							axis5.MaximumScaleIsAuto = false;
							axis5.MinimumScaleIsAuto = false;
							_ = null;
							Axis obj3 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary);
							obj3.MinimumScale = axis.MinimumScale;
							obj3.MaximumScale = axis.MaximumScale;
							obj3.MajorUnit = axis.MajorUnit;
							obj3.MinorUnit = axis.MinorUnit;
							obj3.TickLabels.NumberFormat = axis.TickLabels.NumberFormat;
							obj3.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone;
							_ = null;
							Series obj4 = (Series)chart.SeriesCollection(1);
							obj4.Format.Line.Visible = MsoTriState.msoFalse;
							obj4.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
							obj4.MarkerForegroundColor = ((Series)chart.SeriesCollection(4)).Format.Fill.ForeColor.RGB;
							_ = null;
							Series obj5 = (Series)chart.SeriesCollection(2);
							obj5.Format.Line.Visible = MsoTriState.msoFalse;
							obj5.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
							obj5.MarkerForegroundColor = ((Series)chart.SeriesCollection(5)).Format.Fill.ForeColor.RGB;
							_ = null;
							chart.HasLegend = false;
							((ChartGroup)chart.LineGroups(1)).GapWidth = (int)Math.Round(xD.C);
							chart.DisplayBlanksAs = XlDisplayBlanksAs.xlNotPlotted;
							break;
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						B(ex4.Message);
						clsReporting.LogException(ex4);
						ProjectData.ClearProjectError();
					}
					QuickCharts.RestoreExcel(application, calc);
					QuickCharts.LogActivity(VH.A(75330));
					axis2 = null;
					axis = null;
					series = null;
					range2 = null;
					range4 = null;
					range6 = null;
					range8 = null;
					range = null;
				}
				application = null;
				return;
			}
		}
	}

	private static void A(string A)
	{
		System.Windows.Forms.MessageBox.Show(A, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}

	private static void B(string A)
	{
		System.Windows.Forms.MessageBox.Show(A, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
	}
}
