using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Butterfly
{
	private enum BE
	{
		A,
		B
	}

	private enum CE
	{
		A = 4,
		B = 3,
		C = 2,
		D = 1
	}

	private struct DE
	{
		public float A;

		public float B;

		public bool A;

		public int A;

		public CE A;

		public BE A;

		public bool B;

		public int B;

		public bool C;
	}

	private struct EE
	{
		public Range A;

		public Range B;
	}

	private static readonly string m_A = VH.A(78466);

	private static readonly int m_A = 4;

	private static readonly int B = 5;

	private static readonly int C = 6;

	private static readonly int D = 7;

	private static readonly int E = 8;

	private static readonly int F = 7;

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		checked
		{
			XlCalculation calc = default(XlCalculation);
			string text = default(string);
			Range range5 = default(Range);
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				bool B = true;
				DE dE;
				Range range;
				Range A;
				if (application.Selection is Range)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						A = (Range)application.Selection;
						dE = Butterfly.A(ref A, ref B);
						if (B)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							range = A;
							if (Operators.ConditionalCompareObjectNotEqual(range.Columns.CountLarge, 3, TextCompare: false))
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								Forms.WarningMessage(VH.A(77084));
								B = false;
							}
							else
							{
								if (Butterfly.A((Range)range.Cells[1, 2]))
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										break;
									}
									if (Butterfly.A((Range)range.Cells[1, 3]))
									{
										if (application.WorksheetFunction.CountIf(((Range)range.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2), VH.A(77399)) > 0.0)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												break;
											}
											Forms.WarningMessage(VH.A(77404));
											B = false;
										}
										goto IL_01dc;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										break;
									}
								}
								Forms.WarningMessage(VH.A(77209));
								B = false;
							}
							goto IL_01dc;
						}
						goto IL_2295;
					}
				}
				goto IL_2297;
				IL_2297:
				application = null;
				return;
				IL_2295:
				A = null;
				goto IL_2297;
				IL_01dc:
				range = null;
				if (B)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					int e = QuickCharts2.InputColor();
					int f = QuickCharts2.LinkColor();
					QuickCharts2.PrepareExcel(application, ref calc);
					ChartObject chartObject;
					Chart chart;
					Axis axis;
					Axis axis2;
					Worksheet worksheet;
					try
					{
						int num;
						if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(((Range)A.Cells[1, 2]).Value2)))
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
							num = ((!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(((Range)A.Cells[1, 3]).Value2))) ? 1 : 0);
						}
						else
						{
							num = 0;
						}
						worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(A.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						chartObject = QuickCharts2.AddChart(worksheet, dE.A, dE.B);
						chartObject.Placement = XlPlacement.xlFreeFloating;
						int val;
						if (num != 0)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							val = chartObject.BottomRightCell.Row + 1;
						}
						else
						{
							val = chartObject.BottomRightCell.Row + 2;
						}
						chart = chartObject.Chart;
						QuickCharts2.RequireAxes(chart);
						axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
						axis2 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
						Butterfly.A(D: (num == 0) ? Math.Max(val, F + 3) : Math.Max(val, F + 2), A: dE, B: worksheet, C: ref A, E: e, F: f);
						if (num == 0)
						{
							A = A.get_Offset((object)(-1), (object)0).get_Resize(Operators.AddObject(A.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						int num2 = Conversions.ToInteger(A.Rows.CountLarge);
						if (num == 0)
						{
							((Range)A.Cells[1, 2]).Value2 = VH.A(77461);
							((Range)A.Cells[1, 3]).Value2 = VH.A(77478);
						}
						try
						{
							Range range2 = ((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), (object)2);
							if (!Information.IsDBNull(RuntimeHelpers.GetObjectValue(range2.NumberFormat)))
							{
								range2.NumberFormat = range2.NumberFormat.ToString().Replace(VH.A(77495), "").Replace(VH.A(77500), "")
									.Replace(VH.A(77505), "");
							}
							range2 = null;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (dE.A == BE.A)
						{
							if (Operators.ConditionalCompareObjectGreater(((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).Rows.CountLarge, application.WorksheetFunction.CountA(RuntimeHelpers.GetObjectValue(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
								dE.B = true;
							}
							EE eE = Butterfly.A(worksheet, chartObject, dE);
							chart.ChartType = XlChartType.xlBarStacked100;
							A = A.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.AddObject(A.Columns.CountLarge, 5));
							Range range3 = A;
							((Range)range3.Cells[1, Butterfly.m_A]).Value2 = VH.A(77510);
							((Range)range3.Cells[1, Butterfly.B]).Formula = VH.A(48936) + ((Range)range3.Cells[1, 2]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Range)range3.Cells[1, C]).Value2 = VH.A(77529);
							((Range)range3.Cells[1, D]).Formula = VH.A(48936) + ((Range)range3.Cells[1, 3]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Range)range3.Cells[1, E]).Value2 = VH.A(77536);
							range3 = null;
							int num3 = 0;
							int num4 = 0;
							List<int> list = new List<int>();
							int num5 = num2;
							for (int i = 2; i <= num5; i++)
							{
								Range range4 = (Range)A.Cells[i, 2];
								if (range4.Row > num4)
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
									if (range4.get_End(XlDirection.xlDown).Row > num4)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											break;
										}
										num3 = range4.Row;
										if (Operators.ConditionalCompareObjectNotEqual(range4.Formula, string.Empty, TextCompare: false))
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												break;
											}
											num4 = range4.get_End(XlDirection.xlDown).Row;
										}
										else
										{
											num4 = range4.get_End(XlDirection.xlDown).get_End(XlDirection.xlDown).Row;
										}
										range5 = ((_Worksheet)worksheet).get_Range((object)(VH.A(77555) + num3 + VH.A(77558) + num4), RuntimeHelpers.GetObjectValue(Missing.Value));
										text = VH.A(77563) + range5.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
									}
								}
								if (Operators.ConditionalCompareObjectEqual(range4.Formula, string.Empty, TextCompare: false))
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										break;
									}
									list.Add(i - 1);
								}
								range4 = null;
								string text2 = ((Range)A.Cells[i, 2]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								string text3 = ((Range)A.Cells[i, 3]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								string text4 = VH.A(57636) + eE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378);
								text4 = text4 + VH.A(77574) + ((Range)A.Cells[i, Butterfly.B]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								text4 += VH.A(2378);
								text4 = Conversions.ToString(Operators.ConcatenateObject(text4, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(77583), NewLateBinding.LateGet(range5.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 1, 1 }, null, null, null)), VH.A(68994)), text2), VH.A(39904)), text)));
								text4 += VH.A(39904);
								string text5 = VH.A(57636) + eE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378);
								text5 = text5 + VH.A(77574) + ((Range)A.Cells[i, D]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								text5 += VH.A(2378);
								text5 = Conversions.ToString(Operators.ConcatenateObject(text5, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(77583), NewLateBinding.LateGet(range5.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 1, 1 }, null, null, null)), VH.A(68994)), text3), VH.A(39904)), text)));
								text5 += VH.A(39904);
								Range range6 = A;
								((Range)range6.Cells[i, Butterfly.m_A]).Formula = text4;
								((Range)range6.Cells[i, Butterfly.B]).Formula = VH.A(77602) + text2 + text + VH.A(77613) + text2 + VH.A(39904);
								((Range)range6.Cells[i, C]).Formula = VH.A(48936) + eE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								((Range)range6.Cells[i, D]).Formula = VH.A(77602) + text3 + text + VH.A(77613) + text3 + VH.A(39904);
								((Range)range6.Cells[i, E]).Formula = text5;
								range6 = null;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							Series series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[Butterfly.m_A, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series.Format.Fill.Visible = MsoTriState.msoFalse;
							series.Format.Line.Visible = MsoTriState.msoFalse;
							Series series2 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[Butterfly.B, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							Series series3 = series2;
							if (KH.A.ChartSeriesColors.Count > 1)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								series3.Format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[0]);
							}
							series3.Format.Line.Visible = MsoTriState.msoFalse;
							if (dE.A != CE.D)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									break;
								}
								series3.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								Microsoft.Office.Interop.Excel.DataLabels obj = (Microsoft.Office.Interop.Excel.DataLabels)series3.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
								obj.Position = unchecked((XlDataLabelPosition)dE.A);
								obj.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
								_ = null;
								DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series3.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)));
							}
							series3 = null;
							Series series4 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[C, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							Series series5 = series4;
							series5.Format.Fill.Visible = MsoTriState.msoFalse;
							series5.Format.Line.Visible = MsoTriState.msoFalse;
							series5.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							((Microsoft.Office.Interop.Excel.DataLabels)series5.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionCenter;
							DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series5.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)));
							_ = null;
							Series series6 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[D, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							Series series7 = series6;
							if (KH.A.ChartSeriesColors.Count > 1)
							{
								series7.Format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[1]);
							}
							series7.Format.Line.Visible = MsoTriState.msoFalse;
							if (dE.A != CE.D)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									break;
								}
								series7.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								Microsoft.Office.Interop.Excel.DataLabels obj2 = (Microsoft.Office.Interop.Excel.DataLabels)series7.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
								obj2.Position = unchecked((XlDataLabelPosition)dE.A);
								obj2.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
								_ = null;
								DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series7.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)));
							}
							series7 = null;
							Series series8 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[E, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series8.Format.Fill.Visible = MsoTriState.msoFalse;
							series8.Format.Line.Visible = MsoTriState.msoFalse;
							if (dE.A == CE.A)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideEnd;
								((Microsoft.Office.Interop.Excel.DataLabels)series6.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideBase;
							}
							else if (dE.A == CE.B)
							{
								((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideBase;
								((Microsoft.Office.Interop.Excel.DataLabels)series6.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = XlDataLabelPosition.xlLabelPositionInsideEnd;
							}
							if (chart.HasLegend)
							{
								Legend legend = chart.Legend;
								((LegendEntry)legend.LegendEntries(5)).Delete();
								((LegendEntry)legend.LegendEntries(3)).Delete();
								((LegendEntry)legend.LegendEntries(1)).Delete();
								_ = null;
							}
							Butterfly.A(axis2);
							double num6 = 0.0;
							int count = ((Microsoft.Office.Interop.Excel.DataLabels)series4.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
							for (int j = 1; j <= count; j++)
							{
								DataLabel dataLabel = (DataLabel)series4.DataLabels(j);
								if (dataLabel.Width > num6)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										break;
									}
									if (!list.Contains(j))
									{
										num6 = dataLabel.Width;
									}
								}
								dataLabel = null;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
							Microsoft.Office.Interop.Excel.Point point;
							if (!list.Contains(1))
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								point = (Microsoft.Office.Interop.Excel.Point)series4.Points(1);
							}
							else
							{
								point = (Microsoft.Office.Interop.Excel.Point)series4.Points(2);
							}
							application.Calculation = XlCalculation.xlCalculationSemiautomatic;
							Range b = eE.B;
							while (num6 + 4.0 > point.Width)
							{
								b.Value2 = Operators.AddObject(b.Value2, 1);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							while (num6 + 4.0 < point.Width)
							{
								b.Value2 = Operators.SubtractObject(b.Value2, 1);
							}
							b = null;
							application.Calculation = XlCalculation.xlCalculationManual;
							using (List<int>.Enumerator enumerator = list.GetEnumerator())
							{
								while (enumerator.MoveNext())
								{
									int current = enumerator.Current;
									DataLabel obj3 = (DataLabel)series4.DataLabels(current);
									Microsoft.Office.Interop.Excel.Font font;
									(font = obj3.Font).Size = Operators.AddObject(font.Size, 2);
									obj3.Font.Bold = true;
									obj3.Format.TextFrame2.WordWrap = MsoTriState.msoFalse;
									_ = null;
									_ = null;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_19e8;
									}
									continue;
									end_IL_19e8:
									break;
								}
							}
							if (list.Any())
							{
								int count2 = ((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
								for (int k = 1; k <= count2; k++)
								{
									((DataLabel)series2.DataLabels(k)).NumberFormat = Conversions.ToString(((Range)A.Cells[k + 1, 2]).NumberFormat);
									((DataLabel)series6.DataLabels(k)).NumberFormat = Conversions.ToString(((Range)A.Cells[k + 1, 3]).NumberFormat);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							range5 = null;
							series2 = null;
							series6 = null;
							series4 = null;
							point = null;
						}
						else
						{
							chart.ChartType = XlChartType.xlBarClustered;
							((Range)A.Cells[1, 4]).Formula = VH.A(48936) + ((Range)A.Cells[1, 2]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							int num7 = num2;
							for (int l = 2; l <= num7; l++)
							{
								((Range)A.Cells[l, 4]).Formula = VH.A(77630) + ((Range)A.Cells[l, 2]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							string xValues = VH.A(48936) + ((Range)A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							Series series9 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[4, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series9.XValues = xValues;
							if (KH.A.ChartSeriesColors.Count > 1)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								series9.Format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[0]);
							}
							series9.Format.Line.Visible = MsoTriState.msoFalse;
							if (dE.A != CE.D)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								series9.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series9.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
								dataLabels.Position = unchecked((XlDataLabelPosition)dE.A);
								if (dE.A != CE.C)
								{
									dataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
								}
								dataLabels = null;
								DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series9.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num2 - 1), RuntimeHelpers.GetObjectValue(Missing.Value)));
							}
							else
							{
								series9.HasDataLabels = false;
							}
							series9 = null;
							Series series10 = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add((Range)A.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)], XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
							series10.XValues = xValues;
							if (KH.A.ChartSeriesColors.Count > 1)
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
								series10.Format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[1]);
							}
							series10.Format.Line.Visible = MsoTriState.msoFalse;
							if (dE.A != CE.D)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								series10.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								Microsoft.Office.Interop.Excel.DataLabels dataLabels2 = (Microsoft.Office.Interop.Excel.DataLabels)series10.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
								dataLabels2.Position = unchecked((XlDataLabelPosition)dE.A);
								if (dE.A != CE.C)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										break;
									}
									dataLabels2.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
								}
								dataLabels2 = null;
							}
							else
							{
								series10.HasDataLabels = false;
							}
							series10 = null;
							((ChartGroup)chart.ChartGroups(1)).Overlap = 100;
							axis2.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow;
							string[] array = Strings.Split(Conversions.ToString(((Range)A.Cells[2, 2]).NumberFormat), VH.A(77635));
							string text6 = array[0] + VH.A(77635) + array[0];
							if (array.Length > 2)
							{
								text6 = text6 + VH.A(77635) + array[2];
							}
							axis.TickLabels.NumberFormat = text6;
						}
						if (dE.C)
						{
							Butterfly.A(axis);
							axis.HasMajorGridlines = false;
						}
						else
						{
							axis.HasMajorGridlines = true;
							axis.MajorGridlines.Format.Line.ForeColor.RGB = QuickCharts2.GRIDLINES_COLOR;
						}
						Axis axis3 = axis2;
						axis3.ReversePlotOrder = true;
						axis3.Crosses = XlAxisCrosses.xlAxisCrossesMaximum;
						axis3.HasMajorGridlines = false;
						_ = null;
						QuickCharts2.CleanUpChart(chart);
						Chart chart2 = chart;
						((ChartGroup)chart2.ChartGroups(1)).GapWidth = dE.A;
						chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
						chart2.HasLegend = true;
						chart2.Legend.Position = XlLegendPosition.xlLegendPositionTop;
						chart2.ChartArea.Select();
						_ = null;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						Forms.ErrorMessage(ex4.Message);
						clsReporting.LogException(ex4);
						ProjectData.ClearProjectError();
					}
					QuickCharts2.RestoreExcel(application, calc);
					chartObject = null;
					chart = null;
					axis = null;
					axis2 = null;
					worksheet = null;
					QuickCharts2.LogActivity(VH.A(77638));
				}
				goto IL_2295;
			}
		}
	}

	private static bool A(Range A)
	{
		int num;
		if (Conversions.ToBoolean(Operators.CompareObjectNotEqual(A.Formula, string.Empty, TextCompare: false)))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			num = ((!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(A.Value2))) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		return Conversions.ToBoolean((byte)num != 0);
	}

	private static void A(Axis A)
	{
		A.Format.Line.Visible = MsoTriState.msoFalse;
		A.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone;
	}

	private static DE A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfButterfly wpfButterfly2 = new wpfButterfly();
		wpfButterfly2.Range = A;
		wpfButterfly2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfButterfly2.numChartWidth, wpfButterfly2.numChartHeight);
		wpfButterfly2.numGapWidth.Value = QuickCharts2.GetGapWidth(xmlDoc);
		wpfButterfly2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		BE bE = (BE)Conversions.ToInteger(xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(76386)).InnerText);
		if (bE != BE.A)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (bE == BE.B)
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
				wpfButterfly2.radLabelsLeft.IsChecked = true;
			}
		}
		else
		{
			wpfButterfly2.radLabelsMiddle.IsChecked = true;
		}
		switch ((CE)Conversions.ToInteger(xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77669)).InnerText))
		{
		case CE.A:
			wpfButterfly2.cbxLabelPosn.SelectedValue = VH.A(72273);
			break;
		case CE.B:
			wpfButterfly2.cbxLabelPosn.SelectedValue = VH.A(72296);
			break;
		case CE.C:
			wpfButterfly2.cbxLabelPosn.SelectedValue = VH.A(72317);
			break;
		default:
			wpfButterfly2.cbxLabelPosn.SelectedValue = VH.A(61174);
			break;
		}
		wpfButterfly2.chkCenterLabels.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77704)).InnerText);
		wpfButterfly2.chkHideHorizAxis.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77745)).InnerText);
		wpfButterfly2.ShowDialog();
		DE result = default(DE);
		if (wpfButterfly2.DialogResult.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (wpfButterfly2.DialogResult.Value)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				A = wpfButterfly2.Range;
				XmlNode xmlNode;
				checked
				{
					result = new DE
					{
						A = wpfButterfly2.chkPreserveFormulas.IsChecked.Value,
						A = (float)wpfButterfly2.numChartWidth.Value.Value,
						B = (float)wpfButterfly2.numChartHeight.Value.Value,
						A = (int)Math.Round(wpfButterfly2.numGapWidth.Value.Value)
					};
					if (wpfButterfly2.radLabelsLeft.IsChecked == true)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						result.A = BE.B;
					}
					else
					{
						result.A = BE.A;
					}
					string left = wpfButterfly2.cbxLabelPosn.SelectedValue.ToString();
					if (Operators.CompareString(left, VH.A(72273), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						if (Operators.CompareString(left, VH.A(72296), TextCompare: false) != 0)
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
							if (Operators.CompareString(left, VH.A(72317), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								result.A = CE.D;
							}
							else
							{
								result.A = CE.C;
							}
						}
						else
						{
							result.A = CE.B;
						}
					}
					else
					{
						result.A = CE.A;
					}
					result.B = wpfButterfly2.chkCenterLabels.IsChecked.Value;
					result.C = wpfButterfly2.chkHideHorizAxis.IsChecked.Value;
					QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfButterfly2.numChartWidth.Value.Value));
					QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfButterfly2.numChartHeight.Value.Value));
					QuickCharts2.SetGapWidth(ref xmlDoc, (int)Math.Round(wpfButterfly2.numGapWidth.Value.Value));
					QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfButterfly2.chkPreserveFormulas.IsChecked.Value);
					xmlNode = xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(76386));
				}
				int a = (int)result.A;
				xmlNode.InnerText = a.ToString();
				XmlNode xmlNode2 = xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77669));
				a = (int)result.A;
				xmlNode2.InnerText = a.ToString();
				xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77704)).InnerText = (0 - (result.B ? 1 : 0)).ToString();
				xmlDoc.SelectSingleNode(Butterfly.m_A + VH.A(77745)).InnerText = (0 - (result.C ? 1 : 0)).ToString();
				KH.A.SaveSettings(xmlDoc);
				goto IL_05a4;
			}
		}
		B = false;
		goto IL_05a4;
		IL_05a4:
		wpfButterfly2 = null;
		xmlDoc = null;
		return result;
	}

	private static void A(DE A, Worksheet B, ref Range C, int D, int E, int F)
	{
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.A)
			{
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				((Range)B.Cells[D, 1]).PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num3 = num;
				for (int i = 1; i <= num3; i++)
				{
					int num4 = num2;
					Range range;
					Range range2;
					for (int j = 1; j <= num4; range = null, range2 = null, j++)
					{
						range = (Range)C.Cells[i, j];
						range2 = (Range)B.Cells[D - 1 + i, j];
						if (Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							range2.Clear();
							continue;
						}
						if (Conversions.ToBoolean(range.HasFormula))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							string numberFormat = Conversions.ToString(range.NumberFormat);
							string formula = range.Formula.ToString();
							Range range3 = ((_Worksheet)C.Worksheet).get_Range((object)VH.A(60932), RuntimeHelpers.GetObjectValue(Missing.Value));
							range3.Formula = formula;
							range3.Cut(range2);
							_ = null;
							Range obj = (Range)B.Cells[D - 1 + i, j];
							obj.Font.Color = F;
							obj.NumberFormat = numberFormat;
							_ = null;
							continue;
						}
						if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
						{
							if (!KH.A.AutoColorText)
							{
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						range2.Font.Color = E;
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D, 2]), RuntimeHelpers.GetObjectValue(B.Cells[D + num - 1, num2]));
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D, 1]), RuntimeHelpers.GetObjectValue(B.Cells[Operators.SubtractObject(Operators.AddObject(D, C.Rows.CountLarge), 1), RuntimeHelpers.GetObjectValue(C.Columns.CountLarge)]));
			}
			else
			{
				Range range4 = C;
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range obj2 = (Range)B.Cells[D, 1];
				obj2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj2.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[D + num - 1, num2]));
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[D, 1]), RuntimeHelpers.GetObjectValue(B.Cells[D + num, num2])).Font.Color = F;
				C = (Range)application.Selection;
				int num5 = Conversions.ToInteger(range4.Rows.CountLarge);
				for (int k = 1; k <= num5; k++)
				{
					int num6 = Conversions.ToInteger(range4.Columns.CountLarge);
					for (int l = 1; l <= num6; l++)
					{
						if (!Operators.ConditionalCompareObjectEqual(((Range)range4.Cells[k, l]).Formula, string.Empty, TextCompare: false))
						{
							continue;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
						((Range)C.Cells[k, l]).Value2 = "";
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0518;
						}
						continue;
						end_IL_0518:
						break;
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			application.CutCopyMode = (XlCutCopyMode)0;
			application = null;
		}
	}

	private static EE A(Worksheet A, ChartObject B, DE C)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		EE result = default(EE);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 2]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[F, column + 2])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			Range obj2 = (Range)A.Cells[num + 2, num2];
			obj2.Value2 = VH.A(77772);
			obj2.EntireColumn.AutoFit();
			_ = null;
			Range range2 = (Range)A.Cells[num, num2];
			Range range3 = range2;
			range3.Value2 = C.B;
			try
			{
				range3.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			CheckBox obj3 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range3.Left, range3.Top, range3.Width, range3.Height }, null, null, null);
			obj3.Text = VH.A(77821);
			obj3.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			int num3;
			if (!C.B)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				num3 = -4146;
			}
			else
			{
				num3 = 1;
			}
			obj3.Value = unchecked((Microsoft.Office.Interop.Excel.Constants)num3);
			range3 = null;
			result.A = range2;
			range2 = (Range)A.Cells[num + 2, num2 + 1];
			Range range4 = range2;
			range4.Value2 = 20;
			range4.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			Validation validation = range4.Validation;
			validation.InputMessage = VH.A(77866);
			validation.ShowError = false;
			_ = null;
			_ = null;
			QuickCharts2.FormatOptionsInput(range2);
			result.B = range2;
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[F, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[F, column + 2])).Borders[XlBordersIndex.xlEdgeBottom];
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2;
			border.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			_ = null;
			_ = null;
			range2 = null;
			return result;
		}
	}

	public static Range Example(Worksheet ws)
	{
		Worksheet worksheet = ws;
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78302);
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78317);
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78342);
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78355);
		((_Worksheet)worksheet).get_Range((object)VH.A(78364), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78369);
		((_Worksheet)worksheet).get_Range((object)VH.A(78390), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78395);
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78406);
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 205.329;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 160.396;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 148.233;
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 112.812;
		((_Worksheet)worksheet).get_Range((object)VH.A(61439), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 95.457;
		((_Worksheet)worksheet).get_Range((object)VH.A(61451), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 47.171;
		((_Worksheet)worksheet).get_Range((object)VH.A(57617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(78425);
		((_Worksheet)worksheet).get_Range((object)VH.A(76950), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 173.639;
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 141.772;
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 125.723;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 111.206;
		((_Worksheet)worksheet).get_Range((object)VH.A(61512), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 88.229;
		((_Worksheet)worksheet).get_Range((object)VH.A(61517), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 38.921;
		Range range = ((_Worksheet)worksheet).get_Range((object)VH.A(78444), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range.NumberFormat = QuickCharts2.CURRENCY_FORMAT_1;
			range.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		range = null;
		Range result = ((_Worksheet)worksheet).get_Range((object)VH.A(78455), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = null;
		return result;
	}
}
