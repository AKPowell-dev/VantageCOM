using System;
using System.Collections;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Waterfall
{
	private enum IE
	{
		A,
		B,
		C,
		D
	}

	private struct JE
	{
		public float A;

		public float B;

		public bool A;

		public int A;

		public IE A;

		public int B;

		public int C;

		public int D;

		public bool B;
	}

	private struct KE
	{
		public Range A;

		public Range B;
	}

	private static readonly string m_A = VH.A(81188);

	private static readonly int m_A = 8;

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			XlCalculation calc = default(XlCalculation);
			Range D = default(Range);
			while (true)
			{
				switch (3)
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
						Range A = (Range)application.Selection;
						JE jE = Waterfall.A(ref A, ref B);
						if (B)
						{
							if (Operators.ConditionalCompareObjectNotEqual(A.Columns.CountLarge, 2, TextCompare: false))
							{
								Forms.WarningMessage(VH.A(79920));
								B = false;
							}
							else
							{
								try
								{
									enumerator = ((IEnumerable)NewLateBinding.LateGet(A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(62391), new object[0], null, null, null)).GetEnumerator();
									while (true)
									{
										if (enumerator.MoveNext())
										{
											Range range = (Range)enumerator.Current;
											if (!Operators.ConditionalCompareObjectNotEqual(range.Formula, string.Empty, TextCompare: false) || Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
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
												Forms.WarningMessage(VH.A(80041));
												B = false;
												break;
											}
											break;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_016a;
											}
											continue;
											end_IL_016a:
											break;
										}
										break;
									}
								}
								finally
								{
									if (enumerator is IDisposable)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											(enumerator as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							if (B)
							{
								int f = QuickCharts2.InputColor();
								int g = QuickCharts2.LinkColor();
								QuickCharts2.PrepareExcel(application, ref calc);
								ChartObject chartObject;
								Chart chart;
								Range d;
								Series series;
								try
								{
									Worksheet worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(A.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									chartObject = QuickCharts2.AddChart(worksheet, jE.A, jE.B);
									chartObject.Placement = XlPlacement.xlFreeFloating;
									int val = chartObject.BottomRightCell.Row + 2;
									chart = chartObject.Chart;
									chart.ChartType = XlChartType.xlColumnStacked;
									((_Chart)chart).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary, (object)true);
									KE kE = Waterfall.A(worksheet, chartObject, jE);
									val = Math.Max(val, Waterfall.m_A + 3);
									Waterfall.A(jE, worksheet, ref A, ref D, val, f, g);
									A = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[val, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[Operators.SubtractObject(Operators.AddObject(val, A.Rows.CountLarge), 1), 2]));
									A.Select();
									int num = Conversions.ToInteger(Operators.AddObject(A.Rows.CountLarge, 1));
									D = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[val - 1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[Operators.AddObject(val - 1, A.Rows.CountLarge), 11]));
									Range range2 = D;
									((Range)range2.Cells[1, 3]).Value2 = VH.A(80154);
									((Range)range2.Cells[1, 4]).Value2 = VH.A(51175);
									((Range)range2.Cells[1, 5]).Value2 = VH.A(80163);
									((Range)range2.Cells[1, 6]).Value2 = VH.A(80174);
									((Range)range2.Cells[1, 7]).Value2 = VH.A(80187);
									((Range)range2.Cells[1, 8]).Value2 = VH.A(80200);
									((Range)range2.Cells[1, 9]).Value2 = VH.A(80217);
									((Range)range2.Cells[1, 10]).Value2 = VH.A(80232);
									((Range)range2.Cells[1, 11]).Value2 = VH.A(49999);
									_ = null;
									Range obj = (Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)];
									Range range3 = obj.get_Offset((object)0, (object)4);
									string text = obj.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									string text2 = kE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									int num2 = num;
									for (int i = 2; i <= num2; i++)
									{
										Range range4 = D;
										string text3 = ((Range)range4.Cells[i, 2]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text4 = ((Range)range4.Cells[i, 3]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text5 = ((Range)range4.Cells[i, 4]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text6 = ((Range)range4.Cells[i, 5]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text7 = ((Range)range4.Cells[i, 6]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text8 = ((Range)range4.Cells[i, 7]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										string text9 = kE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										((Range)range4.Cells[i, 3]).Formula = VH.A(80247) + text3 + VH.A(80270) + text + VH.A(54423) + text9 + VH.A(80307) + text8 + VH.A(75557);
										((Range)range4.Cells[i, 4]).Formula = VH.A(80247) + text3 + VH.A(80270) + text + VH.A(54423) + text9 + VH.A(80314) + text + VH.A(80339) + range3.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(80354) + text8 + VH.A(13778) + text3 + VH.A(39904);
										((Range)range4.Cells[i, 5]).Formula = VH.A(80247) + text3 + VH.A(80270) + text + VH.A(54423) + text9 + VH.A(80314) + text + VH.A(80339) + range3.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(80354) + text8 + VH.A(39904);
										((Range)range4.Cells[i, 6]).Formula = VH.A(72965) + text + VH.A(80375) + text + VH.A(80339) + range3.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(80414) + text7 + VH.A(80447);
										((Range)range4.Cells[i, 7]).Formula = VH.A(72965) + text + VH.A(80468) + text8 + VH.A(80491) + text3 + VH.A(80522) + text3 + VH.A(57650);
										((Range)range4.Cells[i, 8]).Formula = VH.A(57636) + text2 + VH.A(80531) + text4 + VH.A(80522) + text4 + VH.A(80552) + text5 + VH.A(2378) + text6 + VH.A(80575);
										((Range)range4.Cells[i, 9]).Formula = VH.A(75501) + text2 + VH.A(80592) + text2 + VH.A(80607) + text4 + VH.A(80522) + text8 + VH.A(80636) + text3 + VH.A(80522) + text3 + VH.A(80661) + text4 + VH.A(80690) + text8 + VH.A(80711) + text5 + VH.A(2378) + text6 + VH.A(80575);
										((Range)range4.Cells[i, 10]).Formula = VH.A(57636) + text2 + VH.A(80724) + text4 + VH.A(80745) + text4 + VH.A(80760) + text4 + VH.A(80767) + text3 + VH.A(80745) + text3 + VH.A(80796) + text5 + VH.A(2378) + text6 + VH.A(80811);
										((Range)range4.Cells[i, 11]).Formula = VH.A(80852) + text3 + VH.A(80522) + text3 + VH.A(2378) + text4 + VH.A(39904);
										range4 = null;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										((Range)A.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)9).NumberFormat = RuntimeHelpers.GetObjectValue(((Range)D.Cells[2, 2]).NumberFormat);
										string text10 = VH.A(48936) + ((Range)A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
										d = ((Range)D.Columns[11, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num - 1), RuntimeHelpers.GetObjectValue(Missing.Value));
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										series.ChartType = XlChartType.xlColumnStacked;
										series.XValues = text10;
										series.HasDataLabels = false;
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[4, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Series series2 = series;
										series2.ChartType = XlChartType.xlLine;
										series2.Format.Line.Visible = MsoTriState.msoFalse;
										series2.XValues = text10;
										series2.HasDataLabels = false;
										_ = null;
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Series series3 = series;
										series3.ChartType = XlChartType.xlLine;
										series3.Format.Line.Visible = MsoTriState.msoFalse;
										series3.XValues = text10;
										series3.HasDataLabels = false;
										_ = null;
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[8, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Waterfall.A(series, XlDataLabelPosition.xlLabelPositionCenter, text10, d);
										((Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(System.Drawing.Color.White);
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[9, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Waterfall.A(series, XlDataLabelPosition.xlLabelPositionAbove, text10, d);
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[10, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Waterfall.A(series, XlDataLabelPosition.xlLabelPositionBelow, text10, d);
										series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(RuntimeHelpers.GetObjectValue(D.Columns[7, RuntimeHelpers.GetObjectValue(Missing.Value)]), XlRowCol.xlColumns, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
										Series series4 = series;
										series4.ChartType = XlChartType.xlXYScatter;
										series4.AxisGroup = XlAxisGroup.xlSecondary;
										series4.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
										series4.XValues = VH.A(48936) + ((Range)D.Columns[6, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).get_Resize((object)(num - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
										series4.HasDataLabels = false;
										_ = null;
										ChartGroup obj2 = (ChartGroup)chart.LineGroups(1);
										obj2.HasUpDownBars = true;
										obj2.GapWidth = jE.A;
										ChartFormat format = obj2.UpBars.Format;
										format.Fill.ForeColor.RGB = jE.C;
										format.Line.Visible = MsoTriState.msoFalse;
										_ = null;
										ChartFormat format2 = obj2.DownBars.Format;
										format2.Fill.ForeColor.RGB = jE.D;
										format2.Line.Visible = MsoTriState.msoFalse;
										_ = null;
										_ = null;
										((_Chart)chart).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary, (object)false);
										((Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)).TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow;
										Chart chart2 = chart;
										((Series)chart2.SeriesCollection(1)).Format.Fill.ForeColor.RGB = jE.B;
										QuickCharts2.CleanUpChart(chart);
										((ChartGroup)chart2.ChartGroups(1)).GapWidth = jE.A;
										chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
										chart2.HasLegend = false;
										chart2.ChartArea.Select();
										_ = null;
										double num3 = ((Microsoft.Office.Interop.Excel.Point)((Series)chart.SeriesCollection(1)).Points(1)).Left - chart.PlotArea.InsideLeft;
										double value = 1.0 - num3 * 1.0 / (((Microsoft.Office.Interop.Excel.Point)series.Points(2)).Left - ((Microsoft.Office.Interop.Excel.Point)series.Points(1)).Left);
										value = Math.Round(value, 6);
										Series series5 = series;
										series5.HasErrorBars = true;
										series5.ErrorBars.EndStyle = XlEndStyleCap.xlNoCap;
										if (jE.B)
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
											series5.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludeBoth, XlErrorBarType.xlErrorBarTypeFixedValue, value, RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										else
										{
											series5.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludeNone, XlErrorBarType.xlErrorBarTypeFixedValue, value, RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										series5.ErrorBar(XlErrorBarDirection.xlY, XlErrorBarInclude.xlErrorBarIncludeNone, XlErrorBarType.xlErrorBarTypeFixedValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										series5 = null;
										break;
									}
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									Forms.ErrorMessage(ex2.Message);
									clsReporting.LogException(ex2);
									ProjectData.ClearProjectError();
								}
								QuickCharts2.RestoreExcel(application, calc);
								chartObject = null;
								chart = null;
								D = null;
								d = null;
								series = null;
								QuickCharts2.LogActivity(VH.A(80869));
							}
						}
						A = null;
					}
				}
				application = null;
				return;
			}
		}
	}

	private static void A(Series A, XlDataLabelPosition B, string C, Range D)
	{
		A.ChartType = XlChartType.xlLine;
		A.AxisGroup = XlAxisGroup.xlSecondary;
		A.Format.Line.Visible = MsoTriState.msoFalse;
		A.XValues = C;
		A.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
		A.HasLeaderLines = false;
		A.HasDataLabels = true;
		((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Position = B;
		DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), D);
		_ = null;
	}

	private static JE A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfWaterfall wpfWaterfall2 = new wpfWaterfall();
		QuickCharts2.HandleColorChange(wpfWaterfall2.btnColorBase);
		QuickCharts2.HandleColorChange(wpfWaterfall2.btnColorIncrease);
		QuickCharts2.HandleColorChange(wpfWaterfall2.btnColorDecrease);
		wpfWaterfall2.Range = A;
		wpfWaterfall2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfWaterfall2.numChartWidth, wpfWaterfall2.numChartHeight);
		wpfWaterfall2.numGapWidth.Value = QuickCharts2.GetGapWidth(xmlDoc);
		wpfWaterfall2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		switch ((IE)Conversions.ToInteger(xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(55388)).InnerText))
		{
		case IE.A:
			wpfWaterfall2.radLabelCenter.IsChecked = true;
			break;
		case IE.B:
			wpfWaterfall2.radLabelAbove.IsChecked = true;
			break;
		case IE.C:
			wpfWaterfall2.radLabelAboveBelow.IsChecked = true;
			break;
		default:
			wpfWaterfall2.radLabelNone.IsChecked = true;
			break;
		}
		QuickCharts2.SetButtonColor(wpfWaterfall2.btnColorBase, xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80900)).InnerText);
		QuickCharts2.SetButtonColor(wpfWaterfall2.btnColorIncrease, xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80919)).InnerText);
		QuickCharts2.SetButtonColor(wpfWaterfall2.btnColorDecrease, xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80946)).InnerText);
		wpfWaterfall2.chkErrorBars.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80973)).InnerText);
		wpfWaterfall2.ShowDialog();
		JE result = default(JE);
		if (wpfWaterfall2.DialogResult.HasValue)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (wpfWaterfall2.DialogResult.Value)
			{
				A = wpfWaterfall2.Range;
				System.Windows.Media.Color buttonColor;
				System.Windows.Media.Color buttonColor2;
				System.Windows.Media.Color buttonColor3;
				XmlNode xmlNode;
				checked
				{
					result = new JE
					{
						A = wpfWaterfall2.chkPreserveFormulas.IsChecked.Value,
						A = (float)wpfWaterfall2.numChartWidth.Value.Value,
						B = (float)wpfWaterfall2.numChartHeight.Value.Value,
						A = (int)Math.Round(wpfWaterfall2.numGapWidth.Value.Value)
					};
					if (wpfWaterfall2.radLabelCenter.IsChecked == true)
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
						result.A = IE.A;
					}
					else if (wpfWaterfall2.radLabelAbove.IsChecked == true)
					{
						result.A = IE.B;
					}
					else if (wpfWaterfall2.radLabelAboveBelow.IsChecked == true)
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
						result.A = IE.C;
					}
					else
					{
						result.A = IE.D;
					}
					buttonColor = QuickCharts2.GetButtonColor(wpfWaterfall2.btnColorBase);
					buttonColor2 = QuickCharts2.GetButtonColor(wpfWaterfall2.btnColorIncrease);
					buttonColor3 = QuickCharts2.GetButtonColor(wpfWaterfall2.btnColorDecrease);
					result.B = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor.R, buttonColor.G, buttonColor.B));
					result.C = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor2.R, buttonColor2.G, buttonColor2.B));
					result.D = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor3.R, buttonColor3.G, buttonColor3.B));
					result.B = wpfWaterfall2.chkErrorBars.IsChecked.Value;
					QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfWaterfall2.numChartWidth.Value.Value));
					QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfWaterfall2.numChartHeight.Value.Value));
					QuickCharts2.SetGapWidth(ref xmlDoc, (int)Math.Round(wpfWaterfall2.numGapWidth.Value.Value));
					QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfWaterfall2.chkPreserveFormulas.IsChecked.Value);
					xmlNode = xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(55388));
				}
				int a = (int)result.A;
				xmlNode.InnerText = a.ToString();
				xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80900)).InnerText = QuickCharts2.Color2RGB(buttonColor);
				xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80919)).InnerText = QuickCharts2.Color2RGB(buttonColor2);
				xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80946)).InnerText = QuickCharts2.Color2RGB(buttonColor3);
				xmlDoc.SelectSingleNode(Waterfall.m_A + VH.A(80973)).InnerText = (0 - (wpfWaterfall2.chkErrorBars.IsChecked.Value ? 1 : 0)).ToString();
				KH.A.SaveSettings(xmlDoc);
				goto IL_061b;
			}
		}
		B = false;
		goto IL_061b;
		IL_061b:
		wpfWaterfall2 = null;
		xmlDoc = null;
		return result;
	}

	private static void A(JE A, Worksheet B, ref Range C, ref Range D, int E, int F, int G)
	{
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.A)
			{
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				((Range)B.Cells[E, 1]).PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num3 = num;
				for (int i = 1; i <= num3; i++)
				{
					int num4 = num2;
					Range range;
					Range range2;
					for (int j = 1; j <= num4; range = null, range2 = null, j++)
					{
						range = (Range)C.Cells[i, j];
						range2 = (Range)B.Cells[E - 1 + i, j];
						if (Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
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
								switch (3)
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
							Range obj = (Range)B.Cells[E - 1 + i, j];
							obj.Font.Color = G;
							obj.NumberFormat = numberFormat;
							_ = null;
							continue;
						}
						if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)))
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
						range2.Font.Color = F;
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
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[Operators.SubtractObject(Operators.AddObject(E, C.Rows.CountLarge), 1), RuntimeHelpers.GetObjectValue(C.Columns.CountLarge)]));
			}
			else
			{
				Range range4 = C;
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range obj2 = (Range)B.Cells[E, 1];
				obj2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj2.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num, num2])).Font.Color = G;
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
							switch (5)
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
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0534;
						}
						continue;
						end_IL_0534:
						break;
					}
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
			application.CutCopyMode = (XlCutCopyMode)0;
			application = null;
		}
	}

	private static KE A(Worksheet A, ChartObject B, JE C)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		KE result = default(KE);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 2]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[Waterfall.m_A, column + 2])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			Range range2 = (Range)A.Cells[num, num2];
			Range range3 = range2;
			range3.Value2 = true;
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
			CheckBox obj2 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				range3.Left,
				range3.Top,
				range3.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2).Width,
				range3.Height
			}, null, null, null);
			obj2.Text = VH.A(81000);
			obj2.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj2.Value = Constants.xlBoth;
			range3 = null;
			result.B = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 2, num2], VH.A(81049));
			range2 = (Range)A.Cells[num + 3, num2];
			Range range4 = range2;
			DropDown dropDown = (DropDown)NewLateBinding.LateGet(A.DropDowns(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(61140), new object[4] { range4.Left, range4.Top, range4.Width, range4.Height }, null, null, null);
			dropDown.AddItem(VH.A(61854), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(61147), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(56560), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(61174), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.LinkedCell = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			switch (C.A)
			{
			case IE.A:
				dropDown.ListIndex = 1;
				break;
			case IE.C:
				dropDown.ListIndex = 2;
				break;
			case IE.B:
				dropDown.ListIndex = 3;
				break;
			default:
				dropDown.ListIndex = 4;
				break;
			}
			try
			{
				range4.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			range4 = null;
			result.A = range2;
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[Waterfall.m_A, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[Waterfall.m_A, column + 2])).Borders[XlBordersIndex.xlEdgeBottom];
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
		((_Worksheet)worksheet).get_Range((object)VH.A(78844), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81078);
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81091);
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81104);
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61854);
		((_Worksheet)worksheet).get_Range((object)VH.A(78364), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81117);
		((_Worksheet)worksheet).get_Range((object)VH.A(78390), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81130);
		((_Worksheet)worksheet).get_Range((object)VH.A(81143), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(81148);
		((_Worksheet)worksheet).get_Range((object)VH.A(81161), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(75064);
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 30;
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 8;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -5;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12;
		((_Worksheet)worksheet).get_Range((object)VH.A(61439), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 4;
		((_Worksheet)worksheet).get_Range((object)VH.A(61451), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -8;
		((_Worksheet)worksheet).get_Range((object)VH.A(61467), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 14;
		Range range = ((_Worksheet)worksheet).get_Range((object)VH.A(81166), RuntimeHelpers.GetObjectValue(Missing.Value));
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
		Range result = ((_Worksheet)worksheet).get_Range((object)VH.A(81177), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = null;
		return result;
	}
}
