using System;
using System.Collections;
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
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class StackedColumn
{
	private struct GE
	{
		public float A;

		public float B;

		public bool A;

		public int A;

		public bool B;

		public bool C;

		public bool D;

		public bool E;
	}

	private struct HE
	{
		public Range A;

		public Range B;

		public Range C;

		public Range D;

		public Range E;

		public Range F;

		public Range G;
	}

	private static readonly string m_A = VH.A(79863);

	private static readonly int m_A = 8;

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		checked
		{
			XlCalculation calc = default(XlCalculation);
			IEnumerator enumerator = default(IEnumerator);
			Range D = default(Range);
			while (true)
			{
				switch (5)
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
						switch (7)
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
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						Range A = (Range)application.Selection;
						GE gE = StackedColumn.A(ref A, ref B);
						if (B && B)
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
							int num = QuickCharts2.DefaultColor();
							int f = QuickCharts2.InputColor();
							int g = QuickCharts2.LinkColor();
							QuickCharts2.PrepareExcel(application, ref calc);
							Worksheet worksheet;
							ChartObject chartObject;
							Chart chart;
							try
							{
								worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(A.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								chartObject = QuickCharts2.AddChart(worksheet, gE.A, gE.B);
								chartObject.Placement = XlPlacement.xlFreeFloating;
								int val = chartObject.BottomRightCell.Row + 1;
								chart = chartObject.Chart;
								chart.ChartType = XlChartType.xlColumnStacked;
								HE hE = StackedColumn.A(worksheet, chartObject, gE);
								val = Math.Max(val, StackedColumn.m_A + 2);
								StackedColumn.A(gE, worksheet, A, ref D, val, f, g);
								int num2 = 0;
								int num3 = Conversions.ToInteger(D.Rows.CountLarge);
								int num4 = Conversions.ToInteger(D.Columns.CountLarge);
								int num5 = num4;
								for (int i = 2; i <= num5; i++)
								{
									try
									{
										if (!Operators.ConditionalCompareObjectEqual(((Range)D.Cells[num3, i]).Value2, application.WorksheetFunction.Sum(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(D.Columns[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57333), new object[2]
										{
											num3 - 1,
											Missing.Value
										}, null, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
										{
											continue;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											num2++;
											break;
										}
									}
									catch (InvalidCastException ex)
									{
										ProjectData.SetProjectError(ex);
										InvalidCastException ex2 = ex;
										num2++;
										ProjectData.ClearProjectError();
									}
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									if (num2 < num4 - 1)
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
										Range range = ((Range)D.Cells[num3, 1]).get_Offset((object)1, (object)0);
										range.Value2 = VH.A(57315);
										range.Font.Color = num;
										_ = null;
										int num6 = num4;
										for (int j = 2; j <= num6; j++)
										{
											Range range2 = ((Range)D.Cells[num3, j]).get_Offset((object)1, (object)0);
											range2.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(54448), NewLateBinding.LateGet(D.Columns[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904));
											range2.Font.Color = num;
											range2 = null;
										}
										D = D.get_Resize((object)(num3 + 1), RuntimeHelpers.GetObjectValue(Missing.Value));
										num3 = Conversions.ToInteger(D.Rows.CountLarge);
									}
									int row = ((Range)D.Rows[num3, RuntimeHelpers.GetObjectValue(Missing.Value)]).Row;
									int num7 = num3 - 1;
									for (int k = 1; k <= num7; k++)
									{
										((Range)D.Cells[k, 1]).get_Offset((object)num3, (object)0).Value2 = Operators.ConcatenateObject(((Range)D.Cells[k, 1]).Text, VH.A(79112));
										int num8 = num4;
										for (int l = 2; l <= num8; l++)
										{
											string text = VH.A(79125) + ((Range)D.Cells[k, l]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75498) + ((Range)worksheet.Cells[row, l]).get_Address((object)1, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
											string text2 = text + VH.A(79142);
											string text3 = text + VH.A(79153);
											Range range3 = ((Range)D.Cells[k, l]).get_Offset((object)num3, (object)0);
											string text4 = VH.A(75501);
											text4 = text4 + VH.A(75516) + hE.D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79166) + ((Range)D.Cells[k, l]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79177) + hE.F.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
											text4 = text4 + VH.A(75528) + hE.E.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text2 + VH.A(75525) + hE.G.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
											text4 += VH.A(75539);
											text4 = text4 + VH.A(75550) + hE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(79182) + ((Range)D.Cells[k, l]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(73137) + ((Range)D.Cells[k, l]).NumberFormat.ToString().Replace(VH.A(39830), VH.A(62386)) + VH.A(79195);
											text4 = text4 + VH.A(75566) + hE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75575) + hE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75702) + text3 + VH.A(79208);
											text4 += VH.A(39904);
											range3.Formula = text4;
											range3 = null;
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												goto end_IL_0b19;
											}
											continue;
											end_IL_0b19:
											break;
										}
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										((Range)D.Cells[num3 * 2 - 1, 1]).get_Offset((object)1, (object)0).Value2 = VH.A(79233);
										int num9 = num4;
										for (int m = 2; m <= num9; m++)
										{
											Range range4 = (Range)D.Cells[num3, m];
											Range range5 = ((Range)D.Cells[num3 * 2 - 1, m]).get_Offset((object)1, (object)0);
											range5.Formula = VH.A(57636) + hE.C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75557);
											range5.NumberFormat = RuntimeHelpers.GetObjectValue(range4.NumberFormat);
											range5 = null;
											range4 = null;
										}
										int num10 = 0;
										string xValues = VH.A(48936) + D.get_Offset((object)(-1), (object)1).get_Resize((object)1, (object)(num4 - 1)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
										Range range6 = JH.A(D, application);
										if (range6 != null)
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
											{
												enumerator = range6.Rows.GetEnumerator();
												try
												{
													while (enumerator.MoveNext())
													{
														Range range7 = (Range)enumerator.Current;
														Series series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Add(range7, XlRowCol.xlRows, true, false, RuntimeHelpers.GetObjectValue(Missing.Value));
														Series series2 = series;
														series2.XValues = xValues;
														series2.HasDataLabels = true;
														if (range7.Row == row)
														{
															series2.ChartType = XlChartType.xlLine;
															series2.Format.Line.Visible = MsoTriState.msoFalse;
															Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
															dataLabels.Position = XlDataLabelPosition.xlLabelPositionAbove;
															if (gE.E)
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
																dataLabels.Format.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
															}
															dataLabels = null;
															DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)D.Rows[num3, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)num3, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(num4 - 1)));
														}
														else
														{
															DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), range7.get_Offset((object)num3, (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(num4 - 1)));
															if (KH.A.ChartSeriesColors.Any())
															{
																num10++;
																try
																{
																	int count = KH.A.ChartSeriesColors.Count;
																	int num11 = unchecked(num10 % count);
																	Microsoft.Office.Interop.Excel.ColorFormat foreColor = series.Format.Fill.ForeColor;
																	string text5;
																	if (num11 <= 0)
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
																		text5 = KH.A.ChartSeriesColors[count - 1];
																	}
																	else
																	{
																		text5 = KH.A.ChartSeriesColors[num11 - 1];
																	}
																	foreColor.RGB = clsColors.RGB2Ole(text5);
																	series.Format.Line.Visible = MsoTriState.msoFalse;
																}
																catch (Exception ex3)
																{
																	ProjectData.SetProjectError(ex3);
																	Exception ex4 = ex3;
																	ProjectData.ClearProjectError();
																}
															}
														}
														series2 = null;
														series = null;
													}
													while (true)
													{
														switch (1)
														{
														case 0:
															break;
														default:
															goto end_IL_0fc7;
														}
														continue;
														end_IL_0fc7:
														break;
													}
												}
												finally
												{
													IDisposable disposable = enumerator as IDisposable;
													if (disposable != null)
													{
														disposable.Dispose();
													}
												}
											}
											range6 = null;
										}
										if (chart.HasLegend)
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
											((LegendEntry)chart.Legend.LegendEntries(num3)).Delete();
										}
										QuickCharts2.CleanUpChart(chart);
										Chart chart2 = chart;
										((ChartGroup)chart2.ChartGroups(1)).GapWidth = gE.A;
										chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
										chart2.ChartArea.Select();
										_ = null;
										hE.F.NumberFormat = RuntimeHelpers.GetObjectValue(((Range)D.Cells[1, 2]).NumberFormat);
										break;
									}
									break;
								}
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								Forms.ErrorMessage(ex6.Message);
								clsReporting.LogException(ex6);
								ProjectData.ClearProjectError();
							}
							QuickCharts2.RestoreExcel(application, calc);
							worksheet = null;
							chartObject = null;
							chart = null;
							D = null;
							QuickCharts2.LogActivity(VH.A(79258));
						}
						A = null;
					}
				}
				application = null;
				return;
			}
		}
	}

	private static GE A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfStackedColumn wpfStackedColumn2 = new wpfStackedColumn();
		wpfStackedColumn2.Range = A;
		wpfStackedColumn2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfStackedColumn2.numChartWidth, wpfStackedColumn2.numChartHeight);
		wpfStackedColumn2.numGapWidth.Value = QuickCharts2.GetGapWidth(xmlDoc);
		wpfStackedColumn2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		wpfStackedColumn2.chkShowAmount.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79299)).InnerText);
		wpfStackedColumn2.chkShowPercent.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79320)).InnerText);
		wpfStackedColumn2.chkTotalsShow.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79343)).InnerText);
		wpfStackedColumn2.chkTotalsBold.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79364)).InnerText);
		wpfStackedColumn2.ShowDialog();
		GE result = default(GE);
		if (wpfStackedColumn2.DialogResult.HasValue && wpfStackedColumn2.DialogResult.Value)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = wpfStackedColumn2.Range;
			checked
			{
				result = new GE
				{
					A = wpfStackedColumn2.chkPreserveFormulas.IsChecked.Value,
					A = (float)wpfStackedColumn2.numChartWidth.Value.Value,
					B = (float)wpfStackedColumn2.numChartHeight.Value.Value,
					A = (int)Math.Round(wpfStackedColumn2.numGapWidth.Value.Value),
					B = wpfStackedColumn2.chkShowAmount.IsChecked.Value,
					C = wpfStackedColumn2.chkShowPercent.IsChecked.Value,
					D = wpfStackedColumn2.chkTotalsShow.IsChecked.Value,
					E = wpfStackedColumn2.chkTotalsBold.IsChecked.Value
				};
				QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfStackedColumn2.numChartWidth.Value.Value));
				QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfStackedColumn2.numChartHeight.Value.Value));
				QuickCharts2.SetGapWidth(ref xmlDoc, (int)Math.Round(wpfStackedColumn2.numGapWidth.Value.Value));
				QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfStackedColumn2.chkPreserveFormulas.IsChecked.Value);
			}
			xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79299)).InnerText = (0 - (wpfStackedColumn2.chkShowAmount.IsChecked.Value ? 1 : 0)).ToString();
			xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79320)).InnerText = (0 - (wpfStackedColumn2.chkShowPercent.IsChecked.Value ? 1 : 0)).ToString();
			xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79343)).InnerText = (0 - (wpfStackedColumn2.chkTotalsShow.IsChecked.Value ? 1 : 0)).ToString();
			xmlDoc.SelectSingleNode(StackedColumn.m_A + VH.A(79364)).InnerText = (0 - (wpfStackedColumn2.chkTotalsBold.IsChecked.Value ? 1 : 0)).ToString();
			KH.A.SaveSettings(xmlDoc);
		}
		else
		{
			B = false;
		}
		wpfStackedColumn2 = null;
		xmlDoc = null;
		return result;
	}

	private static void A(GE A, Worksheet B, Range C, ref Range D, int E, int F, int G)
	{
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.A)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
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
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
							range2.Clear();
							continue;
						}
						if (Conversions.ToBoolean(range.HasFormula))
						{
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
								switch (3)
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
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
			}
			else
			{
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range obj2 = (Range)B.Cells[E, 1];
				obj2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj2.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				obj2.Clear();
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2])).Font.Color = G;
			}
			application.CutCopyMode = (XlCutCopyMode)0;
			application = null;
		}
	}

	private static HE A(Worksheet A, ChartObject B, GE C)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		HE result = default(HE);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 3]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[StackedColumn.m_A, column + 3])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num, num2], VH.A(79397));
			Range range2 = (Range)A.Cells[num + 1, num2];
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
			CheckBox obj2 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range3.Left, range3.Top, range3.Width, range3.Height }, null, null, null);
			obj2.Text = VH.A(76504);
			obj2.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj2.Value = (C.B ? Constants.xlBoth : Constants.xlOff);
			range3 = null;
			result.A = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			Range range4 = range2;
			range4.Value2 = C.C;
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
			CheckBox obj3 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range4.Left, range4.Top, range4.Width, range4.Height }, null, null, null);
			obj3.Text = VH.A(76529);
			obj3.LinkedCell = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			int num3;
			if (!C.C)
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
			obj3.Value = unchecked((Constants)num3);
			range4 = null;
			result.B = range2;
			range2 = (Range)A.Cells[num + 3, num2];
			Range range5 = range2;
			range5.Value2 = C.D;
			try
			{
				range5.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			CheckBox obj4 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range5.Left, range5.Top, range5.Width, range5.Height }, null, null, null);
			obj4.Text = VH.A(79414);
			obj4.LinkedCell = range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj4.Value = (C.D ? Constants.xlBoth : Constants.xlOff);
			range5 = null;
			result.C = range2;
			num2 = column + 2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num, num2], VH.A(79441));
			range2 = (Range)A.Cells[num + 1, num2];
			Range range6 = range2;
			range6.Value2 = false;
			try
			{
				range6.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			CheckBox obj5 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range6.Left, range6.Top, range6.Width, range6.Height }, null, null, null);
			obj5.Text = VH.A(76750);
			obj5.LinkedCell = range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj5.Value = Constants.xlOff;
			range6 = null;
			result.D = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			Range range7 = range2;
			range7.Value2 = false;
			try
			{
				range7.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
			CheckBox obj6 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range7.Left, range7.Top, range7.Width, range7.Height }, null, null, null);
			obj6.Text = VH.A(79480);
			obj6.LinkedCell = range7.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj6.Value = Constants.xlOff;
			range7 = null;
			result.E = range2;
			num2 = column + 3;
			((Range)A.Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]).ColumnWidth = 15;
			range2 = (Range)A.Cells[num + 1, num2];
			range2.Value2 = 0;
			range2.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			Validation validation = range2.Validation;
			validation.InputMessage = VH.A(79515);
			validation.ShowError = false;
			_ = null;
			QuickCharts2.FormatOptionsInput(range2);
			result.F = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			range2.Value2 = 0.05;
			range2.Validation.Add(XlDVType.xlValidateInputOnly, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			Validation validation2 = range2.Validation;
			validation2.InputMessage = VH.A(79683);
			validation2.ShowError = false;
			_ = null;
			try
			{
				range2.NumberFormat = VH.A(75697);
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			QuickCharts2.FormatOptionsInput(range2);
			result.G = range2;
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[StackedColumn.m_A, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[StackedColumn.m_A, column + 3])).Borders[XlBordersIndex.xlEdgeBottom];
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
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76882);
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76898);
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76918);
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 2015;
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 67;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 49;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 41;
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 30;
		((_Worksheet)worksheet).get_Range((object)VH.A(57617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 2016;
		((_Worksheet)worksheet).get_Range((object)VH.A(76950), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 56;
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 43;
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 38;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 22;
		((_Worksheet)worksheet).get_Range((object)VH.A(76893), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 2017;
		((_Worksheet)worksheet).get_Range((object)VH.A(76955), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 51;
		((_Worksheet)worksheet).get_Range((object)VH.A(61532), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 39;
		((_Worksheet)worksheet).get_Range((object)VH.A(61546), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 31;
		((_Worksheet)worksheet).get_Range((object)VH.A(61551), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 18;
		((_Worksheet)worksheet).get_Range((object)VH.A(76913), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 2018;
		((_Worksheet)worksheet).get_Range((object)VH.A(76960), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 45;
		((_Worksheet)worksheet).get_Range((object)VH.A(61576), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 33;
		((_Worksheet)worksheet).get_Range((object)VH.A(61592), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 26;
		((_Worksheet)worksheet).get_Range((object)VH.A(61597), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 12;
		Range range = ((_Worksheet)worksheet).get_Range((object)VH.A(77013), RuntimeHelpers.GetObjectValue(Missing.Value));
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
		Range result = ((_Worksheet)worksheet).get_Range((object)VH.A(77024), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = null;
		return result;
	}
}
