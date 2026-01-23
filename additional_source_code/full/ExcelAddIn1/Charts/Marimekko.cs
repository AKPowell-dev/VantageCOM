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

public sealed class Marimekko
{
	private enum YD
	{
		A,
		B
	}

	private struct ZD
	{
		public float A;

		public float B;

		public bool A;

		public YD A;

		public float C;

		public int A;

		public int B;

		public XlDataLabelPosition A;

		public bool B;

		public bool C;
	}

	private struct AE
	{
		public Range A;

		public Range B;

		public Range C;

		public Range D;

		public Range E;

		public Range F;

		public Range G;

		public Range H;

		public Range I;

		public Range J;

		public Range K;

		public Range L;

		public Range M;

		public Range N;

		public Range O;

		public Range P;

		public Range Q;

		public Range R;
	}

	private static readonly string m_A = VH.A(77035);

	private static readonly int m_A = 15;

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool B = true;
		int num = 4;
		checked
		{
			if (application.Selection is Range)
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
				if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
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
					Range A = (Range)application.Selection;
					ZD a = Marimekko.A(ref A, ref B);
					if (B)
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
						if (B)
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
							int num2 = QuickCharts2.InputColor();
							int g = QuickCharts2.LinkColor();
							XlCalculation calc = default(XlCalculation);
							QuickCharts2.PrepareExcel(application, ref calc);
							Worksheet worksheet;
							ChartObject chartObject;
							Chart chart;
							Range D = default(Range);
							Range range8;
							Range range2;
							try
							{
								worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(A.Worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								chartObject = QuickCharts2.AddChart(worksheet, a.A, a.B);
								chartObject.Placement = XlPlacement.xlFreeFloating;
								int val = chartObject.BottomRightCell.Row + 1;
								chart = chartObject.Chart;
								chart.ChartType = XlChartType.xlAreaStacked100;
								QuickCharts2.RequireAxes(chart);
								AE aE = Marimekko.A(worksheet, chartObject);
								val = Math.Max(val, Marimekko.m_A + 2);
								Marimekko.A(a, worksheet, A, ref D, val, num2, g);
								int num3 = Conversions.ToInteger(D.Rows.CountLarge);
								int num4 = Conversions.ToInteger(D.Columns.CountLarge);
								Range range = D;
								range2 = ((Range)range.Columns[RuntimeHelpers.GetObjectValue(range.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1);
								int num5 = Conversions.ToInteger(range.Rows.CountLarge);
								for (int i = 1; i <= num5; i++)
								{
									((Range)range2.Cells[i, 1]).Formula = VH.A(54448) + ((Range)range.Cells[1, 1]).get_Address((object)1, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2826) + ((Range)range.Cells[i, RuntimeHelpers.GetObjectValue(range.Columns.CountLarge)]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54569) + range.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75410);
									((Range)range2.Cells[i, 1]).NumberFormat = VH.A(75421);
								}
								IEnumerator enumerator = default(IEnumerator);
								IEnumerator enumerator2 = default(IEnumerator);
								IEnumerator enumerator3 = default(IEnumerator);
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									range = null;
									int num6 = 1;
									Range range3 = JH.A(D, application);
									if (range3 != null)
									{
										{
											enumerator = range3.Columns.GetEnumerator();
											try
											{
												while (enumerator.MoveNext())
												{
													Range range4 = (Range)enumerator.Current;
													((Range)worksheet.Cells[val + num3 + 2, range4.Column]).Formula = VH.A(48936) + ((Range)worksheet.Cells[val, range4.Column]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
													((Range)worksheet.Cells[val + num3 + 2, range4.Column + num4]).Value2 = VH.A(75430) + num6;
													((Range)worksheet.Cells[val + num3 + 2, range4.Column + num4 * 2 + 1]).Value2 = VH.A(75435) + num6;
													int num7 = num3 - 1;
													for (int j = 0; j <= num7; j++)
													{
														Range obj = (Range)worksheet.Cells[val + num3 + j * num + 3, range4.Column];
														obj.Value2 = 0;
														obj.Font.Color = num2;
														_ = null;
														string formula = VH.A(48936) + ((Range)D.Cells[j + 1, range4.Column - 1]).get_Address((object)1, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
														((Range)worksheet.Cells[val + num3 + j * num + 4, range4.Column]).Formula = formula;
														((Range)worksheet.Cells[val + num3 + j * num + 5, range4.Column]).Formula = formula;
														((Range)worksheet.Cells[val + num3 + j * num + 6, range4.Column]).Formula = formula;
													}
													while (true)
													{
														switch (1)
														{
														case 0:
															continue;
														}
														break;
													}
													Range obj2 = (Range)worksheet.Cells[val + num3 + num3 * num + num - 1, range4.Column];
													obj2.Value2 = 0;
													obj2.Font.Color = num2;
													_ = null;
													num6++;
													if (!a.C)
													{
														continue;
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
													{
														enumerator2 = range4.Cells.GetEnumerator();
														try
														{
															while (true)
															{
																if (enumerator2.MoveNext())
																{
																	Range range5 = (Range)enumerator2.Current;
																	if (!Operators.ConditionalCompareObjectNotEqual(range5.Formula, string.Empty, TextCompare: false))
																	{
																		continue;
																	}
																	while (true)
																	{
																		switch (6)
																		{
																		case 0:
																			continue;
																		}
																		Marimekko.A(range5, ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[val + num3 + 3, range4.Column]), RuntimeHelpers.GetObjectValue(worksheet.Cells[val + num3 + num3 * num + num - 1, range4.Column])));
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
																		goto end_IL_07e8;
																	}
																	continue;
																	end_IL_07e8:
																	break;
																}
																break;
															}
														}
														finally
														{
															IDisposable disposable2 = enumerator2 as IDisposable;
															if (disposable2 != null)
															{
																disposable2.Dispose();
															}
														}
													}
												}
												while (true)
												{
													switch (4)
													{
													case 0:
														break;
													default:
														goto end_IL_0818;
													}
													continue;
													end_IL_0818:
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
										range3 = null;
									}
									((Range)worksheet.Cells[val + num3 + 2, num4 * 2 + 2]).Value2 = VH.A(75448);
									((Range)worksheet.Cells[val + num3 + 2, num4 * 3 + 3]).Value2 = VH.A(75469);
									Range obj3 = (Range)worksheet.Cells[val + num3 + 3, 1];
									obj3.Value2 = 0;
									obj3.NumberFormat = VH.A(75421);
									obj3.Font.Color = num2;
									_ = null;
									num6 = val + num3 + 4;
									int num8 = num3 - 1;
									for (int k = 0; k <= num8; k++)
									{
										int num9 = num - 1;
										for (int l = 0; l <= num9; l++)
										{
											string formula;
											switch (l)
											{
											case 1:
											{
												string text = VH.A(54414) + ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, 2]), RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, num4 + 1])).get_Address((object)0, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
												string text2 = VH.A(54414) + D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
												range3 = JH.A(D, application);
												if (range3 != null)
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
													try
													{
														enumerator3 = range3.Columns.GetEnumerator();
														while (enumerator3.MoveNext())
														{
															Range range6 = (Range)enumerator3.Current;
															string text3 = ((Range)worksheet.Cells[num6, range6.Column]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
															string text4 = text3 + VH.A(75498) + text;
															formula = VH.A(75501);
															formula = formula + VH.A(75516) + aE.I.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text4 + VH.A(75525) + aE.L.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
															formula = formula + VH.A(75528) + aE.J.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text + VH.A(75498) + text2 + VH.A(75525) + aE.M.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
															formula = formula + VH.A(75528) + aE.K.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text3 + VH.A(75525) + aE.N.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
															formula += VH.A(75539);
															formula = formula + VH.A(75550) + aE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + ((Range)D.Cells[1, range6.Column - 1]).get_Offset((object)(-1), (object)0).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75557);
															formula = formula + VH.A(75566) + aE.B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75575) + aE.A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75584) + ((Range)D.Cells[k + 1, 1]).get_Offset((object)0, (object)(-1)).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75557);
															formula = formula + VH.A(75566) + aE.C.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75613) + ((_Worksheet)worksheet).get_Range((object)aE.A, (object)aE.B).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75628) + text3 + VH.A(2378) + aE.Q.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
															formula = formula + VH.A(75566) + aE.D.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75613) + ((_Worksheet)worksheet).get_Range((object)aE.A, (object)aE.C).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75628) + text4 + VH.A(2378) + aE.O.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
															formula = formula + VH.A(75566) + aE.E.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75613) + ((_Worksheet)worksheet).get_Range((object)aE.A, (object)aE.D).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75628) + text3 + VH.A(75498) + text2 + VH.A(2378) + aE.P.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
															formula += VH.A(39904);
															((Range)worksheet.Cells[num6, range6.Column + num4 * 2 + 1]).Formula = formula;
															Range range7 = (Range)worksheet.Cells[num6, range6.Column + num4];
															range7.Formula = VH.A(54556) + ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, 2]), RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, range6.Column])).get_Address((object)0, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(68994) + text3 + VH.A(75680) + ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, 2]), RuntimeHelpers.GetObjectValue(worksheet.Cells[num6, num4 + 1])).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
															range7.NumberFormat = VH.A(75697);
															range7 = null;
														}
														while (true)
														{
															switch (1)
															{
															case 0:
																break;
															default:
																goto end_IL_146b;
															}
															continue;
															end_IL_146b:
															break;
														}
													}
													finally
													{
														if (enumerator3 is IDisposable)
														{
															while (true)
															{
																switch (1)
																{
																case 0:
																	continue;
																}
																(enumerator3 as IDisposable).Dispose();
																break;
															}
														}
													}
													range3 = null;
												}
												formula = VH.A(60076) + aE.J.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text + VH.A(75498) + text2 + VH.A(75525) + aE.M.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75539);
												formula = formula + VH.A(75550) + aE.F.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + ((Range)worksheet.Cells[val + k + 1, 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75557);
												formula = formula + VH.A(75566) + aE.G.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75575) + aE.F.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75702) + text + VH.A(2378) + aE.R.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
												formula = formula + VH.A(75566) + aE.H.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75613) + ((_Worksheet)worksheet).get_Range((object)aE.F, (object)aE.G).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75628) + text + VH.A(75498) + text2 + VH.A(2378) + aE.P.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75669);
												formula += VH.A(39904);
												((Range)worksheet.Cells[num6, num4 * 3 + 3]).Formula = formula;
												Range obj4 = (Range)worksheet.Cells[num6, num4 * 2 + 2];
												obj4.Value2 = ((a.A == XlDataLabelPosition.xlLabelPositionAbove) ? 1 : 0);
												obj4.NumberFormat = VH.A(75697);
												obj4.Font.Color = num2;
												_ = null;
												formula = VH.A(75741) + ((Range)worksheet.Cells[num6 - 1, 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + ((Range)worksheet.Cells[num6 + 1, 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
												break;
											}
											case 2:
												formula = VH.A(48936) + ((Range)range2.Cells[k + 1, 1]).get_Address((object)1, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
												break;
											default:
												formula = VH.A(48936) + ((Range)worksheet.Cells[num6 - 1, 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
												break;
											}
											Range obj5 = (Range)worksheet.Cells[num6, 1];
											obj5.Formula = formula;
											obj5.NumberFormat = VH.A(75421);
											_ = null;
											num6++;
										}
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										range8 = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[val + num3 + 2, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[val + num3 + 2 + num3 * 4 + 1, num4 * 2 + 2]));
										D = range8;
										if (Operators.ConditionalCompareObjectGreater(D.Columns.CountLarge, 256, TextCompare: false))
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
											D = D.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)256);
											if (num4 + 1 > 256)
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
												Forms.WarningMessage(VH.A(75760));
											}
											else
											{
												Forms.WarningMessage(VH.A(75972));
											}
										}
										chart.SetSourceData(D, RuntimeHelpers.GetObjectValue(Missing.Value));
										chart.PlotBy = XlRowCol.xlColumns;
										num6 = 1;
										int count = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
										int num10 = num4 + 1;
										int num11 = count;
										for (int m = num10; m <= num11; m++)
										{
											Series obj6 = (Series)chart.SeriesCollection(m);
											obj6.ChartType = XlChartType.xlLine;
											Series series = obj6;
											if (series.AxisGroup == XlAxisGroup.xlSecondary)
											{
												series.AxisGroup = XlAxisGroup.xlPrimary;
											}
											series.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
											series.ChartType = XlChartType.xlLine;
											series.Format.Line.Visible = MsoTriState.msoFalse;
											series.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
											series.HasLeaderLines = false;
											DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)obj6.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)range8.Columns[RuntimeHelpers.GetObjectValue(range8.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)num6).get_Resize(Operators.SubtractObject(range8.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value)));
											series = null;
											Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)obj6.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
											dataLabels.ShowLegendKey = false;
											if (m < count)
											{
												dataLabels.Position = XlDataLabelPosition.xlLabelPositionCenter;
												dataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = a.B;
											}
											else
											{
												dataLabels.Position = a.A;
											}
											dataLabels = null;
											num6++;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											if (chart.HasLegend)
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
												int num12 = num4 + 1;
												for (int n = count; n >= num12; n += -1)
												{
													((LegendEntry)chart.Legend.LegendEntries(n)).Delete();
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
											}
											Axis obj7 = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
											obj7.CategoryType = XlCategoryType.xlTimeScale;
											obj7.AxisBetweenCategories = false;
											obj7.MinimumScale = 0.0;
											obj7.MaximumScale = 100.0;
											obj7.MajorUnit = 10.0;
											obj7.MajorUnitScale = XlTimeUnit.xlDays;
											obj7.TickLabels.NumberFormat = VH.A(75421);
											_ = null;
											if (!(a.C > 0f))
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
												if (!a.B)
												{
													goto IL_2002;
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
											int num13 = num4;
											for (int num14 = 1; num14 <= num13; num14++)
											{
												ChartFormat format = ((Series)chart.SeriesCollection(num14)).Format;
												if (a.B)
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
													format.Fill.ForeColor.RGB = clsColors.RGB2Ole(KH.A.ChartSeriesColors[num14]);
												}
												if (a.C > 0f)
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
													LineFormat line = format.Line;
													line.Visible = MsoTriState.msoTrue;
													line.ForeColor.RGB = a.A;
													line.Weight = a.C;
													_ = null;
												}
												format = null;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												break;
											}
											goto IL_2002;
											IL_2002:
											Axis axis = (Axis)chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
											if (a.A == XlDataLabelPosition.xlLabelPositionBelow)
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
												axis.TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone;
												axis.MajorTickMark = XlTickMark.xlTickMarkNone;
												axis.MinorTickMark = XlTickMark.xlTickMarkNone;
												axis.HasTitle = false;
												chart.PlotArea.Height -= 10.0;
												chart.PlotArea.Width -= 10.0;
												chart.PlotArea.Top -= 2.0;
											}
											else
											{
												axis.HasTitle = true;
												axis.AxisTitle.Text = VH.A(76192);
												chart.PlotArea.Height -= 10.0;
												chart.PlotArea.Top += 10.0;
											}
											axis = null;
											Chart chart2 = chart;
											chart2.PlotArea.Left += 15.0;
											Axis obj8 = (Axis)chart2.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue);
											obj8.HasMajorGridlines = false;
											obj8.HasTitle = true;
											obj8.AxisTitle.Text = VH.A(76231);
											_ = null;
											chart2.ChartArea.Format.Line.Visible = MsoTriState.msoFalse;
											chart2.ChartArea.Select();
											_ = null;
											break;
										}
										break;
									}
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
							worksheet = null;
							chartObject = null;
							chart = null;
							D = null;
							range8 = null;
							range2 = null;
							QuickCharts2.LogActivity(VH.A(76270));
						}
					}
					A = null;
				}
			}
			application = null;
		}
	}

	private static ZD A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		wpfMarimekko wpfMarimekko2 = new wpfMarimekko();
		QuickCharts2.HandleColorChange(wpfMarimekko2.btnColorLabel);
		QuickCharts2.HandleColorChange(wpfMarimekko2.btnColorGap);
		wpfMarimekko2.Range = A;
		wpfMarimekko2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfMarimekko2.numChartWidth, wpfMarimekko2.numChartHeight);
		wpfMarimekko2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		if (Conversions.ToInteger(xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(60421)).InnerText) == 0)
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
			wpfMarimekko2.radColumns.IsChecked = true;
		}
		else
		{
			wpfMarimekko2.radRows.IsChecked = true;
		}
		QuickCharts2.SetButtonColor(wpfMarimekko2.btnColorGap, xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76301)).InnerText);
		wpfMarimekko2.numGapWidth.Value = Conversions.ToDouble(xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76328)).InnerText);
		QuickCharts2.SetButtonColor(wpfMarimekko2.btnColorLabel, xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76355)).InnerText);
		if (Conversions.ToInteger(xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76386)).InnerText) == 0)
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
			wpfMarimekko2.radAbove.IsChecked = true;
		}
		else
		{
			wpfMarimekko2.radBelow.IsChecked = true;
		}
		wpfMarimekko2.ShowDialog();
		ZD result = default(ZD);
		if (wpfMarimekko2.DialogResult.HasValue)
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
			if (wpfMarimekko2.DialogResult.Value)
			{
				A = wpfMarimekko2.Range;
				result = new ZD
				{
					A = wpfMarimekko2.chkPreserveFormulas.IsChecked.Value,
					A = (float)wpfMarimekko2.numChartWidth.Value.Value,
					B = (float)wpfMarimekko2.numChartHeight.Value.Value
				};
				if (wpfMarimekko2.radColumns.IsChecked == true)
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
					result.A = YD.A;
				}
				else
				{
					result.A = YD.B;
				}
				System.Windows.Media.Color buttonColor = QuickCharts2.GetButtonColor(wpfMarimekko2.btnColorGap);
				System.Windows.Media.Color buttonColor2 = QuickCharts2.GetButtonColor(wpfMarimekko2.btnColorLabel);
				result.B = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor2.R, buttonColor2.G, buttonColor2.B));
				result.A = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor.R, buttonColor.G, buttonColor.B));
				result.C = (float)wpfMarimekko2.numGapWidth.Value.Value;
				if (wpfMarimekko2.radAbove.IsChecked == true)
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
					result.A = XlDataLabelPosition.xlLabelPositionAbove;
				}
				else
				{
					result.A = XlDataLabelPosition.xlLabelPositionBelow;
				}
				QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfMarimekko2.numChartWidth.Value.Value));
				QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfMarimekko2.numChartHeight.Value.Value));
				QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfMarimekko2.chkPreserveFormulas.IsChecked.Value);
				XmlNode xmlNode = xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(60421));
				int a = (int)result.A;
				xmlNode.InnerText = a.ToString();
				xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76328)).InnerText = wpfMarimekko2.numGapWidth.Value.ToString();
				xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76301)).InnerText = QuickCharts2.Color2RGB(buttonColor);
				xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76355)).InnerText = QuickCharts2.Color2RGB(buttonColor2);
				XmlNode xmlNode2 = xmlDoc.SelectSingleNode(Marimekko.m_A + VH.A(76386));
				a = (int)result.A;
				xmlNode2.InnerText = a.ToString();
				KH.A.SaveSettings(xmlDoc);
				goto IL_054e;
			}
		}
		B = false;
		goto IL_054e;
		IL_054e:
		wpfMarimekko2 = null;
		xmlDoc = null;
		result.C = true;
		return result;
	}

	private static void A(ZD A, Worksheet B, Range C, ref Range D, int E, int F, int G)
	{
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.A)
			{
				if (A.A == YD.A)
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
					C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					Range range = (Range)B.Cells[E, 1];
					range.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					int num3 = num;
					for (int i = 1; i <= num3; i++)
					{
						int num4 = num2;
						Range range3;
						Range range2;
						for (int j = 1; j <= num4; range3 = null, range2 = null, j++)
						{
							range2 = (Range)C.Cells[i, j];
							range3 = (Range)B.Cells[E - 1 + i, j];
							if (Operators.ConditionalCompareObjectEqual(range2.Formula, string.Empty, TextCompare: false))
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
								range3.Clear();
								continue;
							}
							if (Conversions.ToBoolean(range2.HasFormula))
							{
								string numberFormat = Conversions.ToString(range2.NumberFormat);
								string formula = range2.Formula.ToString();
								Range range4 = ((_Worksheet)C.Worksheet).get_Range((object)VH.A(60932), RuntimeHelpers.GetObjectValue(Missing.Value));
								range4.Formula = formula;
								range4.Cut(RuntimeHelpers.GetObjectValue(B.Cells[E - 1 + i, j]));
								_ = null;
								Range obj = (Range)B.Cells[E - 1 + i, j];
								obj.Font.Color = G;
								obj.NumberFormat = numberFormat;
								_ = null;
								continue;
							}
							if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.Value2)))
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
								if (!KH.A.AutoColorText)
								{
									continue;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							range3.Font.Color = F;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_027f;
							}
							continue;
							end_IL_027f:
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
						break;
					}
					range.Clear();
					range = null;
					D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				}
				else
				{
					int num5 = num;
					Range range2;
					Range range5;
					for (int k = 1; k <= num5; k++)
					{
						int num6 = num2;
						for (int l = 1; l <= num6; l++)
						{
							range2 = (Range)C.Cells[k, l];
							range5 = (Range)B.Cells[E - 1 + l, k];
							if (Operators.ConditionalCompareObjectEqual(range2.Formula, string.Empty, TextCompare: false))
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
							if (Conversions.ToBoolean(range2.HasFormula))
							{
								string numberFormat = Conversions.ToString(range2.NumberFormat);
								string formula = range2.Formula.ToString();
								Range range6 = ((_Worksheet)C.Worksheet).get_Range((object)VH.A(60932), RuntimeHelpers.GetObjectValue(Missing.Value));
								range6.Formula = formula;
								range6.Cut(range5);
								_ = null;
								range5 = (Range)B.Cells[E - 1 + l, k];
								Range range7 = range5;
								range7.Font.Color = G;
								range7.NumberFormat = numberFormat;
								_ = null;
								continue;
							}
							if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.Value2)))
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
								if (!KH.A.AutoColorText)
								{
									range2.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
									range5.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									continue;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							range2.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
							range5.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							range5.Font.Color = F;
						}
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
					range2 = null;
					range5 = null;
					D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num2 - 1, num]));
					D.NumberFormat = RuntimeHelpers.GetObjectValue(((Range)C.Cells[2, 2]).NumberFormat);
				}
			}
			else if (A.A == YD.A)
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
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range = (Range)B.Cells[E, 1];
				range.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				range.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				range.Clear();
				range = null;
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num, num2])).Font.Color = G;
			}
			else
			{
				int num7 = num;
				for (int m = 1; m <= num7; m++)
				{
					int num8 = num2;
					for (int n = 1; n <= num8; n++)
					{
						Range range8 = (Range)C.Cells[m, n];
						if (Operators.ConditionalCompareObjectEqual(range8.Formula, string.Empty, TextCompare: false))
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
							break;
						}
						((Range)B.Cells[E - 1 + n, m]).Formula = VH.A(48936) + range8.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						range8 = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_07cf;
						}
						continue;
						end_IL_07cf:
						break;
					}
				}
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num2 - 1, num]));
				D.NumberFormat = RuntimeHelpers.GetObjectValue(((Range)C.Cells[2, 2]).NumberFormat);
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num2, num])).Font.Color = G;
			}
			if (A.C)
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
				if (A.A == YD.A)
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
					int num9 = num;
					for (int num10 = 2; num10 <= num9; num10++)
					{
						int num11 = num2;
						for (int num12 = 2; num12 <= num11; num12++)
						{
							Marimekko.A((Range)C.Cells[num10, num12], (Range)D.Cells[num10 - 1, num12 - 1]);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0961;
							}
							continue;
							end_IL_0961:
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
				else
				{
					int num13 = num;
					for (int num14 = 2; num14 <= num13; num14++)
					{
						int num15 = num2;
						for (int num16 = 2; num16 <= num15; num16++)
						{
							Marimekko.A((Range)C.Cells[num14, num16], (Range)D.Cells[num16 - 1, num14 - 1]);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_09ed;
							}
							continue;
							end_IL_09ed:
							break;
						}
					}
				}
			}
			C.Application.CutCopyMode = (XlCutCopyMode)0;
		}
	}

	private static void A(Range A, Range B)
	{
		if (!Operators.ConditionalCompareObjectNotEqual(A.Formula, string.Empty, TextCompare: false))
		{
			return;
		}
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
			Interior interior = A.Interior;
			if (Operators.ConditionalCompareObjectNotEqual(interior.ColorIndex, Constants.xlNone, TextCompare: false))
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
				B.Interior.Color = RuntimeHelpers.GetObjectValue(interior.Color);
				if ((XlPattern)interior.Pattern != XlPattern.xlPatternNone)
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
					B.Interior.Pattern = RuntimeHelpers.GetObjectValue(interior.Pattern);
					B.Interior.PatternColor = RuntimeHelpers.GetObjectValue(interior.PatternColor);
				}
			}
			interior = null;
			return;
		}
	}

	private static AE A(Worksheet A, ChartObject B)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		AE result = default(AE);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 3]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[Marimekko.m_A, column + 3])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num, num2], VH.A(76429));
			Range range2 = (Range)A.Cells[num + 1, num2];
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
			CheckBox obj2 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range3.Left, range3.Top, range3.Width, range3.Height }, null, null, null);
			obj2.Text = VH.A(76454);
			obj2.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj2.Value = Constants.xlBoth;
			range3 = null;
			result.A = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			Range range4 = range2;
			range4.Value2 = false;
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
			obj3.Text = VH.A(76477);
			obj3.LinkedCell = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj3.Value = Constants.xlOff;
			range4 = null;
			result.B = range2;
			range2 = (Range)A.Cells[num + 3, num2];
			Range range5 = range2;
			range5.Value2 = true;
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
			obj4.Text = VH.A(76504);
			obj4.LinkedCell = range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj4.Value = Constants.xlBoth;
			range5 = null;
			result.C = range2;
			range2 = (Range)A.Cells[num + 4, num2];
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
			obj5.Text = VH.A(76529);
			obj5.LinkedCell = range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj5.Value = Constants.xlOff;
			range6 = null;
			result.D = range2;
			range2 = (Range)A.Cells[num + 5, num2];
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
			obj6.Text = VH.A(76556);
			obj6.LinkedCell = range7.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj6.Value = Constants.xlOff;
			range7 = null;
			result.E = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 7, num2], VH.A(76577));
			range2 = (Range)A.Cells[num + 8, num2];
			Range range8 = range2;
			range8.Value2 = true;
			try
			{
				range8.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			CheckBox obj7 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range8.Left, range8.Top, range8.Width, range8.Height }, null, null, null);
			obj7.Text = VH.A(76477);
			obj7.LinkedCell = range8.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj7.Value = Constants.xlBoth;
			range8 = null;
			result.F = range2;
			range2 = (Range)A.Cells[num + 9, num2];
			Range range9 = range2;
			range9.Value2 = false;
			try
			{
				range9.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
			CheckBox obj8 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range9.Left, range9.Top, range9.Width, range9.Height }, null, null, null);
			obj8.Text = VH.A(76608);
			obj8.LinkedCell = range9.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj8.Value = Constants.xlOff;
			range9 = null;
			result.G = range2;
			range2 = (Range)A.Cells[num + 10, num2];
			Range range10 = range2;
			range10.Value2 = false;
			try
			{
				range10.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex15)
			{
				ProjectData.SetProjectError(ex15);
				Exception ex16 = ex15;
				ProjectData.ClearProjectError();
			}
			CheckBox obj9 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range10.Left, range10.Top, range10.Width, range10.Height }, null, null, null);
			obj9.Text = VH.A(76556);
			obj9.LinkedCell = range10.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj9.Value = Constants.xlOff;
			range10 = null;
			result.H = range2;
			num2 = column + 2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num, num2], VH.A(76637));
			range2 = (Range)A.Cells[num + 1, num2];
			Range range11 = range2;
			range11.Value2 = true;
			try
			{
				range11.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				ProjectData.ClearProjectError();
			}
			CheckBox obj10 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range11.Left, range11.Top, range11.Width, range11.Height }, null, null, null);
			obj10.Text = VH.A(76678);
			obj10.LinkedCell = range11.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj10.Value = Constants.xlBoth;
			range11 = null;
			result.I = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			Range range12 = range2;
			range12.Value2 = false;
			try
			{
				range12.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex19)
			{
				ProjectData.SetProjectError(ex19);
				Exception ex20 = ex19;
				ProjectData.ClearProjectError();
			}
			CheckBox obj11 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range12.Left, range12.Top, range12.Width, range12.Height }, null, null, null);
			obj11.Text = VH.A(76715);
			obj11.LinkedCell = range12.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj11.Value = Constants.xlOff;
			range12 = null;
			result.J = range2;
			range2 = (Range)A.Cells[num + 3, num2];
			Range range13 = range2;
			range13.Value2 = false;
			try
			{
				range13.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex21)
			{
				ProjectData.SetProjectError(ex21);
				Exception ex22 = ex21;
				ProjectData.ClearProjectError();
			}
			CheckBox obj12 = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4] { range13.Left, range13.Top, range13.Width, range13.Height }, null, null, null);
			obj12.Text = VH.A(76750);
			obj12.LinkedCell = range13.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			obj12.Value = Constants.xlOff;
			range13 = null;
			result.K = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 5, num2], VH.A(76781));
			((Range)A.Cells[num + 6, num2]).Value2 = VH.A(76529);
			((Range)A.Cells[num + 7, num2]).Value2 = VH.A(76556);
			((Range)A.Cells[num + 8, num2]).Value2 = VH.A(76504);
			((Range)A.Cells[num + 9, num2]).Value2 = VH.A(76608);
			num2 = column + 3;
			((Range)A.Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]).ColumnWidth = 15;
			range2 = (Range)A.Cells[num + 1, num2];
			try
			{
				range2.NumberFormat = VH.A(75697);
			}
			catch (Exception ex23)
			{
				ProjectData.SetProjectError(ex23);
				Exception ex24 = ex23;
				ProjectData.ClearProjectError();
			}
			range2.Value2 = 0.01;
			QuickCharts2.AddPercentageValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.L = range2;
			range2 = (Range)A.Cells[num + 2, num2];
			try
			{
				range2.NumberFormat = VH.A(75697);
			}
			catch (Exception ex25)
			{
				ProjectData.SetProjectError(ex25);
				Exception ex26 = ex25;
				ProjectData.ClearProjectError();
			}
			range2.Value2 = 0.01;
			QuickCharts2.AddPercentageValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.M = range2;
			range2 = (Range)A.Cells[num + 3, num2];
			range2.Value2 = 0;
			QuickCharts2.FormatOptionsInput(range2);
			result.N = range2;
			range2 = (Range)A.Cells[num + 6, num2];
			range2.Value2 = VH.A(76810);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.O = range2;
			range2 = (Range)A.Cells[num + 7, num2];
			range2.Value2 = VH.A(76843);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.P = range2;
			range2 = (Range)A.Cells[num + 8, num2];
			range2.Value2 = VH.A(76870);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.Q = range2;
			range2 = (Range)A.Cells[num + 9, num2];
			range2.Value2 = VH.A(76870);
			QuickCharts2.AddNumFormatValidation(range2);
			QuickCharts2.FormatOptionsInput(range2);
			result.R = range2;
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[Marimekko.m_A, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[Marimekko.m_A, column + 3])).Borders[XlBordersIndex.xlEdgeBottom];
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
		((_Worksheet)worksheet).get_Range((object)VH.A(76877), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(57617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76882);
		((_Worksheet)worksheet).get_Range((object)VH.A(76893), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76898);
		((_Worksheet)worksheet).get_Range((object)VH.A(76913), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76918);
		((_Worksheet)worksheet).get_Range((object)VH.A(76929), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76934);
		((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 77;
		((_Worksheet)worksheet).get_Range((object)VH.A(76950), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 65;
		((_Worksheet)worksheet).get_Range((object)VH.A(76955), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 60;
		((_Worksheet)worksheet).get_Range((object)VH.A(76960), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 54;
		((_Worksheet)worksheet).get_Range((object)VH.A(76965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76970);
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 61;
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 54;
		((_Worksheet)worksheet).get_Range((object)VH.A(61532), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61576), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 42;
		((_Worksheet)worksheet).get_Range((object)VH.A(76981), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(76986);
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 52;
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 47;
		((_Worksheet)worksheet).get_Range((object)VH.A(61546), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 40;
		((_Worksheet)worksheet).get_Range((object)VH.A(61592), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 35;
		((_Worksheet)worksheet).get_Range((object)VH.A(76997), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(77002);
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 39;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 30;
		((_Worksheet)worksheet).get_Range((object)VH.A(61551), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 26;
		((_Worksheet)worksheet).get_Range((object)VH.A(61597), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 21;
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
