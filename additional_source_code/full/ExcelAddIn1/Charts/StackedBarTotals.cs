using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class StackedBarTotals
{
	public static void Add()
	{
		if (!Licensing.AllowChartAddOnOperation())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Chart activeChart = application.ActiveChart;
		Range range = null;
		Range range2 = null;
		Range range3 = null;
		bool flag = false;
		checked
		{
			if (activeChart != null)
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
				List<Range> list = new List<Range>();
				List<Series> list2 = new List<Series>();
				string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
				bool autoExpandListRange = application.AutoCorrect.AutoExpandListRange;
				application.AutoCorrect.AutoExpandListRange = false;
				application.CutCopyMode = (XlCutCopyMode)0;
				application.ScreenUpdating = false;
				Range range4;
				try
				{
					IEnumerator enumerator = default(IEnumerator);
					XlChartType chartType = default(XlChartType);
					bool flag2 = default(bool);
					Series series;
					try
					{
						enumerator = ((IEnumerable)activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator.MoveNext())
						{
							series = (Series)enumerator.Current;
							if (series.AxisGroup == XlAxisGroup.xlSecondary)
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
							chartType = series.ChartType;
							if (chartType != XlChartType.xlBarStacked)
							{
								_ = 61;
							}
							list2.Add(series);
							string[] array = Helpers.A(series);
							try
							{
								range = ((_Application)application).get_Range((object)array[2], RuntimeHelpers.GetObjectValue(Missing.Value));
								list.Add(range);
								flag2 = range.Columns.Count == 1;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								range = null;
								ProjectData.ClearProjectError();
							}
							if (range2 == null)
							{
								range2 = range;
							}
							else if (range != null)
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
								range2 = application.Union(range2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							try
							{
								range4 = ((_Application)application).get_Range((object)array[0], RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								range4 = null;
								ProjectData.ClearProjectError();
							}
							if (range3 == null)
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
								range3 = range4;
							}
							else
							{
								if (range4 == null)
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
								range3 = application.Union(range3, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0470;
							}
							continue;
							end_IL_0470:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					series = null;
					bool flag3 = false;
					bool flag4 = default(bool);
					Range rng = default(Range);
					if (flag2)
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
						if (range2.Areas.Count == 1)
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
							if (Operators.ConditionalCompareObjectGreater(range2.Columns.CountLarge, 1, TextCompare: false))
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
								int num = 1;
								IEnumerator enumerator2 = default(IEnumerator);
								try
								{
									enumerator2 = list[list.Count - 1].Cells.GetEnumerator();
									while (true)
									{
										if (enumerator2.MoveNext())
										{
											Range range5 = (Range)enumerator2.Current;
											try
											{
												if (Operators.ConditionalCompareObjectNotEqual(range5.Value2, application.WorksheetFunction.Sum(RuntimeHelpers.GetObjectValue(range2.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(range2.Columns.Count - 1)).Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
												{
													while (true)
													{
														switch (5)
														{
														case 0:
															continue;
														}
														flag = false;
														break;
													}
													break;
												}
												flag = true;
											}
											catch (InvalidCastException ex5)
											{
												ProjectData.SetProjectError(ex5);
												InvalidCastException ex6 = ex5;
												flag = false;
												ProjectData.ClearProjectError();
												break;
											}
											num++;
											continue;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_0729;
											}
											continue;
											end_IL_0729:
											break;
										}
										break;
									}
								}
								finally
								{
									if (enumerator2 is IDisposable)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											(enumerator2 as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							else
							{
								flag = false;
							}
						}
						else
						{
							flag = false;
						}
						if (flag)
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
							series = list2.Last();
						}
						else
						{
							int num = 1;
							foreach (Range cell in ((Range)range2.Columns[range2.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)0, (object)1).Cells)
							{
								try
								{
									if (Operators.ConditionalCompareObjectNotEqual(cell.Value2, application.WorksheetFunction.Sum(RuntimeHelpers.GetObjectValue(range2.Rows[num, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											flag4 = false;
											break;
										}
										break;
									}
									flag4 = true;
								}
								catch (InvalidCastException ex7)
								{
									ProjectData.SetProjectError(ex7);
									InvalidCastException ex8 = ex7;
									flag4 = false;
									ProjectData.ClearProjectError();
									break;
								}
								num++;
							}
							if (flag4)
							{
								Range range7 = (Range)range2.Columns[range2.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
								if (range3 != null)
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
									range4 = ((Range)range3.Cells[1, range3.Columns.Count]).get_Offset((object)0, (object)1);
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)(range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + listSeparator + range7.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlColumns,
										true,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
								}
								else
								{
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)range7.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlColumns,
										false,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
									series.Name = VH.A(57315);
								}
								rng = range7.get_Offset((object)0, (object)1);
								range7 = null;
								flag3 = true;
								if (activeChart.HasLegend)
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
									LegendEntries obj = (LegendEntries)activeChart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
									obj.Item(obj.Count).Delete();
									_ = null;
								}
							}
							else if (MessageBox.Show(VH.A(69867), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
								Range range8 = (Range)range2.Columns[range2.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
								range8.get_Offset((object)0, (object)1).EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
								Range range9 = range8.get_Offset((object)0, (object)1);
								if (range2.Areas.Count == 1)
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
									range9.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(54448), NewLateBinding.LateGet(range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904));
								}
								else
								{
									List<string> list3 = new List<string>();
									IEnumerator enumerator4 = default(IEnumerator);
									try
									{
										enumerator4 = range2.Areas.GetEnumerator();
										while (enumerator4.MoveNext())
										{
											Range range10 = (Range)enumerator4.Current;
											list3.Add(Conversions.ToString(NewLateBinding.LateGet(range10.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)));
										}
									}
									finally
									{
										if (enumerator4 is IDisposable)
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												(enumerator4 as IDisposable).Dispose();
												break;
											}
										}
									}
									range9.Formula = VH.A(54448) + Strings.Join(list3.ToArray(), CultureInfo.CurrentCulture.TextInfo.ListSeparator) + VH.A(39904);
									list3 = null;
								}
								range9.Font.Color = ColorTranslator.ToOle(Color.Black);
								range9 = null;
								if (range3 != null)
								{
									range4 = ((Range)range3.Cells[1, range3.Columns.Count]).get_Offset((object)0, (object)1);
									range4.Value2 = VH.A(57315);
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)(range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + listSeparator + range8.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlColumns,
										true,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
								}
								else
								{
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)range8.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlColumns,
										false,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
									series.Name = VH.A(57315);
								}
								rng = range8.get_Offset((object)0, (object)1);
								flag3 = true;
								range8 = null;
								if (activeChart.HasLegend)
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
									LegendEntries obj2 = (LegendEntries)activeChart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
									obj2.Item(obj2.Count).Delete();
									_ = null;
								}
							}
						}
					}
					else
					{
						if (range2.Areas.Count == 1)
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
							if (Operators.ConditionalCompareObjectGreater(range2.Rows.CountLarge, 1, TextCompare: false))
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
								int num2 = 1;
								IEnumerator enumerator5 = default(IEnumerator);
								try
								{
									enumerator5 = list[list.Count - 1].Cells.GetEnumerator();
									while (enumerator5.MoveNext())
									{
										Range range11 = (Range)enumerator5.Current;
										try
										{
											if (Operators.ConditionalCompareObjectNotEqual(range11.Value2, application.WorksheetFunction.Sum(RuntimeHelpers.GetObjectValue(range2.get_Resize((object)(range2.Rows.Count - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														continue;
													}
													flag = false;
													break;
												}
												break;
											}
											flag = true;
										}
										catch (InvalidCastException ex9)
										{
											ProjectData.SetProjectError(ex9);
											InvalidCastException ex10 = ex9;
											flag = false;
											ProjectData.ClearProjectError();
											break;
										}
										num2++;
									}
								}
								finally
								{
									if (enumerator5 is IDisposable)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											(enumerator5 as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							else
							{
								flag = false;
							}
						}
						else
						{
							flag = false;
						}
						if (flag)
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
							series = list2.Last();
						}
						else
						{
							int num2 = 1;
							foreach (Range cell2 in ((Range)range2.Rows[range2.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)1, (object)0).Cells)
							{
								try
								{
									if (Operators.ConditionalCompareObjectNotEqual(cell2.Value2, application.WorksheetFunction.Sum(RuntimeHelpers.GetObjectValue(range2.Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											flag4 = false;
											break;
										}
										break;
									}
									flag4 = true;
								}
								catch (InvalidCastException ex11)
								{
									ProjectData.SetProjectError(ex11);
									InvalidCastException ex12 = ex11;
									flag4 = false;
									ProjectData.ClearProjectError();
									break;
								}
								num2++;
							}
							if (flag4)
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
								Range range13 = (Range)range2.Rows[range2.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
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
									range4 = ((Range)range3.Cells[range3.Rows.Count, 1]).get_Offset((object)1, (object)0);
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)(range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + listSeparator + range13.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlRows,
										true,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
								}
								else
								{
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)range13.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlRows,
										false,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
									series.Name = VH.A(57315);
								}
								rng = range13.get_Offset((object)1, (object)0);
								range13 = null;
								flag3 = true;
								if (activeChart.HasLegend)
								{
									LegendEntries obj3 = (LegendEntries)activeChart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
									obj3.Item(obj3.Count).Delete();
									_ = null;
								}
							}
							else if (MessageBox.Show(VH.A(69867), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
								Range range14 = (Range)range2.Rows[range2.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
								range14.get_Offset((object)1, (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
								Range range15 = range14.get_Offset((object)1, (object)0);
								if (range2.Areas.Count == 1)
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
									range15.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(54448), NewLateBinding.LateGet(range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(39904));
								}
								else
								{
									List<string> list4 = new List<string>();
									IEnumerator enumerator7 = default(IEnumerator);
									try
									{
										enumerator7 = range2.Areas.GetEnumerator();
										while (enumerator7.MoveNext())
										{
											Range range16 = (Range)enumerator7.Current;
											list4.Add(Conversions.ToString(NewLateBinding.LateGet(range16.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)));
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_1bfa;
											}
											continue;
											end_IL_1bfa:
											break;
										}
									}
									finally
									{
										if (enumerator7 is IDisposable)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												(enumerator7 as IDisposable).Dispose();
												break;
											}
										}
									}
									range15.Formula = VH.A(54448) + Strings.Join(list4.ToArray(), CultureInfo.CurrentCulture.TextInfo.ListSeparator) + VH.A(39904);
									list4 = null;
								}
								range15.Font.Color = ColorTranslator.ToOle(Color.Black);
								range15 = null;
								if (range3 != null)
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
									range4 = ((Range)range3.Cells[range3.Rows.Count, 1]).get_Offset((object)1, (object)0);
									range4.Value2 = VH.A(57315);
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)(range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + listSeparator + range14.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlRows,
										true,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
								}
								else
								{
									series = (Series)NewLateBinding.LateGet(activeChart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
									{
										((_Application)application).get_Range((object)range14.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
										XlRowCol.xlRows,
										false,
										false
									}, new string[4]
									{
										VH.A(69016),
										VH.A(69029),
										VH.A(69042),
										VH.A(69067)
									}, null, null);
									series.Name = VH.A(57315);
								}
								rng = range14.get_Offset((object)1, (object)0);
								flag3 = true;
								range14 = null;
								if (activeChart.HasLegend)
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
									LegendEntries obj4 = (LegendEntries)activeChart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
									obj4.Item(obj4.Count).Delete();
									_ = null;
								}
							}
						}
					}
					if (series != null)
					{
						IEnumerator enumerator8 = default(IEnumerator);
						IEnumerator enumerator9 = default(IEnumerator);
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							Series series2 = series;
							if (0 == 0)
							{
								if (series2.ChartType != chartType || series2.AxisGroup == XlAxisGroup.xlSecondary)
								{
									int gapWidth = ((ChartGroup)activeChart.ChartGroups(1)).GapWidth;
									series2.ChartType = chartType;
									series2.AxisGroup = XlAxisGroup.xlPrimary;
									((ChartGroup)activeChart.ChartGroups(1)).GapWidth = gapWidth;
								}
								series2.Format.Fill.Visible = MsoTriState.msoFalse;
								if (series2.HasDataLabels)
								{
									try
									{
										enumerator8 = ((IEnumerable)series2.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
										while (enumerator8.MoveNext())
										{
											Microsoft.Office.Interop.Excel.Point point = (Microsoft.Office.Interop.Excel.Point)enumerator8.Current;
											if (!point.HasDataLabel)
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
												point.HasDataLabel = true;
												point.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
											}
											point.DataLabel.Position = XlDataLabelPosition.xlLabelPositionInsideBase;
											point = null;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_20ef;
											}
											continue;
											end_IL_20ef:
											break;
										}
									}
									finally
									{
										if (enumerator8 is IDisposable)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												(enumerator8 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
								else
								{
									series2.HasDataLabels = true;
									series2.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									NewLateBinding.LateSetComplex(series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60578), new object[1] { XlDataLabelPosition.xlLabelPositionInsideBase }, null, null, OptimisticSet: false, RValueBase: true);
									NewLateBinding.LateSetComplex(series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(70035), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
								}
								List<int> list5 = new List<int>();
								try
								{
									enumerator9 = ((IEnumerable)series2.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator9.MoveNext())
									{
										_ = (Microsoft.Office.Interop.Excel.Point)enumerator9.Current;
										list5.Add(0);
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_2257;
										}
										continue;
										end_IL_2257:
										break;
									}
								}
								finally
								{
									if (enumerator9 is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator9 as IDisposable).Dispose();
											break;
										}
									}
								}
								series.Values = list5.ToArray();
								list5 = null;
								DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), rng);
							}
							else
							{
								if (series2.AxisGroup == XlAxisGroup.xlSecondary)
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
									series2.AxisGroup = XlAxisGroup.xlPrimary;
								}
								series2.ChartType = XlChartType.xlLine;
								series2.Format.Line.Visible = MsoTriState.msoFalse;
								series2.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
								if (series2.HasDataLabels)
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
									try
									{
										IEnumerator enumerator10 = ((IEnumerable)series2.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
										while (enumerator10.MoveNext())
										{
											Microsoft.Office.Interop.Excel.Point point2 = (Microsoft.Office.Interop.Excel.Point)enumerator10.Current;
											if (!point2.HasDataLabel)
											{
												point2.HasDataLabel = true;
												point2.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
											}
											if (flag3)
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
												point2.DataLabel.Position = XlDataLabelPosition.xlLabelPositionRight;
											}
											point2 = null;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_23fc;
											}
											continue;
											end_IL_23fc:
											break;
										}
									}
									finally
									{
										IEnumerator enumerator10;
										if (enumerator10 is IDisposable)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												(enumerator10 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
								else
								{
									series2.HasDataLabels = true;
									series2.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									NewLateBinding.LateSetComplex(series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60578), new object[1] { XlDataLabelPosition.xlLabelPositionRight }, null, null, OptimisticSet: false, RValueBase: true);
									NewLateBinding.LateSetComplex(series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(70035), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
									series2.HasLeaderLines = false;
								}
							}
							NewLateBinding.LateCall(series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
							series2 = null;
							series = null;
							break;
						}
					}
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					Forms.ErrorMessage(ex14.Message);
					clsReporting.LogException(ex14);
					ProjectData.ClearProjectError();
				}
				range = null;
				range2 = null;
				range4 = null;
				range3 = null;
				list = null;
				list2 = null;
				application.AutoCorrect.AutoExpandListRange = autoExpandListRange;
				application.ScreenUpdating = true;
			}
			application = null;
			activeChart = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(70072));
		}
	}
}
