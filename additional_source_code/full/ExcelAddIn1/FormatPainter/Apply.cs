using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Charts;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FormatPainter;

public sealed class Apply
{
	public static void ToSelection(Properties props, Options options)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		List<Chart> list = null;
		ChartObject chartObject = null;
		try
		{
			list = A(application);
			if (list.Any())
			{
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
					application.ScreenUpdating = false;
					try
					{
						using (List<Chart>.Enumerator enumerator = list.GetEnumerator())
						{
							while (enumerator.MoveNext())
							{
								Chart current = enumerator.Current;
								Chart chart = current;
								if (!options.Chart.Format && !options.PlotArea.Format)
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
									if (!options.Series.Format)
									{
										goto IL_05d2;
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
								}
								chartObject = (ChartObject)((Shape)((ChartObject)chart.Parent).Duplicate()).Chart.Parent;
								if (options.Chart.Format)
								{
									chart.ChartArea.ClearFormats();
								}
								if (options.PlotArea.Format)
								{
									chart.PlotArea.ClearFormats();
								}
								checked
								{
									if (options.Series.Format)
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
										FullSeriesCollection fullSeriesCollection = (FullSeriesCollection)chart.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
										int count = fullSeriesCollection.Count;
										for (int i = 1; i <= count; i++)
										{
											fullSeriesCollection.Item(i).ClearFormats();
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
										fullSeriesCollection = null;
									}
									chart.ChartStyle = props.Chart.ChartStyle;
									if (!options.Series.GapWidthOverlap)
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
										ChartGroups chartGroups = (ChartGroups)chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
										int count2 = chartGroups.Count;
										for (int j = 1; j <= count2; j++)
										{
											try
											{
												chartGroups.Item(j).GapWidth = ((ChartGroups)chartObject.Chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(j).GapWidth;
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
											try
											{
												chartGroups.Item(j).Overlap = ((ChartGroups)chartObject.Chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(j).Overlap;
											}
											catch (Exception ex3)
											{
												ProjectData.SetProjectError(ex3);
												Exception ex4 = ex3;
												ProjectData.ClearProjectError();
											}
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
										chartGroups = null;
									}
									if (chart.HasLegend)
									{
										Legend legend = chartObject.Chart.Legend;
										Legend legend2 = chart.Legend;
										if (!options.Legend.Position)
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
											try
											{
												legend2.IncludeInLayout = legend.IncludeInLayout;
												if (legend.Position != XlLegendPosition.xlLegendPositionCustom)
												{
													legend2.Position = legend.Position;
												}
												else
												{
													if (legend.Left > current.ChartArea.Width * 0.5)
													{
														legend2.Position = XlLegendPosition.xlLegendPositionRight;
													}
													else if (legend.Top > current.ChartArea.Height * 0.5)
													{
														legend2.Position = XlLegendPosition.xlLegendPositionBottom;
													}
													else if (legend.Left < current.PlotArea.Left)
													{
														legend2.Position = XlLegendPosition.xlLegendPositionLeft;
													}
													else
													{
														legend2.Position = XlLegendPosition.xlLegendPositionTop;
													}
													legend2.Top = legend.Top;
													legend2.Left = legend.Left;
													legend2.Width = legend.Width;
													legend2.Height = legend.Height;
												}
											}
											catch (Exception ex5)
											{
												ProjectData.SetProjectError(ex5);
												Exception ex6 = ex5;
												ProjectData.ClearProjectError();
											}
										}
										if (!options.Legend.Format)
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
											try
											{
												A(legend2.Font, Properties.GetFontProperties(legend.Font));
												A(legend2.Format.Fill, Properties.GetFillProperties(legend.Format.Fill));
												A(legend2.Format.Line, Properties.GetBorderProperties(legend.Format.Line));
											}
											catch (Exception ex7)
											{
												ProjectData.SetProjectError(ex7);
												Exception ex8 = ex7;
												ProjectData.ClearProjectError();
											}
										}
										legend2 = null;
										legend = null;
									}
									if (chart.HasTitle)
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
										ChartTitle chartTitle = chartObject.Chart.ChartTitle;
										ChartTitle chartTitle2 = chart.ChartTitle;
										if (!options.Title.Position)
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
											try
											{
												chartTitle2.Position = chartTitle.Position;
											}
											catch (Exception ex9)
											{
												ProjectData.SetProjectError(ex9);
												Exception ex10 = ex9;
												ProjectData.ClearProjectError();
											}
										}
										if (!options.Title.Format)
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
												A(chartTitle2.Font, Properties.GetFontProperties(chartTitle.Font));
												A(chartTitle2.Format.Fill, Properties.GetFillProperties(chartTitle.Format.Fill));
												A(chartTitle2.Format.Line, Properties.GetBorderProperties(chartTitle.Format.Line));
											}
											catch (Exception ex11)
											{
												ProjectData.SetProjectError(ex11);
												Exception ex12 = ex11;
												ProjectData.ClearProjectError();
											}
										}
										chartTitle2 = null;
										chartTitle = null;
									}
									try
									{
										chartObject.Delete();
									}
									catch (Exception ex13)
									{
										ProjectData.SetProjectError(ex13);
										Exception ex14 = ex13;
										ProjectData.ClearProjectError();
									}
									goto IL_05d2;
								}
								IL_08b2:
								if (options.Title.Format || options.Title.Position)
								{
									chart.HasTitle = props.Chart.HasTitle;
									if (props.Chart.HasTitle)
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
										A(options, chart.ChartTitle, props.Title);
									}
								}
								if (options.DataTable.Format)
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
									chart.HasDataTable = props.Chart.HasDataTable;
									if (props.Chart.HasDataTable)
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
										A(options, chart.DataTable, props.DataTable);
									}
								}
								A(options, (ChartObject)chart.Parent, props.ChartObject);
								A(options, chart.ChartArea, props.ChartArea);
								A(options, chart.PlotArea, props.PlotArea);
								chart = null;
								continue;
								IL_05d2:
								A(options, current, props.Series);
								int num;
								if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlPrimary)))
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
									num = (props.Chart.HasPrimaryValueAxis ? 1 : 0);
								}
								else
								{
									num = 0;
								}
								if (Conversions.ToBoolean((byte)num != 0))
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
									Options.myPrimaryValueAxis primaryValueAxis = options.PrimaryValueAxis;
									A((Axis)current.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue), props.PrimaryValueAxis, primaryValueAxis.Scale, primaryValueAxis.Gridlines, primaryValueAxis.Ticks, primaryValueAxis.Title);
									primaryValueAxis = null;
								}
								int num2;
								if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlPrimary)))
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
									num2 = (props.Chart.HasPrimaryCategoryAxis ? 1 : 0);
								}
								else
								{
									num2 = 0;
								}
								if (Conversions.ToBoolean((byte)num2 != 0))
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
									Options.myPrimaryCategoryAxis primaryCategoryAxis = options.PrimaryCategoryAxis;
									A((Axis)current.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory), props.PrimaryCategoryAxis, primaryCategoryAxis.Scale, primaryCategoryAxis.Gridlines, primaryCategoryAxis.Ticks, primaryCategoryAxis.Title);
									primaryCategoryAxis = null;
								}
								int num3;
								if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary)))
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
									num3 = (props.Chart.HasSecondaryValueAxis ? 1 : 0);
								}
								else
								{
									num3 = 0;
								}
								if (Conversions.ToBoolean((byte)num3 != 0))
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
									Options.mySecondaryValueAxis secondaryValueAxis = options.SecondaryValueAxis;
									A((Axis)current.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary), props.SecondaryValueAxis, secondaryValueAxis.Scale, secondaryValueAxis.Gridlines, secondaryValueAxis.Ticks, secondaryValueAxis.Title);
									secondaryValueAxis = null;
								}
								int num4;
								if (Conversions.ToBoolean(((_Chart)chart).get_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary)))
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
									num4 = (props.Chart.HasSecondaryCategoryAxis ? 1 : 0);
								}
								else
								{
									num4 = 0;
								}
								if (Conversions.ToBoolean((byte)num4 != 0))
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
									Options.mySecondaryCategoryAxis secondaryCategoryAxis = options.SecondaryCategoryAxis;
									A((Axis)current.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, XlAxisGroup.xlSecondary), props.SecondaryCategoryAxis, secondaryCategoryAxis.Scale, secondaryCategoryAxis.Gridlines, secondaryCategoryAxis.Ticks, secondaryCategoryAxis.Title);
									secondaryCategoryAxis = null;
								}
								if (!options.Legend.Format)
								{
									if (!options.Legend.Position)
									{
										goto IL_08b2;
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
								}
								chart.HasLegend = props.Chart.HasLegend;
								if (props.Chart.HasLegend)
								{
									A(options, chart.Legend, props.Legend);
								}
								goto IL_08b2;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_09c8;
								}
								continue;
								end_IL_09c8:
								break;
							}
						}
						clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, VH.A(172058));
					}
					catch (Exception ex15)
					{
						ProjectData.SetProjectError(ex15);
						Exception ex16 = ex15;
						MessageBox.Show(ex16.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
						clsReporting.LogException(ex16);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
					break;
				}
			}
			else
			{
				A();
			}
		}
		catch (Exception ex17)
		{
			ProjectData.SetProjectError(ex17);
			Exception ex18 = ex17;
			A();
			ProjectData.ClearProjectError();
		}
		application = null;
		JH.A((object)list);
	}

	private static List<Chart> A(Microsoft.Office.Interop.Excel.Application A)
	{
		List<Chart> list = new List<Chart>();
		Chart chart = Helpers.SelectedChart();
		if (chart != null)
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
			list.Add(chart);
			chart = null;
		}
		else if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(A.Selection)), VH.A(56245), TextCompare: false) == 0)
		{
			try
			{
				IEnumerator enumerator = ((IEnumerable)NewLateBinding.LateGet(A.Selection, null, VH.A(56274), new object[0], null, null, null)).GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Shape shape = (Shape)enumerator.Current;
						if (shape.HasChart != MsoTriState.msoTrue)
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
						list.Add(shape.Chart);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_00c8;
						}
						continue;
						end_IL_00c8:
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
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		return list;
	}

	private static void A(Options A, ChartObject B, Properties.ChartObjectProperties C)
	{
		if (A.Chart.Size)
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
			B.ShapeRange.LockAspectRatio = MsoTriState.msoFalse;
			B.Height = C.Height;
			B.Width = C.Width;
			B.ShapeRange.LockAspectRatio = C.LockAspectRatio;
		}
		if (A.Chart.Top)
		{
			B.Top = C.Top;
		}
		if (!A.Chart.Left)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			B.Left = C.Left;
			return;
		}
	}

	private static void A(Options A, ChartArea B, Properties.ChartAreaProperties C)
	{
		if (!A.Chart.Format)
		{
			return;
		}
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
			Apply.A(B.Format.Fill, C.Fill);
			Apply.A(B.Format.Line, C.Border);
			return;
		}
	}

	private static void A(Options A, PlotArea B, Properties.PlotAreaProperties C)
	{
		if (A.PlotArea.Size)
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
			if (!A.PlotArea.Location)
			{
				Apply.A(B, C);
				goto IL_00b7;
			}
		}
		if (A.PlotArea.Location)
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
			if (!A.PlotArea.Size)
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
				Apply.B(B, C);
				goto IL_00b7;
			}
		}
		if (A.PlotArea.Size && A.PlotArea.Location)
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
			Apply.B(B, C);
			Apply.A(B, C);
			Apply.B(B, C);
		}
		goto IL_00b7;
		IL_00b7:
		if (!A.PlotArea.Format)
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
			Apply.A(B.Format.Fill, C.Fill);
			Apply.A(B.Format.Line, C.Border);
			return;
		}
	}

	private static void A(PlotArea A, Properties.PlotAreaProperties B)
	{
		A.InsideHeight = B.InsideHeight;
		A.InsideWidth = B.InsideWidth;
	}

	private static void B(PlotArea A, Properties.PlotAreaProperties B)
	{
		A.InsideTop = B.InsideTop;
		A.InsideLeft = B.InsideLeft;
	}

	private static void A(Options A, Chart B, Properties.SeriesProperties C)
	{
		checked
		{
			if (A.Series.Format)
			{
				FullSeriesCollection fullSeriesCollection = (FullSeriesCollection)B.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count = fullSeriesCollection.Count;
				int num = 1;
				while (true)
				{
					if (num <= count)
					{
						if (!C.ChartType.ContainsKey(num))
						{
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							if (fullSeriesCollection.Item(num).ChartType == C.ChartType[num])
							{
								Series series = fullSeriesCollection.Item(num);
								Apply.A(series.Format.Fill, C.Fill[num]);
								Apply.A(series.Format.Line, C.Border[num]);
								if (C.Markers.ContainsKey(num))
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
									try
									{
										series.MarkerStyle = C.Markers[num].MarkerStyle;
										if (series.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												series.MarkerSize = C.Markers[num].MarkerSize;
												series.MarkerForegroundColor = C.Markers[num].MarkerForegroundColor;
												series.MarkerBackgroundColor = C.Markers[num].MarkerBackgroundColor;
												break;
											}
										}
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
								}
								series = null;
							}
							num++;
							break;
						}
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
					break;
				}
				fullSeriesCollection = null;
			}
			if (A.Series.DataLabels)
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
				FullSeriesCollection fullSeriesCollection2 = (FullSeriesCollection)B.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count2 = fullSeriesCollection2.Count;
				int num2 = 1;
				while (true)
				{
					if (num2 <= count2)
					{
						if (!C.DataLabels.ContainsKey(num2))
						{
							break;
						}
						if (C.DataLabels[num2].HasDataLabels)
						{
							Properties.DataLabelProperties dataLabelProperties = C.DataLabels[num2];
							Series series2 = fullSeriesCollection2.Item(num2);
							try
							{
								if (series2.HasDataLabels)
								{
									Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series2.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
									ChartFormat format = dataLabels.Format;
									Apply.A(format.Fill, dataLabelProperties.Fill);
									Apply.A(format.Line, dataLabelProperties.Border);
									_ = null;
									Apply.A(dataLabels.Font, dataLabelProperties.Font);
									dataLabels.AutoText = dataLabelProperties.AutoText;
									dataLabels.NumberFormatLinked = dataLabelProperties.NumberFormatLinked;
									if (!dataLabels.NumberFormatLinked)
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
										dataLabels.NumberFormat = dataLabelProperties.NumberFormat;
									}
									dataLabels.ShowLegendKey = dataLabelProperties.ShowLegendKey;
									dataLabels = null;
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							try
							{
								series2.HasLeaderLines = dataLabelProperties.HasLeaderLines;
								if (series2.HasLeaderLines)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										Type typeFromHandle = typeof(Apply);
										string memberName = VH.A(172099);
										object[] array = new object[2];
										Series instance = series2;
										string memberName2 = VH.A(140662);
										object[] obj = new object[1] { num2 };
										object[] array2 = obj;
										bool[] obj2 = new bool[1] { true };
										bool[] array3 = obj2;
										object instance2 = NewLateBinding.LateGet(instance, null, memberName2, obj, null, null, obj2);
										if (array3[0])
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
											num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(int));
										}
										object instance3;
										array[0] = NewLateBinding.LateGet(instance3 = NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(172122), new object[0], null, null, null), null, VH.A(60221), new object[0], null, null, null), null, VH.A(57400), new object[0], null, null, null);
										ref Properties.LineProperties line = ref dataLabelProperties.Line;
										array[1] = line;
										object[] array4 = array;
										bool[] array5;
										NewLateBinding.LateCall(null, typeFromHandle, memberName, array, null, null, array5 = new bool[2] { true, true }, IgnoreReturn: true);
										if (array5[0])
										{
											NewLateBinding.LateSetComplex(instance3, null, VH.A(57400), new object[1] { array4[0] }, null, null, OptimisticSet: true, RValueBase: true);
										}
										if (!array5[1])
										{
											break;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											line = (Properties.LineProperties)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[1]), typeof(Properties.LineProperties));
											break;
										}
										break;
									}
								}
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
							series2 = null;
						}
						num2++;
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
					break;
				}
				fullSeriesCollection2 = null;
			}
			if (A.Series.Explosion)
			{
				FullSeriesCollection fullSeriesCollection3 = (FullSeriesCollection)B.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count3 = fullSeriesCollection3.Count;
				int num3 = 1;
				while (true)
				{
					if (num3 <= count3)
					{
						if (!C.Explosion.ContainsKey(num3))
						{
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							try
							{
								fullSeriesCollection3.Item(num3).Explosion = C.Explosion[num3];
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
							num3++;
							break;
						}
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
					break;
				}
				fullSeriesCollection3 = null;
			}
			if (A.Series.ErrorBars)
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
				FullSeriesCollection fullSeriesCollection4 = (FullSeriesCollection)B.FullSeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count4 = fullSeriesCollection4.Count;
				int num4 = 1;
				while (true)
				{
					if (num4 <= count4)
					{
						if (!C.ErrorBars.ContainsKey(num4))
						{
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (C.ErrorBars[num4].HasErrorBars)
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
								try
								{
									if (fullSeriesCollection4.Item(num4).HasErrorBars)
									{
										ErrorBars errorBars = fullSeriesCollection4.Item(num4).ErrorBars;
										Apply.A(errorBars.Format.Line, C.ErrorBars[num4].Line);
										errorBars.EndStyle = C.ErrorBars[num4].EndStyle;
										_ = null;
									}
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
								}
							}
							num4++;
							break;
						}
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
					break;
				}
				fullSeriesCollection4 = null;
			}
			if (A.Series.GapWidthOverlap)
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
				ChartGroups chartGroups = (ChartGroups)B.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count5 = chartGroups.Count;
				int num5 = 1;
				while (true)
				{
					if (num5 <= count5)
					{
						if (!C.GapWidth.ContainsKey(num5))
						{
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							try
							{
								chartGroups.Item(num5).GapWidth = C.GapWidth[num5];
							}
							catch (Exception ex11)
							{
								ProjectData.SetProjectError(ex11);
								Exception ex12 = ex11;
								ProjectData.ClearProjectError();
							}
							num5++;
							break;
						}
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
					break;
				}
				int count6 = chartGroups.Count;
				int num6 = 1;
				while (num6 <= count6 && C.Overlap.ContainsKey(num6))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						try
						{
							chartGroups.Item(num6).Overlap = C.Overlap[num6];
						}
						catch (Exception ex13)
						{
							ProjectData.SetProjectError(ex13);
							Exception ex14 = ex13;
							ProjectData.ClearProjectError();
						}
						num6++;
						break;
					}
				}
				chartGroups = null;
			}
			if (A.Series.FirstSliceAngle)
			{
				ChartGroups chartGroups2 = (ChartGroups)B.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count7 = chartGroups2.Count;
				int num7 = 1;
				while (true)
				{
					if (num7 <= count7)
					{
						if (!C.FirstSliceAngle.ContainsKey(num7))
						{
							break;
						}
						try
						{
							chartGroups2.Item(num7).FirstSliceAngle = C.FirstSliceAngle[num7];
						}
						catch (Exception ex15)
						{
							ProjectData.SetProjectError(ex15);
							Exception ex16 = ex15;
							ProjectData.ClearProjectError();
						}
						num7++;
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
					break;
				}
				chartGroups2 = null;
			}
			if (!A.Series.UpDownBars)
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
				ChartGroups chartGroups3 = (ChartGroups)B.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value));
				int count8 = chartGroups3.Count;
				int num8 = 1;
				while (true)
				{
					if (num8 <= count8)
					{
						if (!C.UpBars.ContainsKey(num8))
						{
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (C.UpBars[num8].HasUpDownBars)
							{
								try
								{
									ChartGroup chartGroup = chartGroups3.Item(num8);
									if (chartGroup.HasUpDownBars)
									{
										ChartFormat format2 = chartGroup.UpBars.Format;
										Apply.A(format2.Fill, C.UpBars[num8].Fill);
										Apply.A(format2.Line, C.UpBars[num8].Border);
										_ = null;
										ChartFormat format3 = chartGroup.DownBars.Format;
										Apply.A(format3.Fill, C.DownBars[num8].Fill);
										Apply.A(format3.Line, C.DownBars[num8].Border);
										_ = null;
									}
									chartGroup = null;
								}
								catch (Exception ex17)
								{
									ProjectData.SetProjectError(ex17);
									Exception ex18 = ex17;
									ProjectData.ClearProjectError();
								}
							}
							num8++;
							break;
						}
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
					break;
				}
				chartGroups3 = null;
				return;
			}
		}
	}

	private static void A(Axis A, Properties.AxisProperties B, bool C, bool D, bool E, bool F)
	{
		if (C)
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
			if (A.Type == Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
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
				if (B.MaximumScaleIsAuto)
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
					A.MaximumScaleIsAuto = true;
				}
				else
				{
					A.MaximumScale = B.MaximumScale;
				}
				if (B.MinimumScaleIsAuto)
				{
					A.MinimumScaleIsAuto = true;
				}
				else
				{
					A.MinimumScale = B.MinimumScale;
				}
				if (B.MajorUnitIsAuto)
				{
					A.MajorUnitIsAuto = true;
				}
				else
				{
					A.MajorUnit = B.MajorUnit;
				}
				if (B.MinorUnitIsAuto)
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
					A.MinorUnitIsAuto = true;
				}
				else
				{
					A.MinorUnit = B.MinorUnit;
				}
			}
			else
			{
				try
				{
					A.AxisBetweenCategories = B.AxisBetweenCategories;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					A.MajorUnitScale = B.MajorUnitScale;
					A.MinorUnitScale = B.MinorUnitScale;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			A.Crosses = B.Crosses;
			if (B.Crosses == XlAxisCrosses.xlAxisCrossesCustom)
			{
				A.CrossesAt = B.CrossesAt;
			}
			A.ReversePlotOrder = B.ReversePlotOrder;
		}
		if (D)
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
			A.HasMajorGridlines = B.HasMajorGridlines;
			A.HasMinorGridlines = B.HasMinorGridlines;
			if (B.HasMajorGridlines)
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
				Apply.A(A.MajorGridlines.Format.Line, B.MajorGridlines.Format);
			}
			if (B.HasMinorGridlines)
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
				Apply.A(A.MinorGridlines.Format.Line, B.MajorGridlines.Format);
			}
		}
		if (E)
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
			ref Properties.BorderProperties border = ref B.Border;
			A.Border.Color = border.Color;
			A.Border.Weight = border.Weight;
			A.Border.LineStyle = border.LineStyle;
			A.MajorTickMark = B.MajorTickMark;
			A.MinorTickMark = B.MinorTickMark;
			A.TickLabelPosition = B.TickLabelPosition;
			if (B.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
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
				if (A.Type != Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
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
					A.TickLabelSpacing = B.TickLabelSpacing;
					A.TickLabelSpacingIsAuto = B.TickLabelSpacingIsAuto;
				}
				ref Properties.TickLabelsProperties tickLabels = ref B.TickLabels;
				try
				{
					Apply.A(A.TickLabels.Format.TextFrame2.TextRange.Font, tickLabels.Format);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					Apply.A(A.TickLabels.Font, tickLabels.Format);
					ProjectData.ClearProjectError();
				}
				try
				{
					A.TickLabels.MultiLevel = tickLabels.MultiLevel;
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				try
				{
					A.TickLabels.Offset = tickLabels.Offset;
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
				A.TickLabels.Orientation = (XlTickLabelOrientation)tickLabels.Orientation;
				A.TickLabels.NumberFormatLinked = tickLabels.NumberFormatLinked;
				if (!tickLabels.NumberFormatLinked)
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
					A.TickLabels.NumberFormat = tickLabels.NumberFormat;
				}
				try
				{
					A.TickLabels.Alignment = tickLabels.Alignment;
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
			}
		}
		if (!F)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			A.HasTitle = B.HasTitle;
			if (!B.HasTitle)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				ref Properties.AxisTitleProperties axisTitle = ref B.AxisTitle;
				try
				{
					Apply.A(A.AxisTitle.Format.TextFrame2.TextRange.Font, axisTitle.Font);
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					Apply.A(A.AxisTitle.Font, axisTitle.Font);
					ProjectData.ClearProjectError();
				}
				Apply.A(A.AxisTitle.Format.Fill, axisTitle.Fill);
				Apply.A(A.AxisTitle.Format.Line, axisTitle.Border);
				A.AxisTitle.IncludeInLayout = axisTitle.IncludeInLayout;
				A.AxisTitle.HorizontalAlignment = RuntimeHelpers.GetObjectValue(axisTitle.HorizontalAlignment);
				A.AxisTitle.VerticalAlignment = RuntimeHelpers.GetObjectValue(axisTitle.VerticalAlignment);
				A.AxisTitle.Orientation = RuntimeHelpers.GetObjectValue(axisTitle.Orientation);
				return;
			}
		}
	}

	private static void A(Options A, Legend B, Properties.LegendProperties C)
	{
		if (A.Legend.Format)
		{
			Apply.A(B.Font, C.Font);
			Apply.A(B.Format.Fill, C.Fill);
			Apply.A(B.Format.Line, C.Border);
		}
		if (!A.Legend.Position)
		{
			return;
		}
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
			B.IncludeInLayout = C.IncludeInLayout;
			if (C.Position != XlLegendPosition.xlLegendPositionCustom)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						B.Position = C.Position;
						return;
					}
				}
			}
			B.Top = C.Top;
			B.Left = C.Left;
			return;
		}
	}

	private static void A(Options A, ChartTitle B, Properties.TitleProperties C)
	{
		if (A.Title.Format)
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
			try
			{
				Apply.A(B.Format.TextFrame2.TextRange.Font, C.Font);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Apply.A(B.Font, C.Font);
				ProjectData.ClearProjectError();
			}
			Apply.A(B.Format.Fill, C.Fill);
			Apply.A(B.Format.Line, C.Border);
		}
		if (!A.Title.Position)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			B.IncludeInLayout = C.IncludeInLayout;
			if (C.Position == XlChartElementPosition.xlChartElementPositionAutomatic)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						B.Position = C.Position;
						return;
					}
				}
			}
			B.Left = C.Left;
			B.Top = C.Top;
			return;
		}
	}

	private static void A(Options A, DataTable B, Properties.DataTableProperties C)
	{
		if (!A.DataTable.Format)
		{
			return;
		}
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
			Apply.A(B.Font, C.Font);
			Apply.A(B.Format.Fill, C.Fill);
			Apply.A(B.Format.Line, C.Border);
			B.ShowLegendKey = C.ShowLegendKey;
			B.HasBorderHorizontal = C.HasBorderHorizontal;
			B.HasBorderOutline = C.HasBorderOutline;
			B.HasBorderVertical = C.HasBorderVertical;
			return;
		}
	}

	private static void A(Font2 A, Properties.FontProperties B)
	{
		try
		{
			A.Fill.ForeColor.RGB = B.ForeColor;
			A.Fill.BackColor.RGB = B.BackColor;
			A.Name = B.Name;
			A.Size = B.Size;
			ref Properties.DecorationProperties decoration = ref B.Decoration;
			A.Bold = decoration.Bold;
			A.Italic = decoration.Italic;
			A.UnderlineStyle = decoration.UnderlineStyle;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Font A, Properties.FontProperties B)
	{
		try
		{
			A.Color = B.ForeColor;
			A.Name = B.Name;
			A.Size = B.Size;
			ref Properties.DecorationProperties decoration = ref B.Decoration;
			A.Bold = decoration.Bold;
			A.Italic = decoration.Italic;
			switch (decoration.UnderlineStyle)
			{
			case MsoTextUnderlineType.msoNoUnderline:
				A.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
				break;
			case MsoTextUnderlineType.msoUnderlineSingleLine:
				A.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
				break;
			case MsoTextUnderlineType.msoUnderlineDoubleLine:
				A.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
				break;
			case MsoTextUnderlineType.msoUnderlineWords:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.FillFormat A, Properties.FillProperties B)
	{
		checked
		{
			try
			{
				if (B.Visible == MsoTriState.msoTrue)
				{
					A.Visible = MsoTriState.msoTrue;
					switch (B.Type)
					{
					case MsoFillType.msoFillSolid:
						A.ForeColor.RGB = B.ForeColor;
						A.BackColor.RGB = B.BackColor;
						A.Solid();
						break;
					case MsoFillType.msoFillPatterned:
						A.Patterned(B.Pattern);
						A.ForeColor.RGB = B.ForeColor;
						A.BackColor.RGB = B.BackColor;
						break;
					case MsoFillType.msoFillGradient:
						switch (B.GradientColorType)
						{
						case MsoGradientColorType.msoGradientTwoColors:
						case MsoGradientColorType.msoGradientMultiColor:
						{
							A.TwoColorGradient(B.GradientStyle, B.GradientVariant);
							A.GradientAngle = B.GradientAngle;
							int num = B.GradientStops.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								if (i < 2)
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
									GradientStop gradientStop = A.GradientStops[i + 1];
									Properties.GradientStopProperties gradientStopProperties = B.GradientStops[i];
									if (gradientStopProperties.Type == MsoColorType.msoColorTypeScheme)
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
										gradientStop.Color.SchemeColor = gradientStopProperties.SchemeColor;
										gradientStop.Color.Brightness = gradientStopProperties.Brightness;
									}
									else
									{
										gradientStop.Color.RGB = gradientStopProperties.RGB;
										gradientStop.Color.Brightness = 0f;
									}
									gradientStop.Position = gradientStopProperties.Position;
									gradientStop.Transparency = gradientStopProperties.Transparency;
									gradientStop = null;
									continue;
								}
								Properties.GradientStopProperties gradientStopProperties2 = B.GradientStops[i];
								if (gradientStopProperties2.Type == MsoColorType.msoColorTypeScheme)
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
									int rGB = MH.A.Application.ActiveWorkbook.Theme.ThemeColorScheme.Colors(unchecked((MsoThemeColorSchemeIndex)gradientStopProperties2.ObjectThemeColor)).RGB;
									A.GradientStops.Insert2(rGB, gradientStopProperties2.Position, gradientStopProperties2.Transparency, i + 1, gradientStopProperties2.Brightness);
								}
								else
								{
									A.GradientStops.Insert2(gradientStopProperties2.RGB, gradientStopProperties2.Position, gradientStopProperties2.Transparency, i + 1);
								}
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
							break;
						}
						case MsoGradientColorType.msoGradientOneColor:
							A.OneColorGradient(B.GradientStyle, B.GradientVariant, B.GradientDegree);
							A.GradientAngle = B.GradientAngle;
							break;
						case MsoGradientColorType.msoGradientPresetColors:
							A.PresetGradient(B.GradientStyle, B.GradientVariant, B.PresetGradientType);
							break;
						}
						break;
					}
					try
					{
						A.Transparency = B.Transparency;
						return;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						return;
					}
				}
				A.Visible = MsoTriState.msoFalse;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void A(LineFormat A, Properties.LineProperties B)
	{
		try
		{
			if (B.Visible == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A.Visible = MsoTriState.msoTrue;
						A.ForeColor.RGB = B.ForeColor;
						A.Transparency = B.Transparency;
						if (B.Style != MsoLineStyle.msoLineStyleMixed)
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
							A.Style = B.Style;
						}
						else
						{
							B.Visible = MsoTriState.msoFalse;
						}
						if (B.DashStyle != MsoLineDashStyle.msoLineDashStyleMixed)
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
							A.DashStyle = B.DashStyle;
						}
						else
						{
							B.Visible = MsoTriState.msoFalse;
						}
						A.Weight = Math.Max(0f, B.Weight);
						return;
					}
				}
			}
			A.Visible = MsoTriState.msoFalse;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Border A, Properties.LineProperties B)
	{
		try
		{
			Operators.ConditionalCompareObjectNotEqual(A.LineStyle, XlLineStyle.xlLineStyleNone, TextCompare: false);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A()
	{
		MessageBox.Show(VH.A(172145), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}
}
