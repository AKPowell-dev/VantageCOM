using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ChartFonts
{
	internal static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, int? C, int? D)
	{
		if (!C.HasValue)
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
			if (!D.HasValue)
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
				break;
			}
		}
		checked
		{
			try
			{
				Chart chart = B.Chart;
				if (chart.HasTitle)
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
						float num = Conversions.ToSingle(chart.ChartTitle.Font.Size);
						if (Fonts.A(num, C))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new ChartTitleFontSize(A, B, num, C.Value, chart.ChartTitle));
								break;
							}
						}
						else if (Fonts.B(num, D))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new ChartTitleFontSize(A, B, num, D.Value, chart.ChartTitle));
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
				if (chart.HasLegend)
				{
					try
					{
						float num = Conversions.ToSingle(chart.Legend.Font.Size);
						if (Fonts.A(num, C))
						{
							Main.Analysis.Errors.Add(new ChartLegendFontSize(A, B, num, C.Value, chart.Legend));
						}
						else if (Fonts.B(num, D))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new ChartLegendFontSize(A, B, num, D.Value, chart.Legend));
								break;
							}
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
						IEnumerator enumerator = ((IEnumerable)chart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						try
						{
							while (true)
							{
								if (enumerator.MoveNext())
								{
									float num = Conversions.ToSingle(((Microsoft.Office.Interop.PowerPoint.LegendEntry)enumerator.Current).Font.Size);
									if (Fonts.A(num, C))
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											Main.Analysis.Errors.Add(new ChartLegendFontSize(A, B, num, C.Value, chart.Legend));
											break;
										}
										break;
									}
									if (!Fonts.B(num, D))
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
										Main.Analysis.Errors.Add(new ChartLegendFontSize(A, B, num, D.Value, chart.Legend));
										break;
									}
									break;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0289;
									}
									continue;
									end_IL_0289:
									break;
								}
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
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
				}
				if (chart.HasDataTable)
				{
					try
					{
						float num = Conversions.ToSingle(chart.DataTable.Font.Size);
						if (Fonts.A(num, C))
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new ChartDataTableFontSize(A, B, num, C.Value, chart.DataTable));
								break;
							}
						}
						else if (Fonts.B(num, D))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new ChartDataTableFontSize(A, B, num, D.Value, chart.DataTable));
								break;
							}
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
				try
				{
					foreach (Axis item in modCharts.AxesList(chart))
					{
						try
						{
							if (item.HasTitle)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									float num = Conversions.ToSingle(item.AxisTitle.Font.Size);
									if (Fonts.A(num, C))
									{
										Main.Analysis.Errors.Add(new ChartAxisTitleFontSize(A, B, num, C.Value, item.AxisTitle));
										break;
									}
									if (!Fonts.B(num, D))
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
										Main.Analysis.Errors.Add(new ChartAxisTitleFontSize(A, B, num, D.Value, item.AxisTitle));
										break;
									}
									break;
								}
							}
						}
						catch (Exception ex9)
						{
							ProjectData.SetProjectError(ex9);
							Exception ex10 = ex9;
							ProjectData.ClearProjectError();
						}
						try
						{
							float num = Conversions.ToSingle(item.TickLabels.Font.Size);
							if (Fonts.A(num, C))
							{
								Main.Analysis.Errors.Add(new ChartTickLabelsFontSize(A, B, num, C.Value, item));
							}
							else
							{
								if (!Fonts.B(num, D))
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
									Main.Analysis.Errors.Add(new ChartTickLabelsFontSize(A, B, num, D.Value, item));
									break;
								}
								continue;
							}
						}
						catch (Exception ex11)
						{
							ProjectData.SetProjectError(ex11);
							Exception ex12 = ex11;
							ProjectData.ClearProjectError();
						}
					}
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					ProjectData.ClearProjectError();
				}
				try
				{
					IEnumerator enumerator3 = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					try
					{
						while (enumerator3.MoveNext())
						{
							IMsoSeries msoSeries = (IMsoSeries)enumerator3.Current;
							try
							{
								float num = Conversions.ToSingle(((IMsoDataLabels)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Size);
								if (Fonts.A(num, C))
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										Main.Analysis.Errors.Add(new ChartDataLabelsFontSize(A, B, num, C.Value, msoSeries));
										break;
									}
									continue;
								}
								if (Fonts.B(num, D))
								{
									Main.Analysis.Errors.Add(new ChartDataLabelsFontSize(A, B, num, D.Value, msoSeries));
									continue;
								}
							}
							catch (Exception ex15)
							{
								ProjectData.SetProjectError(ex15);
								Exception ex16 = ex15;
								ProjectData.ClearProjectError();
							}
							try
							{
								int num2 = 0;
								int num3 = Conversions.ToInteger(NewLateBinding.LateGet(msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value)), null, AH.A(13955), new object[0], null, null, null));
								int num4 = 1;
								while (num4 <= num3 && num2 != 25)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										object instance = msoSeries.Points(num4);
										if (Conversions.ToBoolean(NewLateBinding.LateGet(instance, null, AH.A(13966), new object[0], null, null, null)))
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
											float num = Conversions.ToSingle(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, AH.A(13991), new object[0], null, null, null), null, AH.A(14010), new object[0], null, null, null), null, AH.A(14019), new object[0], null, null, null));
											if (Fonts.A(num, C))
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
												Main.Analysis.Errors.Add(new ChartDataLabelFontSize(A, B, num, C.Value, (IMsoDataLabel)NewLateBinding.LateGet(instance, null, AH.A(13991), new object[0], null, null, null)));
											}
											else if (Fonts.B(num, D))
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
												Main.Analysis.Errors.Add(new ChartDataLabelFontSize(A, B, num, D.Value, (IMsoDataLabel)NewLateBinding.LateGet(instance, null, AH.A(13991), new object[0], null, null, null)));
											}
										}
										instance = null;
										num2++;
										num4++;
										break;
									}
								}
							}
							catch (Exception ex17)
							{
								ProjectData.SetProjectError(ex17);
								Exception ex18 = ex17;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_07f0;
							}
							continue;
							end_IL_07f0:
							break;
						}
					}
					finally
					{
						IDisposable disposable2 = enumerator3 as IDisposable;
						if (disposable2 != null)
						{
							disposable2.Dispose();
						}
					}
				}
				catch (Exception ex19)
				{
					ProjectData.SetProjectError(ex19);
					Exception ex20 = ex19;
					ProjectData.ClearProjectError();
				}
				chart = null;
			}
			catch (Exception ex21)
			{
				ProjectData.SetProjectError(ex21);
				Exception ex22 = ex21;
				ProjectData.ClearProjectError();
			}
		}
	}
}
