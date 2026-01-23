using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Colors;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ChartColors
{
	public static void ChartArea(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int rGB = shp.Chart.ChartArea.Format.Fill.ForeColor.RGB;
			if (!listColors.Contains(rGB))
			{
				Main.Analysis.Errors.Add(new ChartAreaFillColor(sld, shp, rGB, sev));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			int rGB = shp.Chart.ChartArea.Format.Line.ForeColor.RGB;
			if (listColors.Contains(rGB))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Main.Analysis.Errors.Add(new ChartAreaBorderColor(sld, shp, rGB, sev));
				return;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	public static void PlotArea(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int rGB = shp.Chart.PlotArea.Format.Fill.ForeColor.RGB;
			if (!listColors.Contains(rGB))
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Main.Analysis.Errors.Add(new PlotAreaFillColor(sld, shp, rGB, shp.Chart.PlotArea, sev));
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
		try
		{
			int rGB = shp.Chart.PlotArea.Format.Line.ForeColor.RGB;
			if (listColors.Contains(rGB))
			{
				return;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				Main.Analysis.Errors.Add(new PlotAreaBorderColor(sld, shp, rGB, shp.Chart.PlotArea, sev));
				return;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	public static void Axes(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0502: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_015e: Unknown result type (might be due to invalid IL or missing references)
		//IL_023f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0424: Unknown result type (might be due to invalid IL or missing references)
		//IL_0313: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_039b: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			bool flag = clsCharts.A(shp.Chart, new int[2] { 90, 91 });
			using List<Axis>.Enumerator enumerator = modCharts.AxesList(shp.Chart).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Axis current = enumerator.Current;
				try
				{
					Axis axis = current;
					Microsoft.Office.Interop.PowerPoint.XlAxisType type = current.Type;
					XlAxisGroup axisGroup = current.AxisGroup;
					if (axisGroup != XlAxisGroup.xlPrimary)
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
						if (flag)
						{
							goto IL_016e;
						}
					}
					if (axis.HasMajorGridlines)
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
						int rGB = axis.MajorGridlines.Format.Line.ForeColor.RGB;
						if (Color.ColorNotInPalette(rGB, listColors))
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
							Main.Analysis.Errors.Add(new GridlinesColor(sld, shp, rGB, shp.Chart.PlotArea, sev, type, axisGroup, areMajor: true));
						}
					}
					if (axis.HasMinorGridlines)
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
						int rGB = axis.MinorGridlines.Format.Line.ForeColor.RGB;
						if (Color.ColorNotInPalette(rGB, listColors))
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
							Main.Analysis.Errors.Add(new GridlinesColor(sld, shp, rGB, shp.Chart.PlotArea, sev, type, axisGroup, areMajor: false));
						}
					}
					goto IL_016e;
					IL_016e:
					if (axis.HasTitle)
					{
						AxisTitle axisTitle = axis.AxisTitle;
						AxisTitle axisTitle2 = axis.AxisTitle;
						Font2 font = axisTitle2.Format.TextFrame2.TextRange.Font;
						int rGB = font.Fill.ForeColor.RGB;
						if (Color.ColorNotInPalette(rGB, listColors))
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
							Main.Analysis.Errors.Add(new AxisTitleFontColor(sld, shp, rGB, axisGroup, axisTitle, sev));
						}
						if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
							rGB = font.UnderlineColor.RGB;
							if (rGB >= 0 && Color.ColorNotInPalette(rGB, listColors))
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
								Main.Analysis.Errors.Add(new AxisTitleUnderlineColor(sld, shp, rGB, axisGroup, axisTitle, sev));
							}
						}
						if (font.Line.Visible == MsoTriState.msoTrue)
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
							rGB = font.Line.ForeColor.RGB;
							if (rGB >= 0)
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
								if (Color.ColorNotInPalette(rGB, listColors))
								{
									Main.Analysis.Errors.Add(new AxisTitleOutlineColor(sld, shp, rGB, axisGroup, axisTitle, sev));
								}
							}
						}
						font = null;
						Microsoft.Office.Interop.PowerPoint.FillFormat fill = axisTitle2.Format.Fill;
						if (fill.Visible == MsoTriState.msoTrue)
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
							rGB = fill.ForeColor.RGB;
							if (Color.ColorNotInPalette(rGB, listColors))
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
								Main.Analysis.Errors.Add(new AxisTitleFillColor(sld, shp, rGB, axisGroup, axisTitle, sev));
							}
						}
						fill = null;
						Microsoft.Office.Interop.PowerPoint.LineFormat line = axisTitle2.Format.Line;
						if (line.Weight > 0f)
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
							if (line.Visible == MsoTriState.msoTrue)
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
								rGB = line.ForeColor.RGB;
								if (Color.ColorNotInPalette(rGB, listColors))
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
									Main.Analysis.Errors.Add(new AxisTitleBorderColor(sld, shp, rGB, axisGroup, axisTitle, sev));
								}
							}
						}
						line = null;
						_ = null;
						axisTitle = null;
					}
					try
					{
						Microsoft.Office.Interop.PowerPoint.LineFormat line2 = axis.Format.Line;
						if (line2.Weight > 0f)
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
							if (line2.Visible == MsoTriState.msoTrue)
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
								int rGB = line2.ForeColor.RGB;
								if (Color.ColorNotInPalette(rGB, listColors))
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
									Main.Analysis.Errors.Add(new AxisBorderColor(sld, shp, rGB, axisGroup, current, sev));
								}
							}
						}
						line2 = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					try
					{
						Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = axis.Format.Fill;
						if (fill2.Visible == MsoTriState.msoTrue)
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
							int rGB = fill2.ForeColor.RGB;
							if (Color.ColorNotInPalette(rGB, listColors))
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
								Main.Analysis.Errors.Add(new AxisFillColor(sld, shp, rGB, axisGroup, current, sev));
							}
						}
						fill2 = null;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					try
					{
						int rGB = Conversions.ToInteger(axis.TickLabels.Font.Color);
						if (Color.ColorNotInPalette(rGB, listColors))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new AxisFontColor(sld, shp, rGB, axisGroup, current, sev));
								break;
							}
						}
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception a = ex5;
						Main.A(a, null, shp.Chart);
						ProjectData.ClearProjectError();
					}
					axis = null;
				}
				finally
				{
					current = null;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		catch (Exception ex6)
		{
			ProjectData.SetProjectError(ex6);
			Exception ex7 = ex6;
			ProjectData.ClearProjectError();
		}
	}

	public static void Series(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_01e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_048c: Unknown result type (might be due to invalid IL or missing references)
		//IL_064d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0184: Unknown result type (might be due to invalid IL or missing references)
		//IL_0805: Unknown result type (might be due to invalid IL or missing references)
		//IL_0402: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ce: Unknown result type (might be due to invalid IL or missing references)
		Chart chart = shp.Chart;
		bool flag = clsCharts.UsesFormatFillForSeriesClrs(chart);
		bool flag2 = clsCharts.UsesLegendLinesForSeriesClrs(chart);
		bool flag3 = clsCharts.UsesMarkers(chart);
		LegendEntries legendEntries = null;
		try
		{
			if (clsCharts.CanIgnoreLegendClrs(chart))
			{
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
					break;
				}
			}
			else
			{
				legendEntries = (LegendEntries)chart.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value));
				if (Operators.ConditionalCompareObjectNotEqual(legendEntries.Count, NewLateBinding.LateGet(chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, AH.A(13955), new object[0], null, null, null), TextCompare: false))
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						legendEntries = null;
						break;
					}
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		int num = 0;
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
					num++;
					int? num2 = null;
					int? num3 = null;
					if (flag)
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
							num2 = msoSeries.Format.Fill.ForeColor.RGB;
							if (!listColors.Contains(num2.Value))
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									Main.Analysis.Errors.Add(new SeriesColor(sld, shp, num2.Value, msoSeries, sev));
									break;
								}
							}
						}
						catch (Exception projectError2)
						{
							ProjectData.SetProjectError(projectError2);
							ProjectData.ClearProjectError();
						}
						try
						{
							num3 = msoSeries.Format.Line.ForeColor.RGB;
							if (!listColors.Contains(num3.Value))
							{
								Main.Analysis.Errors.Add(new SeriesLineColor(sld, shp, num3.Value, msoSeries, sev));
							}
						}
						catch (Exception projectError3)
						{
							ProjectData.SetProjectError(projectError3);
							ProjectData.ClearProjectError();
						}
					}
					else if (legendEntries != null)
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
						object objectValue;
						try
						{
							objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(legendEntries.Cast<object>().ElementAtOrDefault(num - 1), null, AH.A(13177), new object[0], null, null, null));
							if (flag2)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									num3 = (int?)NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[0], null, null, null);
									break;
								}
							}
							else
							{
								num2 = (int?)NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null);
							}
						}
						catch (Exception projectError4)
						{
							ProjectData.SetProjectError(projectError4);
							ProjectData.ClearProjectError();
						}
						objectValue = null;
					}
					if (Charts.ImplsPoints(msoSeries))
					{
						try
						{
							enumerator2 = ((IEnumerable)msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								if (!(RuntimeHelpers.GetObjectValue(enumerator2.Current) is ChartPoint chartPoint))
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_037f;
										}
										continue;
										end_IL_037f:
										break;
									}
									break;
								}
								object obj = null;
								object obj2 = null;
								try
								{
									obj = RuntimeHelpers.GetObjectValue(chartPoint.Interior.Color);
									if (!listColors.Contains(Conversions.ToInteger(obj)))
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											if (object.Equals(RuntimeHelpers.GetObjectValue(obj), num2))
											{
												break;
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												Main.Analysis.Errors.Add(new SeriesPointColor(sld, shp, Conversions.ToInteger(obj), chartPoint, sev));
												break;
											}
											break;
										}
									}
								}
								catch (Exception projectError5)
								{
									ProjectData.SetProjectError(projectError5);
									ProjectData.ClearProjectError();
								}
								try
								{
									obj2 = chartPoint.Format.Line.ForeColor.RGB;
									if (!listColors.Contains(Conversions.ToInteger(obj2)))
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											if (!object.Equals(RuntimeHelpers.GetObjectValue(obj2), num3))
											{
												Main.Analysis.Errors.Add(new SeriesPointLineColor(sld, shp, Conversions.ToInteger(obj2), chartPoint, sev));
											}
											break;
										}
									}
								}
								catch (Exception projectError6)
								{
									ProjectData.SetProjectError(projectError6);
									ProjectData.ClearProjectError();
								}
								if (!flag3)
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
								if (chartPoint.MarkerStyle == XlMarkerStyle.xlMarkerStyleNone)
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
								int markerForegroundColor = chartPoint.MarkerForegroundColor;
								int markerBackgroundColor = chartPoint.MarkerBackgroundColor;
								bool flag4 = object.Equals(markerForegroundColor, markerBackgroundColor);
								int? num4 = null;
								int? num5 = null;
								if (legendEntries != null)
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
										object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(legendEntries.Cast<object>().ElementAtOrDefault(num - 1), null, AH.A(13177), new object[0], null, null, null));
										num4 = (int?)NewLateBinding.LateGet(objectValue2, null, AH.A(14093), new object[0], null, null, null);
										num5 = (int?)NewLateBinding.LateGet(objectValue2, null, AH.A(14136), new object[0], null, null, null);
									}
									catch (Exception projectError7)
									{
										ProjectData.SetProjectError(projectError7);
										ProjectData.ClearProjectError();
									}
								}
								object[] source = new object[5]
								{
									Base.TRANSPARENT,
									num2,
									num3,
									obj,
									obj2
								};
								if (!listColors.Contains(markerForegroundColor) && !source.Contains(markerForegroundColor))
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
									if (!object.Equals(markerForegroundColor, num4))
									{
										bool? isFore = true;
										if (flag4)
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
											isFore = null;
										}
										Main.Analysis.Errors.Add(new SeriesPointMarkerColor(sld, shp, markerForegroundColor, chartPoint, sev, isFore));
									}
								}
								if (listColors.Contains(markerBackgroundColor))
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
								if (source.Contains(markerBackgroundColor))
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
								if (object.Equals(markerBackgroundColor, num5))
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
								if (flag4)
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
								Main.Analysis.Errors.Add(new SeriesPointMarkerColor(sld, shp, markerBackgroundColor, chartPoint, sev, false));
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (3)
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
					try
					{
						int num6 = Conversions.ToInteger(((IMsoDataLabels)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color);
						if (listColors.Contains(num6))
						{
							continue;
						}
						List<IMsoDataLabel> list = new List<IMsoDataLabel>();
						try
						{
							enumerator3 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator3.MoveNext())
							{
								IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator3.Current;
								if (listColors.Contains(Conversions.ToInteger(msoDataLabel.Font.Color)))
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
									break;
								}
								list.Add(msoDataLabel);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_07c5;
								}
								continue;
								end_IL_07c5:
								break;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									(enumerator3 as IDisposable).Dispose();
									break;
								}
							}
						}
						Main.Analysis.Errors.Add(new LabelsColor(sld, shp, num6, msoSeries, list, sev));
						list = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
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
			legendEntries = null;
			chart = null;
		}
	}

	public static void Legend(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0245: Unknown result type (might be due to invalid IL or missing references)
		//IL_015a: Unknown result type (might be due to invalid IL or missing references)
		//IL_069b: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0651: Unknown result type (might be due to invalid IL or missing references)
		Chart chart = shp.Chart;
		try
		{
			if (!clsCharts.CanIgnoreLegendClrs(chart))
			{
				IEnumerator enumerator = default(IEnumerator);
				IEnumerator enumerator2 = default(IEnumerator);
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
					Legend legend = shp.Chart.Legend;
					Font2 font = legend.Format.TextFrame2.TextRange.Font;
					int rGB = font.Fill.ForeColor.RGB;
					if (!listColors.Contains(rGB))
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
						Main.Analysis.Errors.Add(new LegendFontColor(sld, shp, rGB, legend, sev));
					}
					if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
					{
						rGB = font.UnderlineColor.RGB;
						if (rGB >= 0)
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
							if (Color.ColorNotInPalette(rGB, listColors))
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
								Main.Analysis.Errors.Add(new LegendUnderlineColor(sld, shp, rGB, legend, sev));
							}
						}
					}
					if (font.Line.Visible == MsoTriState.msoTrue)
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
						rGB = font.Line.ForeColor.RGB;
						if (rGB >= 0)
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
							if (Color.ColorNotInPalette(rGB, listColors))
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
								Main.Analysis.Errors.Add(new LegendOutlineColor(sld, shp, rGB, legend, sev));
							}
						}
					}
					font = null;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = legend.Format.Fill;
					if (fill.Visible == MsoTriState.msoTrue)
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
						Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = fill.ForeColor;
						if (!listColors.Contains(foreColor.RGB))
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
							Main.Analysis.Errors.Add(new LegendFillColor(sld, shp, foreColor.RGB, legend, sev));
						}
						foreColor = null;
					}
					fill = null;
					Microsoft.Office.Interop.PowerPoint.LineFormat line = legend.Format.Line;
					if (line.Visible == MsoTriState.msoTrue)
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
						Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = line.ForeColor;
						if (!listColors.Contains(foreColor2.RGB))
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
							Main.Analysis.Errors.Add(new LegendBorderColor(sld, shp, foreColor2.RGB, legend, sev));
						}
						foreColor2 = null;
					}
					line = null;
					bool flag = clsCharts.UsesLegendLinesForSeriesClrs(chart);
					if (!clsCharts.UsesFormatFillForSeriesClrs(chart))
					{
						int num = 0;
						try
						{
							enumerator = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								object objectValue2;
								try
								{
									objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, AH.A(13177), new object[0], null, null, null));
									object obj;
									if (!flag)
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
										obj = NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null);
									}
									else
									{
										obj = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue2, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[0], null, null, null);
									}
									object objectValue3 = RuntimeHelpers.GetObjectValue(obj);
									if (!listColors.Contains(Conversions.ToInteger(objectValue3)))
									{
										Main.Analysis.Errors.Add(new LegendEntryColor(sld, shp, Conversions.ToInteger(objectValue3), RuntimeHelpers.GetObjectValue(objectValue2), sev, num));
									}
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									ProjectData.ClearProjectError();
								}
								objectValue2 = null;
								num = checked(num + 1);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_03e5;
								}
								continue;
								end_IL_03e5:
								break;
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (1)
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
					if (clsCharts.UsesMarkers(chart))
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
						int num2 = 0;
						{
							enumerator2 = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							try
							{
								while (enumerator2.MoveNext())
								{
									object objectValue4 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
									object objectValue5;
									try
									{
										objectValue5 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue4, null, AH.A(13177), new object[0], null, null, null));
										object objectValue6 = RuntimeHelpers.GetObjectValue(flag ? NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue5, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[0], null, null, null) : NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue5, null, AH.A(14076), new object[0], null, null, null), null, AH.A(13587), new object[0], null, null, null));
										object objectValue7 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue5, null, AH.A(14136), new object[0], null, null, null));
										object objectValue8 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue5, null, AH.A(14093), new object[0], null, null, null));
										bool flag2 = !object.Equals(RuntimeHelpers.GetObjectValue(objectValue7), RuntimeHelpers.GetObjectValue(objectValue6)) && !listColors.Contains(Conversions.ToInteger(objectValue7));
										int num3;
										if (!object.Equals(RuntimeHelpers.GetObjectValue(objectValue8), RuntimeHelpers.GetObjectValue(objectValue6)))
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
											num3 = ((!listColors.Contains(Conversions.ToInteger(objectValue8))) ? 1 : 0);
										}
										else
										{
											num3 = 0;
										}
										bool flag3 = (byte)num3 != 0;
										if (!flag2)
										{
											goto IL_066d;
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
										if (!flag3)
										{
											goto IL_066d;
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
										if (!object.Equals(RuntimeHelpers.GetObjectValue(objectValue6), RuntimeHelpers.GetObjectValue(objectValue7)))
										{
											goto IL_066d;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											Main.Analysis.Errors.Add(new LegendMarkerColor(sld, shp, Conversions.ToInteger(objectValue8), RuntimeHelpers.GetObjectValue(objectValue5), sev, num2, null));
											break;
										}
										goto end_IL_045c;
										IL_066d:
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
											Main.Analysis.Errors.Add(new LegendMarkerColor(sld, shp, Conversions.ToInteger(objectValue8), RuntimeHelpers.GetObjectValue(objectValue5), sev, num2, true));
										}
										if (flag2)
										{
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												Main.Analysis.Errors.Add(new LegendMarkerColor(sld, shp, Conversions.ToInteger(objectValue8), RuntimeHelpers.GetObjectValue(objectValue5), sev, num2, false));
												break;
											}
										}
										end_IL_045c:;
									}
									catch (Exception projectError2)
									{
										ProjectData.SetProjectError(projectError2);
										ProjectData.ClearProjectError();
									}
									objectValue5 = null;
									num2 = checked(num2 + 1);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0710;
									}
									continue;
									end_IL_0710:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator2 as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
						}
					}
					legend = null;
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
		chart = null;
	}

	public static void Title(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_021c: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!shp.Chart.HasTitle)
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
				ChartTitle chartTitle = shp.Chart.ChartTitle;
				Font2 font = chartTitle.Format.TextFrame2.TextRange.Font;
				int rGB = font.Fill.ForeColor.RGB;
				if (!listColors.Contains(rGB))
				{
					Main.Analysis.Errors.Add(new TitleFontColor(sld, shp, rGB, chartTitle, sev));
				}
				if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
					rGB = font.UnderlineColor.RGB;
					if (rGB >= 0)
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
						if (Color.ColorNotInPalette(rGB, listColors))
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
							Main.Analysis.Errors.Add(new TitleUnderlineColor(sld, shp, rGB, chartTitle, sev));
						}
					}
				}
				if (font.Line.Visible == MsoTriState.msoTrue)
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
					rGB = font.Line.ForeColor.RGB;
					if (rGB >= 0)
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
						if (Color.ColorNotInPalette(rGB, listColors))
						{
							Main.Analysis.Errors.Add(new TitleOutlineColor(sld, shp, rGB, chartTitle, sev));
						}
					}
				}
				font = null;
				Microsoft.Office.Interop.PowerPoint.FillFormat fill = chartTitle.Format.Fill;
				if (fill.Visible == MsoTriState.msoTrue)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = fill.ForeColor;
					if (!listColors.Contains(foreColor.RGB))
					{
						Main.Analysis.Errors.Add(new TitleFillColor(sld, shp, foreColor.RGB, chartTitle, sev));
					}
					foreColor = null;
				}
				fill = null;
				Microsoft.Office.Interop.PowerPoint.LineFormat line = chartTitle.Format.Line;
				if (line.Visible == MsoTriState.msoTrue)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = line.ForeColor;
					if (!listColors.Contains(foreColor2.RGB))
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
						Main.Analysis.Errors.Add(new TitleBorderColor(sld, shp, foreColor2.RGB, chartTitle, sev));
					}
					foreColor2 = null;
				}
				line = null;
				chartTitle = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DataTable(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0158: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!shp.Chart.HasDataTable)
			{
				return;
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
				DataTable dataTable = shp.Chart.DataTable;
				_ = dataTable.Format;
				Microsoft.Office.Interop.PowerPoint.ChartFont font = dataTable.Font;
				if (!listColors.Contains(Conversions.ToInteger(font.Color)))
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
					Main.Analysis.Errors.Add(new DataTableFontColor(sld, shp, Conversions.ToInteger(font.Color), dataTable, sev));
				}
				font = null;
				Microsoft.Office.Interop.PowerPoint.FillFormat fill = dataTable.Format.Fill;
				if (fill.Visible == MsoTriState.msoTrue)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = fill.ForeColor;
					if (!listColors.Contains(foreColor.RGB))
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
						Main.Analysis.Errors.Add(new DataTableFillColor(sld, shp, foreColor.RGB, dataTable, sev));
					}
					foreColor = null;
				}
				fill = null;
				Microsoft.Office.Interop.PowerPoint.LineFormat line = dataTable.Format.Line;
				if (line.Visible == MsoTriState.msoTrue)
				{
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = line.ForeColor;
					if (!listColors.Contains(foreColor2.RGB))
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
						Main.Analysis.Errors.Add(new DataTableBorderColor(sld, shp, foreColor2.RGB, dataTable, sev));
					}
					foreColor2 = null;
				}
				line = null;
				_ = null;
				dataTable = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Groups(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_0193: Unknown result type (might be due to invalid IL or missing references)
		//IL_0121: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f9: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int count = ((ChartGroups)shp.Chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				ChartGroup chartGroup = (ChartGroup)shp.Chart.ChartGroups(i);
				ChartGroup chartGroup2 = chartGroup;
				if (chartGroup2.HasUpDownBars)
				{
					_ = chartGroup2.UpBars.Format.Fill;
					_ = null;
					_ = chartGroup2.DownBars.Format.Fill;
					_ = null;
					_ = chartGroup2.UpBars.Format.Line;
					_ = null;
					_ = chartGroup2.DownBars.Format.Line;
					_ = null;
				}
				if (chartGroup2.HasHiLoLines)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = chartGroup2.HiLoLines.Format.Line.ForeColor;
					if (!listColors.Contains(foreColor.RGB))
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
						Main.Analysis.Errors.Add(new HiLoLinesColor(sld, shp, foreColor.RGB, chartGroup, sev));
					}
					foreColor = null;
				}
				if (chartGroup2.HasDropLines)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = chartGroup2.DropLines.Format.Line.ForeColor;
					if (!listColors.Contains(foreColor2.RGB))
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
						Main.Analysis.Errors.Add(new DropLinesColor(sld, shp, foreColor2.RGB, chartGroup, sev));
					}
					foreColor2 = null;
				}
				if (Charts.HasRadarAxisLabels(chartGroup))
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
						int num = Conversions.ToInteger(chartGroup2.RadarAxisLabels.Font.Color);
						if (!listColors.Contains(num))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new AxisRadarFontColor(sld, shp, num, chartGroup, sev));
								break;
							}
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
				chartGroup2 = null;
				chartGroup = null;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
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
}
