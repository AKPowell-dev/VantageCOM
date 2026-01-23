using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Recolor
{
	public static void SeriesToDefaults()
	{
		if (!A())
		{
			return;
		}
		Application application = MH.A.Application;
		Chart chart = null;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		XlCalculation calculation = application.Calculation;
		application.Calculation = XlCalculation.xlCalculationManual;
		try
		{
			if (application.ActiveSheet is Worksheet)
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
				chart = application.ActiveChart;
			}
			else
			{
				chart = (Chart)application.ActiveSheet;
			}
			if (chart != null)
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
				A(chart);
			}
			else
			{
				A(application, A);
			}
			A(VH.A(69441));
		}
		catch (NotImplementedException ex)
		{
			ProjectData.SetProjectError(ex);
			NotImplementedException ex2 = ex;
			Helpers.B();
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application.Calculation = calculation;
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		JH.A((object)chart);
		chart = null;
	}

	private static void A(Chart A)
	{
		int C = 1;
		int count = KH.A.ChartSeriesColors.Count;
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		if (count > 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						enumerator = ((IEnumerable)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator.MoveNext())
						{
							Series series = (Series)enumerator.Current;
							XlChartType chartType = series.ChartType;
							if (chartType <= XlChartType.xl3DPie)
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
								if (chartType != XlChartType.xlRadar)
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
									if (chartType != XlChartType.xlDoughnut)
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
										if (chartType != XlChartType.xl3DPie)
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
											goto IL_01de;
										}
									}
									goto IL_0169;
								}
							}
							else if (chartType <= XlChartType.xlLine)
							{
								if (chartType != XlChartType.xl3DLine)
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
									if (chartType != XlChartType.xlLine)
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
										goto IL_01de;
									}
								}
							}
							else
							{
								if (chartType == XlChartType.xlPie)
								{
									goto IL_0169;
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
								switch (chartType)
								{
								case XlChartType.xlLineStacked:
								case XlChartType.xlLineStacked100:
								case XlChartType.xlLineMarkers:
								case XlChartType.xlLineMarkersStacked:
								case XlChartType.xlLineMarkersStacked100:
									break;
								case XlChartType.xlPieOfPie:
								case XlChartType.xlPieExploded:
								case XlChartType.xl3DPieExploded:
								case XlChartType.xlBarOfPie:
								case XlChartType.xlDoughnutExploded:
									goto IL_0169;
								default:
									goto IL_01de;
								}
							}
							B(series.Format, count, ref C);
							continue;
							IL_0169:
							try
							{
								enumerator2 = ((IEnumerable)series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
								while (enumerator2.MoveNext())
								{
									Recolor.A(((Point)enumerator2.Current).Format, count, ref C);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										return;
									}
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
											break;
										default:
											(enumerator2 as IDisposable).Dispose();
											goto end_IL_01c7;
										}
										continue;
										end_IL_01c7:
										break;
									}
								}
							}
							IL_01de:
							Recolor.A(series.Format, count, ref C);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return;
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
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_020f;
								}
								continue;
								end_IL_020f:
								break;
							}
						}
					}
				}
			}
		}
		Forms.WarningMessage(VH.A(69494));
	}

	private static void A(ChartFormat A, int B, ref int C)
	{
		Recolor.A(A.Fill.ForeColor, B, C);
		checked
		{
			C++;
		}
	}

	private static void B(ChartFormat A, int B, ref int C)
	{
		Recolor.A(A.Line.ForeColor, B, C);
		checked
		{
			C++;
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.ColorFormat A, int B, int C)
	{
		int num = C % B;
		A.RGB = clsColors.RGB2Ole(checked((num > 0) ? KH.A.ChartSeriesColors[num - 1] : KH.A.ChartSeriesColors[B - 1]));
	}

	public static void PointsToSource()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = MH.A.Application;
			Chart chart = null;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			XlCalculation calculation = application.Calculation;
			application.Calculation = XlCalculation.xlCalculationManual;
			try
			{
				if (application.ActiveSheet is Worksheet)
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
					chart = application.ActiveChart;
				}
				else
				{
					chart = (Chart)application.ActiveSheet;
				}
				if (chart != null)
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
					B(chart);
				}
				else
				{
					A(application, B);
				}
				A(VH.A(69607));
			}
			catch (NotImplementedException ex)
			{
				ProjectData.SetProjectError(ex);
				NotImplementedException ex2 = ex;
				Helpers.B();
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(ex4.Message);
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			application.Calculation = calculation;
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			JH.A((object)chart);
			chart = null;
			return;
		}
	}

	private static void B(Chart A)
	{
		int num = 0;
		IEnumerator enumerator = ((IEnumerable)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
		checked
		{
			Range range;
			try
			{
				while (true)
				{
					Series series;
					if (enumerator.MoveNext())
					{
						series = (Series)enumerator.Current;
						string[] array = Helpers.A(series);
						try
						{
							range = ((_Application)A.Application).get_Range((object)array[2], RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
							continue;
						}
						XlChartType chartType = series.ChartType;
						if (chartType <= XlChartType.xl3DLine)
						{
							if (chartType == XlChartType.xlXYScatter)
							{
								goto IL_020b;
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							if (chartType == XlChartType.xlRadar || chartType == XlChartType.xl3DLine)
							{
								goto IL_0160;
							}
						}
						else
						{
							if (chartType > XlChartType.xlArea)
							{
								if (chartType != XlChartType.xlLine)
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
									switch (chartType)
									{
									case XlChartType.xlLineStacked:
									case XlChartType.xlLineStacked100:
									case XlChartType.xlLineMarkers:
									case XlChartType.xlLineMarkersStacked:
									case XlChartType.xlLineMarkersStacked100:
										break;
									case XlChartType.xlXYScatterSmooth:
									case XlChartType.xlXYScatterSmoothNoMarkers:
									case XlChartType.xlXYScatterLines:
									case XlChartType.xlXYScatterLinesNoMarkers:
									case XlChartType.xlRadarMarkers:
										goto IL_020b;
									case XlChartType.xlAreaStacked:
									case XlChartType.xlAreaStacked100:
									case XlChartType.xl3DAreaStacked:
									case XlChartType.xl3DAreaStacked100:
										goto IL_02f9;
									default:
										goto IL_0344;
									}
								}
								goto IL_0160;
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
							if (chartType == XlChartType.xl3DArea || chartType == XlChartType.xlArea)
							{
								goto IL_02f9;
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
						goto IL_0344;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0425;
						}
						continue;
						end_IL_0425:
						break;
					}
					break;
					IL_0344:
					int num2 = Conversions.ToInteger(NewLateBinding.LateGet(series.Points(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
					int num3 = num2;
					for (int i = 1; i <= num3; i++)
					{
						Recolor.A(((Point)series.Points(i)).Format.Fill, ((Range)range.Cells[i, RuntimeHelpers.GetObjectValue(Missing.Value)]).Interior);
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
					num += num2;
					goto IL_03f2;
					IL_03f2:
					if (num <= 1000)
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
						Forms.WarningMessage(VH.A(69656));
						break;
					}
					break;
					IL_020b:
					num2 = Conversions.ToInteger(NewLateBinding.LateGet(series.Points(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
					int num4 = num2;
					for (int j = 1; j <= num4; j++)
					{
						ChartFormat format = ((Point)series.Points(j)).Format;
						Recolor.A(format.Fill, ((Range)range.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)]).Interior);
						Recolor.A(format.Line, ((Range)range.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)]).Interior);
						_ = null;
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
					num += num2;
					goto IL_03f2;
					IL_02f9:
					Recolor.A(series.Format.Fill, ((Range)range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Interior);
					num++;
					goto IL_03f2;
					IL_0160:
					num2 = Conversions.ToInteger(NewLateBinding.LateGet(series.Points(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null));
					int num5 = num2;
					for (int k = 1; k <= num5; k++)
					{
						Recolor.A(((Point)series.Points(k)).Format.Line, ((Range)range.Cells[k, RuntimeHelpers.GetObjectValue(Missing.Value)]).Interior);
					}
					num += num2;
					goto IL_03f2;
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
			range = null;
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.FillFormat A, Interior B)
	{
		Microsoft.Office.Interop.Excel.FillFormat fillFormat = A;
		if (Operators.ConditionalCompareObjectNotEqual(B.ColorIndex, Constants.xlNone, TextCompare: false))
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
			XlPattern xlPattern = (XlPattern)Conversions.ToInteger(B.Pattern);
			if (xlPattern != XlPattern.xlPatternSolid)
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
				if (xlPattern != XlPattern.xlPatternNone)
				{
					if (xlPattern <= XlPattern.xlPatternUp)
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
						if (xlPattern != XlPattern.xlPatternVertical)
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
							if (xlPattern != XlPattern.xlPatternUp)
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
							}
							else
							{
								fillFormat.Patterned(MsoPatternType.msoPatternDarkUpwardDiagonal);
							}
						}
						else
						{
							fillFormat.Patterned(MsoPatternType.msoPatternDarkVertical);
						}
					}
					else
					{
						switch (xlPattern)
						{
						case XlPattern.xlPatternHorizontal:
							fillFormat.Patterned(MsoPatternType.msoPatternDarkHorizontal);
							break;
						case XlPattern.xlPatternLightHorizontal:
							fillFormat.Patterned(MsoPatternType.msoPatternLightHorizontal);
							break;
						case XlPattern.xlPatternLightVertical:
							fillFormat.Patterned(MsoPatternType.msoPatternLightVertical);
							break;
						case XlPattern.xlPatternGrid:
							fillFormat.Patterned(MsoPatternType.msoPatternSmallGrid);
							break;
						case XlPattern.xlPatternChecker:
							fillFormat.Patterned(MsoPatternType.msoPatternSmallCheckerBoard);
							break;
						case XlPattern.xlPatternCrissCross:
							fillFormat.Patterned(MsoPatternType.msoPattern30Percent);
							break;
						case XlPattern.xlPatternSemiGray75:
							fillFormat.Patterned(MsoPatternType.msoPatternTrellis);
							break;
						case XlPattern.xlPatternDown:
							fillFormat.Patterned(MsoPatternType.msoPatternDarkDownwardDiagonal);
							break;
						case XlPattern.xlPatternLightDown:
							fillFormat.Patterned(MsoPatternType.msoPatternLightDownwardDiagonal);
							break;
						case XlPattern.xlPatternLightUp:
							fillFormat.Patterned(MsoPatternType.msoPatternLightUpwardDiagonal);
							break;
						case XlPattern.xlPatternGray75:
							fillFormat.Patterned(MsoPatternType.msoPattern70Percent);
							break;
						case XlPattern.xlPatternGray50:
							fillFormat.Patterned(MsoPatternType.msoPattern50Percent);
							break;
						case XlPattern.xlPatternGray25:
							fillFormat.Patterned(MsoPatternType.msoPattern25Percent);
							break;
						case XlPattern.xlPatternGray16:
							fillFormat.Patterned(MsoPatternType.msoPattern20Percent);
							break;
						case XlPattern.xlPatternGray8:
							fillFormat.Patterned(MsoPatternType.msoPattern10Percent);
							break;
						}
					}
					fillFormat.ForeColor.RGB = Conversions.ToInteger(B.PatternColor);
					fillFormat.BackColor.RGB = Conversions.ToInteger(B.Color);
					goto IL_0227;
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
			fillFormat.BackColor.RGB = Conversions.ToInteger(B.Color);
			fillFormat.ForeColor.RGB = Conversions.ToInteger(B.Color);
		}
		goto IL_0227;
		IL_0227:
		fillFormat = null;
	}

	private static void A(LineFormat A, Interior B)
	{
		LineFormat lineFormat = A;
		if (Operators.ConditionalCompareObjectNotEqual(B.ColorIndex, Constants.xlNone, TextCompare: false))
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
			lineFormat.ForeColor.RGB = Conversions.ToInteger(B.Color);
		}
		lineFormat = null;
	}

	public static void LabelsToPoints()
	{
		if (!A())
		{
			return;
		}
		Application application = MH.A.Application;
		Chart chart = null;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		XlCalculation calculation = application.Calculation;
		application.Calculation = XlCalculation.xlCalculationManual;
		try
		{
			object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
			if (objectValue is Microsoft.Office.Interop.Excel.DataLabels)
			{
				A((Series)((Microsoft.Office.Interop.Excel.DataLabels)objectValue).Parent);
			}
			else if (objectValue is DataLabel)
			{
				A((DataLabel)objectValue);
			}
			else if (objectValue is Series)
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
				A((Series)objectValue);
			}
			else
			{
				if (application.ActiveSheet is Worksheet)
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
					chart = application.ActiveChart;
				}
				else
				{
					chart = (Chart)application.ActiveSheet;
				}
				if (chart != null)
				{
					C(chart);
				}
				else
				{
					A(application, C);
				}
			}
			A(VH.A(69699));
		}
		catch (NotImplementedException ex)
		{
			ProjectData.SetProjectError(ex);
			NotImplementedException ex2 = ex;
			Helpers.B();
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application.Calculation = calculation;
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		JH.A((object)chart);
		chart = null;
	}

	private static void C(Chart A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Recolor.A((Series)enumerator.Current);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void A(Series A)
	{
		try
		{
			_ = (Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
			return;
		}
		XlChartType chartType = A.ChartType;
		if (chartType <= XlChartType.xlRadar)
		{
			if (chartType == XlChartType.xlXYScatter)
			{
				goto IL_00f2;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (chartType != XlChartType.xlRadar)
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
				goto IL_00f9;
			}
		}
		else if (chartType != XlChartType.xl3DLine)
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
			if (chartType != XlChartType.xlLine)
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
				switch (chartType)
				{
				case XlChartType.xlLineStacked:
				case XlChartType.xlLineStacked100:
				case XlChartType.xlLineMarkers:
				case XlChartType.xlLineMarkersStacked:
				case XlChartType.xlLineMarkersStacked100:
					break;
				case XlChartType.xlXYScatterSmooth:
				case XlChartType.xlXYScatterSmoothNoMarkers:
				case XlChartType.xlXYScatterLines:
				case XlChartType.xlXYScatterLinesNoMarkers:
				case XlChartType.xlRadarMarkers:
					goto IL_00f2;
				default:
					goto IL_00f9;
				}
			}
		}
		B(A);
		return;
		IL_00f2:
		D(A);
		return;
		IL_00f9:
		C(A);
	}

	private static void B(Series A)
	{
		if (A.Format.Line.Visible != MsoTriState.msoTrue)
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
			((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = A.Format.Line.ForeColor.RGB;
			return;
		}
	}

	private static void C(Series A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)A.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Recolor.A((Point)enumerator.Current);
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
				return;
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

	private static void A(Point A)
	{
		if (!A.HasDataLabel || A.Format.Fill.Visible != MsoTriState.msoTrue)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = A.Format.Fill.ForeColor.RGB;
			return;
		}
	}

	private static void D(Series A)
	{
		IEnumerator enumerator = ((IEnumerable)A.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				B((Point)enumerator.Current);
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
				return;
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

	private static void B(Point A)
	{
		if (A.HasDataLabel)
		{
			A.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = A.MarkerForegroundColor;
		}
	}

	private static void A(DataLabel A)
	{
		Point point = (Point)A.Parent;
		Series series = (Series)point.Parent;
		XlChartType chartType = series.ChartType;
		if (chartType <= XlChartType.xlRadar)
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
			if (chartType == XlChartType.xlXYScatter)
			{
				goto IL_0140;
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
			if (chartType != XlChartType.xlRadar)
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
				goto IL_0148;
			}
		}
		else if (chartType != XlChartType.xl3DLine)
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
			switch (chartType)
			{
			case XlChartType.xlLine:
			case XlChartType.xlLineStacked:
			case XlChartType.xlLineStacked100:
			case XlChartType.xlLineMarkers:
			case XlChartType.xlLineMarkersStacked:
			case XlChartType.xlLineMarkersStacked100:
				break;
			case XlChartType.xlXYScatterSmooth:
			case XlChartType.xlXYScatterSmoothNoMarkers:
			case XlChartType.xlXYScatterLines:
			case XlChartType.xlXYScatterLinesNoMarkers:
			case XlChartType.xlRadarMarkers:
				goto IL_0140;
			default:
				goto IL_0148;
			}
		}
		if (series.Format.Line.Visible == MsoTriState.msoTrue)
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
			A.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = series.Format.Line.ForeColor.RGB;
		}
		goto IL_014e;
		IL_0140:
		B(point);
		goto IL_014e;
		IL_014e:
		point = null;
		series = null;
		return;
		IL_0148:
		Recolor.A(point);
		goto IL_014e;
	}

	private static void A(Application A, Action<Chart> B)
	{
		if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(A.Selection)), VH.A(56245), TextCompare: false) != 0)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = ((IEnumerable)NewLateBinding.LateGet(A.Selection, null, VH.A(56274), new object[0], null, null, null)).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape shape = (Shape)enumerator.Current;
					if (shape.HasChart != MsoTriState.msoTrue)
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
					B(shape.Chart);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
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
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, A);
	}
}
