using System;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class StatLines
{
	private enum UD
	{
		A,
		B,
		C
	}

	private enum VD
	{
		A,
		B,
		C,
		D
	}

	private struct WD
	{
		public float A;

		public int A;

		public MsoLineDashStyle A;

		public VD A;

		public bool A;

		public bool B;
	}

	public static void Average()
	{
		A(UD.A);
	}

	public static void Median()
	{
		A(UD.B);
	}

	public static void Value()
	{
		A(UD.C);
	}

	private static void A(UD A)
	{
		if (!Licensing.AllowChartAddOnOperation())
		{
			return;
		}
		bool e = default(bool);
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
			Chart chart = Helpers.SelectedChart();
			if (chart != null)
			{
				while (true)
				{
					string text;
					string[] array;
					string f;
					string g;
					bool A2;
					WD h;
					Application application;
					Series series;
					Range range;
					Range d;
					switch (1)
					{
					case 0:
						break;
					default:
						{
							text = "";
							series = ((SeriesCollection)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).Item(1);
							XlChartType chartType = series.ChartType;
							if (chartType <= XlChartType.xlColumnStacked)
							{
								if (chartType != XlChartType.xlLine)
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
									if ((uint)(chartType - 51) > 1u)
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
										goto IL_0121;
									}
								}
							}
							else if ((uint)(chartType - 57) > 1u)
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
								if (chartType != XlChartType.xlLineStacked)
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
									if ((uint)(chartType - 65) > 1u)
									{
										goto IL_0121;
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
							}
							if (series.AxisGroup == XlAxisGroup.xlSecondary)
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
								text = VH.A(68500);
							}
							else if (((Points)series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).Count < 2)
							{
								text = VH.A(72531);
							}
							goto IL_012f;
						}
						IL_0121:
						text = VH.A(72626);
						goto IL_012f;
						IL_012f:
						if (text.Length > 0)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									Forms.WarningMessage(text);
									chart = null;
									series = null;
									return;
								}
							}
						}
						application = MH.A.Application;
						array = Helpers.A(series);
						try
						{
							range = ((_Application)application).get_Range((object)array[2], RuntimeHelpers.GetObjectValue(Missing.Value));
							e = range.Rows.Count == 1;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							range = null;
							ProjectData.ClearProjectError();
						}
						try
						{
							d = ((_Application)application).get_Range((object)array[0], RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							d = null;
							ProjectData.ClearProjectError();
						}
						if (A != UD.A)
						{
							if (A != UD.B)
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
								f = VH.A(41636);
								g = "";
							}
							else
							{
								f = VH.A(72846);
								g = VH.A(72859);
							}
						}
						else
						{
							f = VH.A(72816);
							g = VH.A(72831);
						}
						A2 = false;
						h = StatLines.A(ref A2);
						if (!A2)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
								{
									bool autoExpandListRange = application.AutoCorrect.AutoExpandListRange;
									application.AutoCorrect.AutoExpandListRange = false;
									application.CutCopyMode = (XlCutCopyMode)0;
									application.ScreenUpdating = false;
									try
									{
										XlChartType chartType2 = series.ChartType;
										if (chartType2 <= XlChartType.xlColumnStacked)
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
											if (chartType2 != XlChartType.xlLine)
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
												if ((uint)(chartType2 - 51) > 1u)
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
													goto IL_02f5;
												}
											}
										}
										else if (chartType2 != XlChartType.xlLineStacked)
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
											if ((uint)(chartType2 - 65) > 1u)
											{
												goto IL_02f5;
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
										}
										StatLines.A(chart, A, range, d, e, f, g, h);
										goto end_IL_0284;
										IL_02f5:
										B(chart, A, range, d, e, f, g, h);
										end_IL_0284:;
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										Forms.ErrorMessage(ex6.Message);
										clsReporting.LogException(ex6);
										ProjectData.ClearProjectError();
									}
									application.AutoCorrect.AutoExpandListRange = autoExpandListRange;
									application.ScreenUpdating = true;
									application = null;
									chart = null;
									series = null;
									range = null;
									d = null;
									clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(72872));
									return;
								}
								}
							}
						}
						return;
					}
				}
			}
			Helpers.NoChartMessage();
			return;
		}
	}

	private static void A(Chart A, UD B, Range C, Range D, bool E, string F, string G, WD H)
	{
		_ = A.Application;
		string xValues = StatLines.A(ref C, ref D, E);
		Series series = ((!E) ? StatLines.B(A, B, C, D, F, G) : StatLines.A(A, B, C, D, F, G));
		Series series2 = series;
		series2.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		series2.AxisGroup = XlAxisGroup.xlPrimary;
		series2.XValues = xValues;
		series2.HasErrorBars = true;
		series2.ErrorBars.EndStyle = XlEndStyleCap.xlNoCap;
		series2.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludePlusValues, XlErrorBarType.xlErrorBarTypeFixedValue, 1, RuntimeHelpers.GetObjectValue(Missing.Value));
		series2.ErrorBar(XlErrorBarDirection.xlY, XlErrorBarInclude.xlErrorBarIncludeNone, XlErrorBarType.xlErrorBarTypeFixedValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		StatLines.A(series2.ErrorBars.Format, H);
		if (H.A == MsoLineDashStyle.msoLineSolid)
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
			StatLines.A(series2.Format, H);
		}
		else
		{
			series2.Format.Line.Visible = MsoTriState.msoFalse;
		}
		series2 = null;
		StatLines.A(series, C, D, F, H, E, XlDataLabelPosition.xlLabelPositionAbove);
		series = null;
	}

	private static void B(Chart A, UD B, Range C, Range D, bool E, string F, string G, WD H)
	{
		Application application = A.Application;
		Series series;
		Range xValues;
		if (E)
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
			series = StatLines.A(A, B, C, D, F, G);
			Range range = (Range)C.Rows[C.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
			xValues = ((_Application)application).get_Range((object)range.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2);
			range = null;
		}
		else
		{
			series = StatLines.B(A, B, C, D, F, G);
			Range range2 = (Range)C.Columns[C.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
			xValues = ((_Application)application).get_Range((object)range2.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value));
			range2 = null;
		}
		Series series2 = series;
		series2.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
		series2.Values = VH.A(72891);
		series2.XValues = xValues;
		StatLines.A(series2.Format, H);
		_ = null;
		Axis obj = (Axis)A.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlSecondary);
		obj.MaximumScale = 1.0;
		obj.MinimumScale = 0.0;
		_ = null;
		((_Chart)A).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary, (object)false);
		StatLines.A(series, C, D, F, H, E, XlDataLabelPosition.xlLabelPositionRight);
		application = null;
		series = null;
		xValues = null;
	}

	private static string A(ref Range A, ref Range B, bool C)
	{
		checked
		{
			string result;
			if (C)
			{
				Range obj = (Range)A.Rows[A.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
				obj.get_Offset((object)1, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range = obj.get_Offset((object)1, (object)0);
				int num = Conversions.ToInteger(A.Columns.CountLarge);
				for (int i = 1; i <= num; i++)
				{
					((Range)range.Cells[1, i]).Formula = VH.A(72904) + ((Range)range.Cells[1, 1]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72945) + ((Range)range.Cells[1, i - 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72958);
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
				result = VH.A(48936) + range.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				range = null;
				_ = null;
				if (B != null)
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
					B.get_Offset((object)1, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
					B = ((Range)B.Cells[RuntimeHelpers.GetObjectValue(B.Rows.CountLarge), 1]).get_Offset((object)1, (object)0);
					B.Value2 = VH.A(57387);
				}
				A = A.get_Resize(Operators.AddObject(A.Rows.CountLarge, 1), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			else
			{
				Range obj2 = (Range)A.Columns[A.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
				obj2.get_Offset((object)0, (object)1).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
				Range range2 = obj2.get_Offset((object)0, (object)1);
				int num2 = Conversions.ToInteger(A.Rows.CountLarge);
				for (int j = 1; j <= num2; j++)
				{
					((Range)range2.Cells[j, 1]).Formula = VH.A(72965) + ((Range)range2.Cells[1, 1]).get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72945) + ((Range)range2.Cells[j - 1, 1]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(72958);
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
				result = VH.A(48936) + range2.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				range2 = null;
				_ = null;
				if (B != null)
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
					B.get_Offset((object)0, (object)1).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
					B = ((Range)B.Cells[RuntimeHelpers.GetObjectValue(B.Columns.CountLarge), 1]).get_Offset((object)0, (object)1);
					B.Value2 = VH.A(57387);
				}
				A = A.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.AddObject(A.Columns.CountLarge, 1));
			}
			return result;
		}
	}

	private static Series A(Chart A, UD B, Range C, Range D, string E, string F)
	{
		Application application = A.Application;
		Range range = (Range)C.Rows[C.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
		range.get_Offset((object)1, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range2 = range.get_Offset((object)1, (object)0);
		if (B != UD.C)
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
			range2.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(48936) + F + VH.A(39848), NewLateBinding.LateGet(C.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 1, 1 }, null, null, null)), VH.A(39904));
		}
		else
		{
			range2.Value2 = StatLines.A(application.WorksheetFunction.Median(RuntimeHelpers.GetObjectValue(C.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
		}
		range2 = null;
		Series series;
		if (D != null)
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
			D.get_Offset((object)1, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
			D = ((Range)D.Cells[D.Rows.Count, 1]).get_Offset((object)1, (object)0);
			D.Value2 = E;
			series = (Series)NewLateBinding.LateGet(A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				((_Application)application).get_Range((object)(D.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + StatLines.A() + range.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
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
			series = (Series)NewLateBinding.LateGet(A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				((_Application)application).get_Range((object)range.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
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
			series.Name = E;
		}
		range = null;
		application = null;
		return series;
	}

	private static Series B(Chart A, UD B, Range C, Range D, string E, string F)
	{
		Application application = A.Application;
		Range range = (Range)C.Columns[C.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
		range.get_Offset((object)0, (object)1).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range2 = range.get_Offset((object)0, (object)1);
		if (B != UD.C)
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
			range2.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(48936) + F + VH.A(39848), NewLateBinding.LateGet(C.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 1, 1 }, null, null, null)), VH.A(39904));
		}
		else
		{
			range2.Value2 = StatLines.A(application.WorksheetFunction.Median(RuntimeHelpers.GetObjectValue(C.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
		}
		range2 = null;
		Series series;
		if (D != null)
		{
			D.get_Offset((object)0, (object)1).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
			D = ((Range)D.Cells[1, D.Columns.Count]).get_Offset((object)0, (object)1);
			D.Value2 = E;
			series = (Series)NewLateBinding.LateGet(A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				((_Application)application).get_Range((object)(D.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + StatLines.A() + range.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))), RuntimeHelpers.GetObjectValue(Missing.Value)),
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
			series = (Series)NewLateBinding.LateGet(A.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				((_Application)application).get_Range((object)range.get_Offset((object)0, (object)1).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)),
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
			series.Name = E;
		}
		range = null;
		application = null;
		return series;
	}

	private static void A(Series A, Range B, Range C, string D, WD E, bool F, XlDataLabelPosition G)
	{
		if (E.A == VD.A)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.HasDataLabels = false;
					return;
				}
			}
		}
		A.HasDataLabels = true;
		Range b;
		try
		{
			b = ((!F) ? StatLines.B(A, B, C, D, G, E.A) : StatLines.A(A, B, C, D, G, E.A));
			StatLines.A(A, b, G, E);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(VH.A(72994));
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		b = null;
	}

	private static Range A(Series A, Range B, Range C, string D, XlDataLabelPosition E, VD F)
	{
		Range obj = (Range)B.Rows[B.Rows.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
		obj.get_Offset((object)2, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range = obj.get_Offset((object)2, (object)0);
		if (E == XlDataLabelPosition.xlLabelPositionAbove)
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
			StatLines.A((Range)range.Cells[RuntimeHelpers.GetObjectValue(range.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], F, D);
		}
		else
		{
			StatLines.A((Range)range.Cells[2, RuntimeHelpers.GetObjectValue(Missing.Value)], F, D);
		}
		Range cells = range.Cells;
		range.Font.Color = 0;
		range = null;
		_ = null;
		if (C != null)
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
			C.get_Offset((object)2, (object)0).Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
			C = ((Range)C.Cells[RuntimeHelpers.GetObjectValue(C.Rows.CountLarge), 1]).get_Offset((object)2, (object)0);
			C.Value2 = VH.A(49999);
		}
		return cells;
	}

	private static Range B(Series A, Range B, Range C, string D, XlDataLabelPosition E, VD F)
	{
		Range obj = (Range)B.Columns[B.Columns.Count, RuntimeHelpers.GetObjectValue(Missing.Value)];
		obj.get_Offset((object)0, (object)2).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range = obj.get_Offset((object)0, (object)2);
		if (E == XlDataLabelPosition.xlLabelPositionAbove)
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
			StatLines.B((Range)range.Cells[RuntimeHelpers.GetObjectValue(range.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], F, D);
		}
		else
		{
			StatLines.B((Range)range.Cells[2, RuntimeHelpers.GetObjectValue(Missing.Value)], F, D);
		}
		Range cells = range.Cells;
		range.Font.Color = 0;
		range = null;
		_ = null;
		if (C != null)
		{
			C.get_Offset((object)0, (object)2).Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
			C = ((Range)C.Cells[1, RuntimeHelpers.GetObjectValue(C.Columns.CountLarge)]).get_Offset((object)0, (object)2);
			C.Value2 = VH.A(49999);
		}
		return cells;
	}

	private static void A(Range A, VD B, string C)
	{
		StatLines.A(A, A.get_Offset((object)(-1), (object)0), B, C);
	}

	private static void B(Range A, VD B, string C)
	{
		StatLines.A(A, A.get_Offset((object)0, (object)(-1)), B, C);
	}

	private static void A(Range A, Range B, VD C, string D)
	{
		Range range = A;
		switch (C)
		{
		case VD.B:
			try
			{
				range.Formula = VH.A(73111) + D + VH.A(73116) + B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(73137) + B.NumberFormat.ToString().Replace(VH.A(39830), VH.A(73142)) + VH.A(73151);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				range.Value2 = D;
				ProjectData.ClearProjectError();
			}
			break;
		case VD.D:
			range.Formula = VH.A(48936) + B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			break;
		case VD.C:
			range.Value2 = D;
			break;
		}
		range = null;
	}

	private static void A(Series A, Range B, XlDataLabelPosition C, WD D)
	{
		A.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		A.HasLeaderLines = false;
		Application application = A.Application;
		if (Conversion.Val(application.Version) < 15.0)
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
			if (C == XlDataLabelPosition.xlLabelPositionRight || C == XlDataLabelPosition.xlLabelPositionLeft)
			{
				application.ScreenUpdating = true;
				application.ScreenUpdating = false;
			}
		}
		application = null;
		DataLabels.LinkToRange((Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), B);
		Microsoft.Office.Interop.Excel.DataLabels obj = (Microsoft.Office.Interop.Excel.DataLabels)A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		Font2 font = obj.Format.TextFrame2.TextRange.Font;
		if (D.A)
		{
			font.Fill.ForeColor.RGB = D.A;
		}
		if (D.B)
		{
			font.Bold = MsoTriState.msoTrue;
		}
		font = null;
		obj.Position = C;
		_ = null;
		_ = null;
	}

	private static string A()
	{
		return CultureInfo.CurrentCulture.TextInfo.ListSeparator;
	}

	private static void A(ChartFormat A, WD B)
	{
		LineFormat line = A.Line;
		line.Weight = B.A;
		line.ForeColor.RGB = B.A;
		line.DashStyle = B.A;
		_ = null;
	}

	private static double A(double A)
	{
		if (double.TryParse(Forms.InputBox(VH.A(73156), VH.A(73177), A.ToString()), out var result))
		{
			return result;
		}
		return StatLines.A(A);
	}

	private static WD A(ref bool A)
	{
		WD wD = default(WD);
		try
		{
			XmlDocument A2 = KH.A.SettingsXml;
			wD = StatLines.A(ref A2);
			wpfStatLine wpfStatLine2 = new wpfStatLine();
			System.Drawing.Color color = ColorTranslator.FromOle(wD.A);
			wpfStatLine2.btnColor.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
			wpfStatLine2.numWeight.Value = wD.A;
			wpfStatLine2.cbxStyle.SelectedIndex = B(wD.A);
			wpfStatLine2.cbxLabel.SelectedIndex = (int)wD.A;
			wpfStatLine2.chkColor.IsChecked = wD.A;
			wpfStatLine2.chkBold.IsChecked = wD.B;
			wpfStatLine2.ShowDialog();
			if (wpfStatLine2.DialogResult.HasValue)
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
				if (wpfStatLine2.DialogResult.Value)
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
					System.Windows.Media.Color color2 = ((SolidColorBrush)wpfStatLine2.btnColor.Foreground).Color;
					wD = new WD
					{
						A = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(color2.R, color2.G, color2.B)),
						A = (float)wpfStatLine2.numWeight.Value.Value,
						A = StatLines.A(wpfStatLine2.cbxStyle.SelectedIndex),
						A = (VD)wpfStatLine2.cbxLabel.SelectedIndex,
						A = wpfStatLine2.chkColor.IsChecked.Value,
						B = wpfStatLine2.chkBold.IsChecked.Value
					};
					StatLines.A(wD, ref A2);
					goto IL_0217;
				}
			}
			A = true;
			goto IL_0217;
			IL_0217:
			wpfStatLine2 = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(69125));
			wD.A = 1.5f;
			wD.A = 49407;
			wD.A = MsoLineDashStyle.msoLineSolid;
			wD.A = VD.B;
			wD.A = true;
			wD.B = true;
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlDocument A2 = null;
		}
		return wD;
	}

	private static WD A(ref XmlDocument A)
	{
		WD result = default(WD);
		XmlNode xmlNode = A.DocumentElement.SelectSingleNode(VH.A(1422));
		result.A = clsColors.RGB2Ole(xmlNode.SelectSingleNode(VH.A(55331)).InnerText);
		result.A = float.Parse(xmlNode.SelectSingleNode(VH.A(73254)).InnerText, CultureInfo.InvariantCulture);
		result.A = (MsoLineDashStyle)Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(73267)).InnerText);
		result.A = (VD)Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(73286)).InnerText);
		result.A = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(73305)).InnerText);
		result.B = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(55289)).InnerText);
		xmlNode = null;
		return result;
	}

	private static void A(WD A, ref XmlDocument B)
	{
		XmlNode xmlNode = B.DocumentElement.SelectSingleNode(VH.A(1422));
		xmlNode.SelectSingleNode(VH.A(55331)).InnerText = clsColors.Color2RGB(ColorTranslator.FromOle(A.A));
		xmlNode.SelectSingleNode(VH.A(73254)).InnerText = A.A.ToString(CultureInfo.InvariantCulture);
		XmlNode xmlNode2 = xmlNode.SelectSingleNode(VH.A(73267));
		int a = (int)A.A;
		xmlNode2.InnerText = a.ToString();
		XmlNode xmlNode3 = xmlNode.SelectSingleNode(VH.A(73286));
		a = (int)A.A;
		xmlNode3.InnerText = a.ToString();
		xmlNode.SelectSingleNode(VH.A(73305)).InnerText = A.A.ToString();
		xmlNode.SelectSingleNode(VH.A(55289)).InnerText = A.B.ToString();
		_ = null;
		KH.A.SaveSettings(B);
	}

	private static MsoLineDashStyle A(int A)
	{
		MsoLineDashStyle msoLineDashStyle = default(MsoLineDashStyle);
		return A switch
		{
			0 => MsoLineDashStyle.msoLineSolid, 
			1 => MsoLineDashStyle.msoLineSysDot, 
			2 => MsoLineDashStyle.msoLineSysDash, 
			3 => MsoLineDashStyle.msoLineDash, 
			4 => MsoLineDashStyle.msoLineDashDot, 
			5 => MsoLineDashStyle.msoLineLongDash, 
			6 => MsoLineDashStyle.msoLineLongDashDot, 
			7 => MsoLineDashStyle.msoLineLongDashDotDot, 
			_ => msoLineDashStyle, 
		};
	}

	private static int B(MsoLineDashStyle A)
	{
		int num = default(int);
		return A switch
		{
			MsoLineDashStyle.msoLineSolid => 0, 
			MsoLineDashStyle.msoLineSysDot => 1, 
			MsoLineDashStyle.msoLineSysDash => 2, 
			MsoLineDashStyle.msoLineDash => 3, 
			MsoLineDashStyle.msoLineDashDot => 4, 
			MsoLineDashStyle.msoLineLongDash => 5, 
			MsoLineDashStyle.msoLineLongDashDot => 6, 
			MsoLineDashStyle.msoLineLongDashDotDot => 7, 
			_ => num, 
		};
	}
}
