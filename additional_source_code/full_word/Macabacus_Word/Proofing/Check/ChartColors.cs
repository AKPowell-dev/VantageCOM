using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class ChartColors
{
	public static void ChartArea(object shp, List<int> listColors, Severity sev)
	{
		//IL_00e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d9: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22348), new object[0], null, null, null), null, XC.A(22367), new object[0], null, null, null), null, XC.A(22380), new object[0], null, null, null), null, XC.A(22389), new object[0], null, null, null), null, XC.A(22408), new object[0], null, null, null));
			if (!listColors.Contains(num))
			{
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
					Main.Analysis.Errors.Add(new NonconformingChartAreaFillColor(RuntimeHelpers.GetObjectValue(shp), num, sev));
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
			int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22348), new object[0], null, null, null), null, XC.A(22367), new object[0], null, null, null), null, XC.A(22415), new object[0], null, null, null), null, XC.A(22389), new object[0], null, null, null), null, XC.A(22408), new object[0], null, null, null));
			if (listColors.Contains(num))
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
				Main.Analysis.Errors.Add(new NonconformingChartAreaBorderColor(RuntimeHelpers.GetObjectValue(shp), num, sev));
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

	public static void PlotArea(object shp, List<int> listColors, Severity sev)
	{
		//IL_0126: Unknown result type (might be due to invalid IL or missing references)
		//IL_025b: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), null, XC.A(22367), new object[0], null, null, null), null, XC.A(22380), new object[0], null, null, null), null, XC.A(22389), new object[0], null, null, null), null, XC.A(22408), new object[0], null, null, null));
			if (!listColors.Contains(num))
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
					Main.Analysis.Errors.Add(new NonconformingPlotAreaFillColor(RuntimeHelpers.GetObjectValue(shp), num, (PlotArea)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), sev));
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
			int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), null, XC.A(22367), new object[0], null, null, null), null, XC.A(22415), new object[0], null, null, null), null, XC.A(22389), new object[0], null, null, null), null, XC.A(22408), new object[0], null, null, null));
			if (listColors.Contains(num))
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
				Main.Analysis.Errors.Add(new NonconformingPlotAreaBorderColor(RuntimeHelpers.GetObjectValue(shp), num, (PlotArea)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), sev));
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

	public static void Axes(object shp, List<int> listColors, Severity sev)
	{
		//IL_05e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0581: Unknown result type (might be due to invalid IL or missing references)
		//IL_031d: Unknown result type (might be due to invalid IL or missing references)
		//IL_027a: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_040c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0507: Unknown result type (might be due to invalid IL or missing references)
		//IL_0480: Unknown result type (might be due to invalid IL or missing references)
		checked
		{
			try
			{
				Microsoft.Office.Interop.Word.XlAxisType[] array = new Microsoft.Office.Interop.Word.XlAxisType[2]
				{
					Microsoft.Office.Interop.Word.XlAxisType.xlValue,
					Microsoft.Office.Interop.Word.XlAxisType.xlCategory
				};
				for (int i = 0; i < array.Length; i++)
				{
					Microsoft.Office.Interop.Word.XlAxisType xlAxisType = array[i];
					XlAxisGroup[] array2 = new XlAxisGroup[2]
					{
						XlAxisGroup.xlPrimary,
						XlAxisGroup.xlSecondary
					};
					for (int j = 0; j < array2.Length; j++)
					{
						XlAxisGroup xlAxisGroup = array2[j];
						object[] array3;
						bool[] array4;
						object value = NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(13913), array3 = new object[2] { xlAxisType, xlAxisGroup }, null, null, array4 = new bool[2] { true, true });
						if (array4[0])
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
							xlAxisType = (Microsoft.Office.Interop.Word.XlAxisType)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(Microsoft.Office.Interop.Word.XlAxisType));
						}
						if (array4[1])
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
							xlAxisGroup = (XlAxisGroup)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(XlAxisGroup));
						}
						if (!Conversions.ToBoolean(value))
						{
							continue;
						}
						object instance = NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null);
						string memberName = XC.A(14441);
						object[] obj = new object[2] { xlAxisType, xlAxisGroup };
						array3 = obj;
						bool[] obj2 = new bool[2] { true, true };
						array4 = obj2;
						object obj3 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
						if (array4[0])
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
							xlAxisType = (Microsoft.Office.Interop.Word.XlAxisType)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(Microsoft.Office.Interop.Word.XlAxisType));
						}
						if (array4[1])
						{
							xlAxisGroup = (XlAxisGroup)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(XlAxisGroup));
						}
						Axis axis = (Axis)obj3;
						Axis axis2 = axis;
						int rGB;
						if (axis2.HasMajorGridlines)
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
							rGB = axis2.MajorGridlines.Format.Line.ForeColor.RGB;
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
								Main.Analysis.Errors.Add(new NonconformingGridlinesColor(RuntimeHelpers.GetObjectValue(shp), rGB, (PlotArea)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), sev));
							}
						}
						if (axis2.HasMinorGridlines)
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
							rGB = axis2.MinorGridlines.Format.Line.ForeColor.RGB;
							if (Color.ColorNotInPalette(rGB, listColors))
							{
								Main.Analysis.Errors.Add(new NonconformingGridlinesColor(RuntimeHelpers.GetObjectValue(shp), rGB, (PlotArea)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22424), new object[0], null, null, null), sev));
							}
						}
						if (axis2.HasTitle)
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
							AxisTitle axisTitle = axis2.AxisTitle;
							AxisTitle axisTitle2 = axis2.AxisTitle;
							rGB = axisTitle2.Format.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor.RGB;
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
								Main.Analysis.Errors.Add(new NonconformingAxisTitleFontColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axisTitle, sev));
							}
							Microsoft.Office.Interop.Word.FillFormat fill = axisTitle2.Format.Fill;
							if (fill.Visible == MsoTriState.msoTrue)
							{
								rGB = fill.ForeColor.RGB;
								if (Color.ColorNotInPalette(rGB, listColors))
								{
									Main.Analysis.Errors.Add(new NonconformingAxisTitleFillColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axisTitle, sev));
								}
							}
							fill = null;
							Microsoft.Office.Interop.Word.LineFormat line = axisTitle2.Format.Line;
							if (line.Weight > 0f && line.Visible == MsoTriState.msoTrue)
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
								rGB = line.ForeColor.RGB;
								if (Color.ColorNotInPalette(rGB, listColors))
								{
									Main.Analysis.Errors.Add(new NonconformingAxisTitleBorderColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axisTitle, sev));
								}
							}
							line = null;
							_ = null;
							axisTitle = null;
						}
						try
						{
							Microsoft.Office.Interop.Word.LineFormat line2 = axis2.Format.Line;
							if (line2.Weight > 0f && line2.Visible == MsoTriState.msoTrue)
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
								rGB = line2.ForeColor.RGB;
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
									Main.Analysis.Errors.Add(new NonconformingAxisBorderColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axis, sev));
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
							Microsoft.Office.Interop.Word.FillFormat fill2 = axis2.Format.Fill;
							if (fill2.Visible == MsoTriState.msoTrue)
							{
								rGB = fill2.ForeColor.RGB;
								if (Color.ColorNotInPalette(rGB, listColors))
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
									Main.Analysis.Errors.Add(new NonconformingAxisFillColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axis, sev));
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
						rGB = Conversions.ToInteger(axis2.TickLabels.Font.Color);
						if (Color.ColorNotInPalette(rGB, listColors))
						{
							Main.Analysis.Errors.Add(new NonconformingAxisFontColor(RuntimeHelpers.GetObjectValue(shp), rGB, xlAxisGroup, axis, sev));
						}
						axis2 = null;
						axis = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0601;
						}
						continue;
						end_IL_0601:
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
		}
	}

	public static void Series(object shp, List<int> listColors, Severity sev)
	{
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cc: Unknown result type (might be due to invalid IL or missing references)
		IEnumerator enumerator2 = default(IEnumerator);
		foreach (IMsoSeries item in (IEnumerable)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22441), new object[0], null, null, null))
		{
			try
			{
				Microsoft.Office.Core.ColorFormat foreColor = item.Format.Fill.ForeColor;
				if (!listColors.Contains(foreColor.RGB))
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
					Main.Analysis.Errors.Add(new NonconformingSeriesColor(RuntimeHelpers.GetObjectValue(shp), foreColor.RGB, item, sev));
				}
				foreColor = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				int num = Conversions.ToInteger(((IMsoDataLabels)item.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color);
				if (listColors.Contains(num))
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
					List<IMsoDataLabel> list = new List<IMsoDataLabel>();
					try
					{
						enumerator2 = ((IEnumerable)item.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator2.MoveNext())
						{
							IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
							if (listColors.Contains(Conversions.ToInteger(msoDataLabel.Font.Color)))
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
							list.Add(msoDataLabel);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0188;
							}
							continue;
							end_IL_0188:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					Main.Analysis.Errors.Add(new NonconformingLabelsColor(RuntimeHelpers.GetObjectValue(shp), num, item, list, sev));
					list = null;
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void Legend(object shp, List<int> listColors, Severity sev)
	{
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_016d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!Conversions.ToBoolean(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22474), new object[0], null, null, null)))
			{
				return;
			}
			Legend legend = (Legend)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22493), new object[0], null, null, null);
			Microsoft.Office.Core.ColorFormat foreColor = legend.Format.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor;
			if (!listColors.Contains(foreColor.RGB))
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
				Main.Analysis.Errors.Add(new NonconformingLegendFontColor(RuntimeHelpers.GetObjectValue(shp), foreColor.RGB, legend, sev));
			}
			foreColor = null;
			Microsoft.Office.Interop.Word.FillFormat fill = legend.Format.Fill;
			if (fill.Visible == MsoTriState.msoTrue)
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
				Microsoft.Office.Interop.Word.ColorFormat foreColor2 = fill.ForeColor;
				if (!listColors.Contains(foreColor2.RGB))
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
					Main.Analysis.Errors.Add(new NonconformingLegendFillColor(RuntimeHelpers.GetObjectValue(shp), foreColor2.RGB, legend, sev));
				}
				foreColor2 = null;
			}
			fill = null;
			Microsoft.Office.Interop.Word.LineFormat line = legend.Format.Line;
			if (line.Visible == MsoTriState.msoTrue)
			{
				Microsoft.Office.Interop.Word.ColorFormat foreColor3 = line.ForeColor;
				if (!listColors.Contains(foreColor3.RGB))
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
					Main.Analysis.Errors.Add(new NonconformingLegendBorderColor(RuntimeHelpers.GetObjectValue(shp), foreColor3.RGB, legend, sev));
				}
				foreColor3 = null;
			}
			line = null;
			legend = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Title(object shp, List<int> listColors, Severity sev)
	{
		//IL_0106: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!Conversions.ToBoolean(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22506), new object[0], null, null, null)))
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
				ChartTitle chartTitle = (ChartTitle)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22523), new object[0], null, null, null);
				_ = chartTitle.Format;
				Microsoft.Office.Core.ColorFormat foreColor = chartTitle.Format.TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor;
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
					Main.Analysis.Errors.Add(new NonconformingTitleFontColor(RuntimeHelpers.GetObjectValue(shp), foreColor.RGB, chartTitle, sev));
				}
				foreColor = null;
				Microsoft.Office.Interop.Word.FillFormat fill = chartTitle.Format.Fill;
				if (fill.Visible == MsoTriState.msoTrue)
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
					Microsoft.Office.Interop.Word.ColorFormat foreColor2 = fill.ForeColor;
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
						Main.Analysis.Errors.Add(new NonconformingTitleFillColor(RuntimeHelpers.GetObjectValue(shp), foreColor2.RGB, chartTitle, sev));
					}
					foreColor2 = null;
				}
				fill = null;
				Microsoft.Office.Interop.Word.LineFormat line = chartTitle.Format.Line;
				if (line.Visible == MsoTriState.msoTrue)
				{
					Microsoft.Office.Interop.Word.ColorFormat foreColor3 = line.ForeColor;
					if (!listColors.Contains(foreColor3.RGB))
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
						Main.Analysis.Errors.Add(new NonconformingTitleBorderColor(RuntimeHelpers.GetObjectValue(shp), foreColor3.RGB, chartTitle, sev));
					}
					foreColor3 = null;
				}
				line = null;
				_ = null;
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

	public static void DataTable(object shp, List<int> listColors, Severity sev)
	{
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_015a: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d2: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!Conversions.ToBoolean(NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22544), new object[0], null, null, null)))
			{
				return;
			}
			DataTable dataTable = (DataTable)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22569), new object[0], null, null, null);
			_ = dataTable.Format;
			Microsoft.Office.Interop.Word.ChartFont font = dataTable.Font;
			if (!listColors.Contains(Conversions.ToInteger(font.Color)))
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
				Main.Analysis.Errors.Add(new NonconformingDataTableFontColor(RuntimeHelpers.GetObjectValue(shp), Conversions.ToInteger(font.Color), dataTable, sev));
			}
			font = null;
			Microsoft.Office.Interop.Word.FillFormat fill = dataTable.Format.Fill;
			if (fill.Visible == MsoTriState.msoTrue)
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
				Microsoft.Office.Interop.Word.ColorFormat foreColor = fill.ForeColor;
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
					Main.Analysis.Errors.Add(new NonconformingDataTableFillColor(RuntimeHelpers.GetObjectValue(shp), foreColor.RGB, dataTable, sev));
				}
				foreColor = null;
			}
			fill = null;
			Microsoft.Office.Interop.Word.LineFormat line = dataTable.Format.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				Microsoft.Office.Interop.Word.ColorFormat foreColor2 = line.ForeColor;
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
					Main.Analysis.Errors.Add(new NonconformingDataTableBorderColor(RuntimeHelpers.GetObjectValue(shp), foreColor2.RGB, dataTable, sev));
				}
				foreColor2 = null;
			}
			line = null;
			_ = null;
			dataTable = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Groups(object shp, List<int> listColors, Severity sev)
	{
		//IL_021c: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bf: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int count = ((ChartGroups)NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22588), new object[0], null, null, null)).Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				object[] array;
				bool[] array2;
				object obj = NewLateBinding.LateGet(NewLateBinding.LateGet(shp, null, XC.A(13902), new object[0], null, null, null), null, XC.A(22588), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				ChartGroup chartGroup = (ChartGroup)obj;
				ChartGroup chartGroup2 = chartGroup;
				if (chartGroup2.HasUpDownBars)
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
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					Microsoft.Office.Interop.Word.ColorFormat foreColor = chartGroup2.HiLoLines.Format.Line.ForeColor;
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
						Main.Analysis.Errors.Add(new NonconformingHiLoLinesColor(RuntimeHelpers.GetObjectValue(shp), foreColor.RGB, chartGroup, sev));
					}
					foreColor = null;
				}
				if (chartGroup2.HasDropLines)
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
					Microsoft.Office.Interop.Word.ColorFormat foreColor2 = chartGroup2.DropLines.Format.Line.ForeColor;
					Main.Analysis.Errors.Add(new NonconformingDropLinesColor(RuntimeHelpers.GetObjectValue(shp), foreColor2.RGB, chartGroup, sev));
					foreColor2 = null;
				}
				chartGroup2 = null;
				chartGroup = null;
			}
			while (true)
			{
				switch (1)
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
