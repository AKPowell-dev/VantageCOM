using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Colors;

public sealed class Index
{
	public static void Font(Range rng, object obj, ref List<Tuple<int, object>> FontColors)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = rng.Paragraphs.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Paragraph paragraph = (Paragraph)enumerator.Current;
					FontColors.Add(A(paragraph.Range.Font.TextColor.RGB, RuntimeHelpers.GetObjectValue(obj)));
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
					return;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Font(Microsoft.Office.Interop.Word.Shape shp, object obj, ref List<Tuple<int, object>> FontColors)
	{
		Microsoft.Office.Interop.Word.Shape shape = shp;
		try
		{
			if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
			{
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
					A(shape.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(obj), ref FontColors);
					break;
				}
			}
			else
			{
				B(shape.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(obj), ref FontColors);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		shape = null;
	}

	public static void Fill(Microsoft.Office.Interop.Word.Shape shp, object obj, ref List<Tuple<int, object>> FillColors)
	{
		A(shp.Fill, RuntimeHelpers.GetObjectValue(obj), ref FillColors);
	}

	public static void Fill(InlineShape shp, object obj, ref List<Tuple<int, object>> FillColors)
	{
		A(shp.Fill, RuntimeHelpers.GetObjectValue(obj), ref FillColors);
	}

	private static void A(Microsoft.Office.Interop.Word.FillFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Interop.Word.FillFormat fillFormat = A;
			if (fillFormat.Visible == MsoTriState.msoTrue)
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
				Index.A(fillFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(fillFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			fillFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Border(Microsoft.Office.Interop.Word.Shape shp, object obj, ref List<Tuple<int, object>> BorderColors)
	{
		A(shp.Line, RuntimeHelpers.GetObjectValue(obj), ref BorderColors);
	}

	public static void Border(InlineShape shp, object obj, ref List<Tuple<int, object>> BorderColors)
	{
		A(shp.Line, RuntimeHelpers.GetObjectValue(obj), ref BorderColors);
	}

	private static void A(Microsoft.Office.Interop.Word.LineFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Interop.Word.LineFormat lineFormat = A;
			if (lineFormat.Visible == MsoTriState.msoTrue)
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
				Index.A(lineFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(lineFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			lineFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Table(Table tbl, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, object>> FontColors, ref List<Tuple<int, object>> FillColors, ref List<Tuple<int, object>> BorderColors)
	{
		Table table = tbl;
		int count = table.Rows.Count;
		int count2 = table.Columns.Count;
		int num = count;
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			int num2 = count2;
			for (int j = 1; j <= num2; j = checked(j + 1))
			{
				Cell cell = table.Cell(i, j);
				if (blnFont)
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
					Font(cell.Range, tbl, ref FontColors);
				}
				if (blnFill)
				{
					try
					{
						FillColors.Add(A((int)cell.Shading.ForegroundPatternColor, tbl));
						FillColors.Add(A((int)cell.Shading.BackgroundPatternColor, tbl));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				if (blnBorder)
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
						BorderColors.Add(A((int)cell.Borders.InsideColor, tbl));
						BorderColors.Add(A((int)cell.Borders.OutsideColor, tbl));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				cell = null;
			}
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			table = null;
			return;
		}
	}

	public static void SmartArt(SmartArt art, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, object>> FontColors, ref List<Tuple<int, object>> FillColors, ref List<Tuple<int, object>> BorderColors)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = art.Nodes.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				if (blnFont)
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
					try
					{
						if (smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								A(smartArtNode.TextFrame2.TextRange, art, ref FontColors);
								break;
							}
						}
						else
						{
							B(smartArtNode.TextFrame2.TextRange, art, ref FontColors);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				try
				{
					enumerator2 = smartArtNode.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator2.Current;
						if (blnFill)
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
								Microsoft.Office.Core.FillFormat fill = shape.Fill;
								if (fill.Visible == MsoTriState.msoTrue)
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
									A(fill.ForeColor, art, ref FillColors);
									A(fill.BackColor, art, ref FillColors);
								}
								fill = null;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						if (!blnBorder)
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
						try
						{
							Microsoft.Office.Core.LineFormat line = shape.Line;
							if (line.Visible == MsoTriState.msoTrue)
							{
								A(line.ForeColor, art, ref BorderColors);
								A(line.BackColor, art, ref BorderColors);
							}
							line = null;
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (5)
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

	public static void Chart(Chart cht, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, object>> FontColors, ref List<Tuple<int, object>> FillColors, ref List<Tuple<int, object>> BorderColors)
	{
		Chart chart = cht;
		if (blnFill)
		{
			B(chart.ChartArea.Format.Fill, cht.ChartArea, ref FillColors);
			_ = null;
			B(chart.PlotArea.Format.Fill, cht.PlotArea, ref FillColors);
			_ = null;
		}
		if (blnBorder)
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
			try
			{
				Microsoft.Office.Interop.Word.LineFormat line = chart.ChartArea.Format.Line;
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
					A(line.ForeColor, cht.ChartArea, ref BorderColors);
				}
				line = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				Microsoft.Office.Interop.Word.LineFormat line2 = chart.PlotArea.Format.Line;
				if (line2.Visible == MsoTriState.msoTrue)
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
					A(line2.ForeColor, cht.PlotArea, ref BorderColors);
				}
				line2 = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		int count = ((ChartGroups)chart.get_ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			ChartGroup chartGroup = (ChartGroup)chart.get_ChartGroups((object)i);
			if (chartGroup.HasUpDownBars)
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
				if (blnFill)
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
					Microsoft.Office.Interop.Word.FillFormat fill = chartGroup.UpBars.Format.Fill;
					A(fill.ForeColor, cht, ref FillColors);
					A(fill.BackColor, cht, ref FillColors);
					_ = null;
					Microsoft.Office.Interop.Word.FillFormat fill2 = chartGroup.DownBars.Format.Fill;
					A(fill2.ForeColor, cht, ref FillColors);
					A(fill2.BackColor, cht, ref FillColors);
					_ = null;
				}
				if (blnBorder)
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
					Microsoft.Office.Interop.Word.LineFormat line3 = chartGroup.UpBars.Format.Line;
					A(line3.ForeColor, cht, ref BorderColors);
					A(line3.BackColor, cht, ref BorderColors);
					_ = null;
					Microsoft.Office.Interop.Word.LineFormat line4 = chartGroup.DownBars.Format.Line;
					A(line4.ForeColor, cht, ref BorderColors);
					A(line4.BackColor, cht, ref BorderColors);
					_ = null;
				}
			}
			if (chartGroup.HasHiLoLines)
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
				if (blnBorder)
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
						Microsoft.Office.Interop.Word.LineFormat line5 = chartGroup.HiLoLines.Format.Line;
						A(line5.ForeColor, cht, ref BorderColors);
						A(line5.BackColor, cht, ref BorderColors);
						_ = null;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (chartGroup.HasDropLines)
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
				if (blnBorder)
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
					try
					{
						Microsoft.Office.Interop.Word.LineFormat line6 = chartGroup.DropLines.Format.Line;
						A(line6.ForeColor, cht, ref BorderColors);
						A(line6.BackColor, cht, ref BorderColors);
						_ = null;
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
			}
			chartGroup = null;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			try
			{
				enumerator = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				while (enumerator.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
					if (blnBorder)
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
						if (msoSeries.HasErrorBars)
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
								A(msoSeries.ErrorBars.Format.Line, cht, ref BorderColors);
							}
							catch (Exception ex9)
							{
								ProjectData.SetProjectError(ex9);
								Exception ex10 = ex9;
								ProjectData.ClearProjectError();
							}
						}
						if (msoSeries.HasLeaderLines)
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
							try
							{
								A(msoSeries.LeaderLines.Format.Line, cht, ref BorderColors);
							}
							catch (Exception ex11)
							{
								ProjectData.SetProjectError(ex11);
								Exception ex12 = ex11;
								ProjectData.ClearProjectError();
							}
						}
						foreach (IMsoTrendline item in (IEnumerable)msoSeries.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value)))
						{
							A(item.Format.Line, cht, ref BorderColors);
						}
					}
					if (blnFill)
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
						A(msoSeries.Format.Fill, cht, ref FillColors);
					}
					if (blnBorder)
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
						A(msoSeries.Format.Line, cht, ref BorderColors);
					}
					foreach (ChartPoint item2 in (IEnumerable)msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value)))
					{
						if (blnFill)
						{
							A(item2.Format.Fill, cht, ref FillColors);
							try
							{
								if (item2.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										if (item2.MarkerBackgroundColor == Base.TRANSPARENT)
										{
											break;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											FillColors.Add(A(item2.MarkerBackgroundColor, cht));
											break;
										}
										break;
									}
								}
							}
							catch (Exception ex13)
							{
								ProjectData.SetProjectError(ex13);
								Exception ex14 = ex13;
								ProjectData.ClearProjectError();
							}
						}
						if (blnBorder)
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
							A(item2.Format.Line, cht, ref BorderColors);
							try
							{
								if (item2.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										if (item2.MarkerForegroundColor == Base.TRANSPARENT)
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
											BorderColors.Add(A(item2.MarkerForegroundColor, cht));
											break;
										}
										break;
									}
								}
							}
							catch (Exception ex15)
							{
								ProjectData.SetProjectError(ex15);
								Exception ex16 = ex15;
								ProjectData.ClearProjectError();
							}
						}
						if (!blnFont)
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
						try
						{
							FontColors.Add(A(Conversions.ToInteger(item2.DataLabel.Font.Color), cht));
						}
						catch (Exception ex17)
						{
							ProjectData.SetProjectError(ex17);
							Exception ex18 = ex17;
							ProjectData.ClearProjectError();
						}
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0662;
					}
					continue;
					end_IL_0662:
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
			Microsoft.Office.Interop.Word.XlAxisType[] array = new Microsoft.Office.Interop.Word.XlAxisType[2]
			{
				Microsoft.Office.Interop.Word.XlAxisType.xlValue,
				Microsoft.Office.Interop.Word.XlAxisType.xlCategory
			};
			foreach (Microsoft.Office.Interop.Word.XlAxisType xlAxisType in array)
			{
				try
				{
					if (!Conversions.ToBoolean(chart.get_HasAxis((object)xlAxisType, RuntimeHelpers.GetObjectValue(Missing.Value))))
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
						Axis axis = (Axis)chart.Axes(xlAxisType);
						if (axis.HasTitle)
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
							AxisTitle axisTitle = axis.AxisTitle;
							if (blnFont)
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
								FontColors.Add(A(Conversions.ToInteger(axisTitle.Font.Color), cht));
							}
							if (blnFill)
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
								B(axisTitle.Format.Fill, cht, ref FillColors);
							}
							if (blnBorder)
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
								B(axisTitle.Format.Line, cht, ref BorderColors);
							}
							axisTitle = null;
						}
						if (blnBorder)
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
							if (axis.HasMajorGridlines)
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
								A(axis.MajorGridlines.Format.Line.ForeColor, cht, ref BorderColors);
							}
							if (axis.HasMinorGridlines)
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
								A(axis.MinorGridlines.Format.Line.ForeColor, cht, ref BorderColors);
							}
							Microsoft.Office.Interop.Word.LineFormat line7 = axis.Format.Line;
							if (line7.Weight > 0f)
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
								if (line7.Visible == MsoTriState.msoTrue)
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
									A(line7.ForeColor, cht, ref BorderColors);
								}
							}
							line7 = null;
						}
						axis = null;
						break;
					}
				}
				catch (Exception ex19)
				{
					ProjectData.SetProjectError(ex19);
					Exception ex20 = ex19;
					ProjectData.ClearProjectError();
				}
			}
			if (chart.HasTitle)
			{
				ChartTitle chartTitle = chart.ChartTitle;
				if (blnFont)
				{
					FontColors.Add(A(Conversions.ToInteger(chartTitle.Font.Color), cht.ChartTitle));
				}
				if (blnFill)
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
					B(chartTitle.Format.Fill, cht.ChartTitle, ref FillColors);
				}
				if (blnBorder)
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
					B(chartTitle.Format.Line, cht.ChartTitle, ref BorderColors);
				}
				chartTitle = null;
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
				Legend legend = chart.Legend;
				if (blnFont)
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
						enumerator4 = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator4.MoveNext())
						{
							LegendEntry legendEntry = (LegendEntry)enumerator4.Current;
							FontColors.Add(A(Conversions.ToInteger(legendEntry.Font.Color), cht.Legend));
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_09bc;
							}
							continue;
							end_IL_09bc:
							break;
						}
					}
					finally
					{
						if (enumerator4 is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator4 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				if (blnFill)
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
					B(legend.Format.Fill, cht.Legend, ref FillColors);
				}
				if (blnBorder)
				{
					B(legend.Format.Line, cht.Legend, ref BorderColors);
				}
				legend = null;
			}
			if (chart.HasDataTable)
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
				DataTable dataTable = chart.DataTable;
				if (blnFont)
				{
					FontColors.Add(A(Conversions.ToInteger(dataTable.Font.Color), cht.DataTable));
				}
				if (blnFill)
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
					B(dataTable.Format.Fill, cht.DataTable, ref FillColors);
				}
				if (blnBorder)
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
					B(dataTable.Format.Line, cht.DataTable, ref BorderColors);
				}
				dataTable = null;
			}
			chart = null;
			return;
		}
	}

	private static void A(TextRange2 A, object B, ref List<Tuple<int, object>> C)
	{
		TextRange2 textRange = A;
		int count = textRange.get_Runs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			Index.B(textRange.get_Runs(i, -1), RuntimeHelpers.GetObjectValue(B), ref C);
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
			textRange = null;
			return;
		}
	}

	private static void B(TextRange2 A, object B, ref List<Tuple<int, object>> C)
	{
		Font2 font = A.Font;
		if (font.Fill.ForeColor.RGB != Base.TRANSPARENT)
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
			C.Add(Index.A(font.Fill.ForeColor.RGB, RuntimeHelpers.GetObjectValue(B)));
		}
		Index.A(font.Fill.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
		Index.A(font.Highlight, RuntimeHelpers.GetObjectValue(B), ref C);
		font = null;
	}

	private static void A(Microsoft.Office.Core.FillFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Core.FillFormat fillFormat = A;
			if (fillFormat.Visible == MsoTriState.msoTrue)
			{
				Index.A(fillFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(fillFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			fillFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B(Microsoft.Office.Interop.Word.FillFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Interop.Word.FillFormat fillFormat = A;
			if (fillFormat.Visible == MsoTriState.msoTrue)
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
				Index.A(fillFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(fillFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			fillFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Core.LineFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Core.LineFormat lineFormat = A;
			if (lineFormat.Visible == MsoTriState.msoTrue)
			{
				Index.A(lineFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(lineFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			lineFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B(Microsoft.Office.Interop.Word.LineFormat A, object B, ref List<Tuple<int, object>> C)
	{
		try
		{
			Microsoft.Office.Interop.Word.LineFormat lineFormat = A;
			if (lineFormat.Visible == MsoTriState.msoTrue)
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
				Index.A(lineFormat.ForeColor, RuntimeHelpers.GetObjectValue(B), ref C);
				Index.A(lineFormat.BackColor, RuntimeHelpers.GetObjectValue(B), ref C);
			}
			lineFormat = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.Word.ColorFormat A, object B, ref List<Tuple<int, object>> C)
	{
		int rGB = A.RGB;
		if (rGB == Base.TRANSPARENT)
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
			if (rGB == 0)
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
				if (A.Type != MsoColorType.msoColorTypeRGB)
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
					break;
				}
			}
			C.Add(Index.A(rGB, RuntimeHelpers.GetObjectValue(B)));
			return;
		}
	}

	private static void A(Microsoft.Office.Core.ColorFormat A, object B, ref List<Tuple<int, object>> C)
	{
		int rGB = A.RGB;
		if (rGB == Base.TRANSPARENT)
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
			if (rGB == 0)
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
				if (A.Type != MsoColorType.msoColorTypeRGB)
				{
					return;
				}
			}
			C.Add(Index.A(rGB, RuntimeHelpers.GetObjectValue(B)));
			return;
		}
	}

	private static Tuple<int, object> A(int A, object B)
	{
		return new Tuple<int, object>(A, RuntimeHelpers.GetObjectValue(B));
	}
}
