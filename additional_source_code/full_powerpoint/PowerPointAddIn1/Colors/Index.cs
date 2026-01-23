using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;
using PowerPointAddIn1.DeckCheck.Reformat;

namespace PowerPointAddIn1.Colors;

public sealed class Index
{
	public static void Font(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> FontColors)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
		try
		{
			if (shape.HasTextFrame == MsoTriState.msoTrue)
			{
				if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
				{
					A(shape.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref FontColors);
				}
				else
				{
					TextRange(shape.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref FontColors);
				}
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					BulletFormat2 bullet;
					for (enumerator = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator(); enumerator.MoveNext(); bullet = null)
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						bullet = textRange.ParagraphFormat.Bullet;
						if (bullet.Type != MsoBulletType.msoBulletNumbered)
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
							if (bullet.Type != MsoBulletType.msoBulletUnnumbered)
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
						}
						if (bullet.UseTextColor != MsoTriState.msoFalse)
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
						Microsoft.Office.Core.FillFormat fill = bullet.Font.Fill;
						AddColorIfNotTransparent(fill.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, textRange.ParagraphFormat.Bullet, ref FontColors);
						if (fill.BackColor.RGB != fill.ForeColor.RGB)
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
							AddColorIfNotTransparent(fill.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, textRange.ParagraphFormat.Bullet, ref FontColors);
						}
						fill = null;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0194;
						}
						continue;
						end_IL_0194:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		shape = null;
	}

	public static void Fill(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> FillColors)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = shp.Fill;
			if (fill.Visible == MsoTriState.msoTrue)
			{
				AddColorIfNotTransparent(fill.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref FillColors);
				AddColorIfNotTransparent(fill.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref FillColors);
			}
			fill = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Border(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> BorderColors)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = shp.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				AddColorIfNotTransparent(line.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref BorderColors);
				AddColorIfNotTransparent(line.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref BorderColors);
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Table(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, Table tbl, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, IndexedObject>> FontColors, ref List<Tuple<int, IndexedObject>> FillColors, ref List<Tuple<int, IndexedObject>> BorderColors)
	{
		Table table = tbl;
		int count = table.Rows.Count;
		int count2 = table.Columns.Count;
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				int num2 = count2;
				for (int j = 1; j <= num2; j++)
				{
					Cell cell = table.Cell(i, j);
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						Font(RuntimeHelpers.GetObjectValue(parent), cell.Shape, tbl.Cell(i, j), ref FontColors);
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
						Fill(RuntimeHelpers.GetObjectValue(parent), cell.Shape, tbl.Cell(i, j), ref FillColors);
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
						Border(RuntimeHelpers.GetObjectValue(parent), cell.Shape, tbl.Cell(i, j), ref BorderColors);
					}
					cell = null;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				table = null;
				return;
			}
		}
	}

	public static void SmartArt(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, SmartArt art, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, IndexedObject>> FontColors, ref List<Tuple<int, IndexedObject>> FillColors, ref List<Tuple<int, IndexedObject>> BorderColors)
	{
		IEnumerator enumerator = art.AllNodes.GetEnumerator();
		try
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
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
								A(smartArtNode.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(parent), shp, art, ref FontColors);
								break;
							}
						}
						else
						{
							TextRange(smartArtNode.TextFrame2.TextRange, RuntimeHelpers.GetObjectValue(parent), shp, art, ref FontColors);
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
								switch (2)
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
									AddColorIfNotTransparent(fill.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, art, ref FillColors);
									AddColorIfNotTransparent(fill.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, art, ref FillColors);
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
							switch (4)
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
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									break;
								}
								AddColorIfNotTransparent(line.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, art, ref BorderColors);
								AddColorIfNotTransparent(line.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, art, ref BorderColors);
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
							switch (1)
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
					goto end_IL_01db;
				}
				continue;
				end_IL_01db:
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
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shp.Fill;
			if (fill2.Visible == MsoTriState.msoTrue)
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
				AddColorIfNotTransparent(fill2.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, shp, ref FillColors);
				AddColorIfNotTransparent(fill2.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, shp, ref FillColors);
			}
			fill2 = null;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line2 = shp.Line;
			if (line2.Visible == MsoTriState.msoTrue)
			{
				AddColorIfNotTransparent(line2.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, shp, ref BorderColors);
				AddColorIfNotTransparent(line2.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, shp, ref BorderColors);
			}
			line2 = null;
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			ProjectData.ClearProjectError();
		}
	}

	public static void Chart(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, Chart cht, bool blnFont, bool blnFill, bool blnBorder, ref List<Tuple<int, IndexedObject>> FontColors, ref List<Tuple<int, IndexedObject>> FillColors, ref List<Tuple<int, IndexedObject>> BorderColors)
	{
		Chart chart = cht;
		if (blnFill)
		{
			if (!(chart.ChartArea.Format.Fill.Transparency < 0f))
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
				FillFormat(chart.ChartArea.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, cht.ChartArea, ref FillColors);
			}
			if (!(chart.PlotArea.Format.Fill.Transparency < 0f))
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
				FillFormat(chart.PlotArea.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, cht.PlotArea, ref FillColors);
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
			try
			{
				Microsoft.Office.Interop.PowerPoint.LineFormat line = chart.ChartArea.Format.Line;
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
					if (line.Weight > 0f)
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
						AddColorIfNotTransparent(line.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, cht.ChartArea, ref BorderColors);
					}
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
				Microsoft.Office.Interop.PowerPoint.LineFormat line2 = chart.PlotArea.Format.Line;
				if (line2.Visible == MsoTriState.msoTrue && line2.Weight > 0f)
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
					AddColorIfNotTransparent(line2.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, cht.PlotArea, ref BorderColors);
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
		int count = ((ChartGroups)chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			ChartGroup chartGroup = (ChartGroup)chart.ChartGroups(i);
			ChartGroup chartGroup2 = chartGroup;
			if (chartGroup2.HasUpDownBars)
			{
				if (blnFill)
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
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = chartGroup2.UpBars.Format.Fill;
					AddColorIfNotTransparent(fill.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.UpBars, ref FillColors);
					AddColorIfNotTransparent(fill.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.UpBars, ref FillColors);
					_ = null;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = chartGroup2.DownBars.Format.Fill;
					AddColorIfNotTransparent(fill2.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DownBars, ref FillColors);
					AddColorIfNotTransparent(fill2.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DownBars, ref FillColors);
					_ = null;
				}
				if (blnBorder)
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line3 = chartGroup2.UpBars.Format.Line;
					AddColorIfNotTransparent(line3.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.UpBars, ref BorderColors);
					AddColorIfNotTransparent(line3.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.UpBars, ref BorderColors);
					_ = null;
					Microsoft.Office.Interop.PowerPoint.LineFormat line4 = chartGroup2.DownBars.Format.Line;
					AddColorIfNotTransparent(line4.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DownBars, ref BorderColors);
					AddColorIfNotTransparent(line4.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DownBars, ref BorderColors);
					_ = null;
				}
			}
			if (chartGroup2.HasHiLoLines && blnBorder)
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
				try
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line5 = chartGroup2.HiLoLines.Format.Line;
					AddColorIfNotTransparent(line5.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.HiLoLines, ref BorderColors);
					AddColorIfNotTransparent(line5.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.HiLoLines, ref BorderColors);
					_ = null;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			if (chartGroup2.HasDropLines)
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
					try
					{
						Microsoft.Office.Interop.PowerPoint.LineFormat line6 = chartGroup2.DropLines.Format.Line;
						AddColorIfNotTransparent(line6.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DropLines, ref BorderColors);
						AddColorIfNotTransparent(line6.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, chartGroup.DropLines, ref BorderColors);
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
			if (Charts.HasRadarAxisLabels(chartGroup))
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
					AddColorIntIfNotTransparent(Conversions.ToInteger(chartGroup2.RadarAxisLabels.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, chartGroup2.RadarAxisLabels, ref FontColors, isMarker: false, areRadarLabels: true);
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
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		IEnumerator enumerator6 = default(IEnumerator);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (!clsCharts.SeriesClrsAreUnusable(cht))
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
				bool flag = clsCharts.UsesFormatFillForSeriesClrs(cht);
				try
				{
					enumerator = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator.MoveNext())
					{
						IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
						if (blnBorder)
						{
							if (Charts.ImplsAndHasErrorBars(msoSeries))
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
									LineFormat(msoSeries.ErrorBars.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, msoSeries.ErrorBars, ref BorderColors);
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
									switch (1)
									{
									case 0:
										continue;
									}
									break;
								}
								try
								{
									LineFormat(msoSeries.LeaderLines.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, msoSeries.LeaderLines, ref BorderColors);
								}
								catch (Exception ex11)
								{
									ProjectData.SetProjectError(ex11);
									Exception ex12 = ex11;
									ProjectData.ClearProjectError();
								}
							}
							if (Charts.ImplsTrendLines(msoSeries))
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
									enumerator2 = ((IEnumerable)msoSeries.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator2.MoveNext())
									{
										IMsoTrendline msoTrendline = (IMsoTrendline)enumerator2.Current;
										LineFormat(msoTrendline.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, msoTrendline, ref BorderColors);
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_066d;
										}
										continue;
										end_IL_066d:
										break;
									}
								}
								finally
								{
									if (enumerator2 is IDisposable)
									{
										while (true)
										{
											switch (7)
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
						}
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
								FillFormat(msoSeries.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, msoSeries, ref FillColors);
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
								LineFormat(msoSeries.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, msoSeries, ref BorderColors);
							}
						}
						if (!Charts.ImplsPoints(msoSeries))
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
						foreach (ChartPoint item in (IEnumerable)msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value)))
						{
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
								FillFormat(item.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, item, ref FillColors);
								try
								{
									if (item.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											if (item.MarkerBackgroundColor == Base.TRANSPARENT)
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
												FillColors.Add(NewTuple(item.MarkerBackgroundColor, RuntimeHelpers.GetObjectValue(parent), shp, item));
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
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
								LineFormat(item.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, item, ref BorderColors);
								try
								{
									if (item.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											if (item.MarkerForegroundColor == Base.TRANSPARENT)
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
												BorderColors.Add(NewTuple(item.MarkerForegroundColor, RuntimeHelpers.GetObjectValue(parent), shp, item));
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
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							try
							{
								FontColors.Add(NewTuple(Conversions.ToInteger(item.DataLabel.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, item.DataLabel));
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
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0910;
						}
						continue;
						end_IL_0910:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
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
			using (List<Axis>.Enumerator enumerator4 = modCharts.AxesList(cht).GetEnumerator())
			{
				while (enumerator4.MoveNext())
				{
					Axis current = enumerator4.Current;
					try
					{
						_ = current.Type;
						_ = current.AxisGroup;
						Axis axis = current;
						if (axis.HasTitle)
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
								FontColors.Add(NewTuple(Conversions.ToInteger(axisTitle.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, current.AxisTitle));
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
								if (!(axisTitle.Format.Fill.Transparency < 0f))
								{
									FillFormat(axisTitle.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, current.AxisTitle, ref FillColors);
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
								LineFormat(axisTitle.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, current.AxisTitle, ref BorderColors);
							}
							axisTitle = null;
						}
						if (blnBorder)
						{
							if (axis.HasMajorGridlines)
							{
								AddColorIfNotTransparent(axis.MajorGridlines.Format.Line.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, axis.MajorGridlines, ref BorderColors);
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
								AddColorIfNotTransparent(axis.MinorGridlines.Format.Line.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, axis.MinorGridlines, ref BorderColors);
							}
							Microsoft.Office.Interop.PowerPoint.LineFormat line7 = axis.Format.Line;
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
									AddColorIfNotTransparent(line7.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, current, ref BorderColors);
								}
							}
							line7 = null;
						}
						if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
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
							TickLabels tickLabels = axis.TickLabels;
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
								FontColors.Add(NewTuple(Conversions.ToInteger(tickLabels.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, current.TickLabels));
							}
							tickLabels = null;
						}
						axis = null;
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						ProjectData.ClearProjectError();
					}
					finally
					{
						current = null;
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0be0;
					}
					continue;
					end_IL_0be0:
					break;
				}
			}
			if (chart.HasTitle)
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
				ChartTitle chartTitle = chart.ChartTitle;
				if (blnFont)
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
						if (clsCharts.ImplsFont(cht.ChartTitle))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								FontColors.Add(NewTuple(Conversions.ToInteger(chartTitle.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, cht.ChartTitle));
								break;
							}
						}
						else
						{
							int rGB = chartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
							if (rGB != 0)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									FontColors.Add(NewTuple(rGB, RuntimeHelpers.GetObjectValue(parent), shp, chartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor));
									break;
								}
							}
						}
					}
					catch (Exception projectError3)
					{
						ProjectData.SetProjectError(projectError3);
						ProjectData.ClearProjectError();
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
					if (!(chartTitle.Format.Fill.Transparency < 0f))
					{
						FillFormat(chartTitle.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, cht.ChartTitle, ref FillColors);
					}
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
					LineFormat(chartTitle.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, cht.ChartTitle, ref BorderColors);
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
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					try
					{
						enumerator5 = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Microsoft.Office.Core.LegendEntry legendEntry = (Microsoft.Office.Core.LegendEntry)enumerator5.Current;
							if (!Charts.ImplsFont(legendEntry))
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
							FontColors.Add(NewTuple(Conversions.ToInteger(legendEntry.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, cht.Legend));
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0e51;
							}
							continue;
							end_IL_0e51:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (3)
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
					if (clsCharts.UsesLegendsForSeriesClrs(cht))
					{
						{
							enumerator6 = ((IEnumerable)cht.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
							try
							{
								while (enumerator6.MoveNext())
								{
									object objectValue = RuntimeHelpers.GetObjectValue(enumerator6.Current);
									try
									{
										IMsoLegendKey obj = (IMsoLegendKey)NewLateBinding.LateGet(objectValue, null, AH.A(13177), new object[0], null, null, null);
										FillFormat(RuntimeHelpers.GetObjectValue(parent), shp, obj, ref FillColors, isMarker: false);
										FillFormat(RuntimeHelpers.GetObjectValue(parent), shp, obj, ref FillColors, isMarker: true);
									}
									catch (Exception projectError4)
									{
										ProjectData.SetProjectError(projectError4);
										ProjectData.ClearProjectError();
									}
									finally
									{
										IMsoLegendKey obj = null;
									}
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										goto end_IL_0f3d;
									}
									continue;
									end_IL_0f3d:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator6 as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
						}
					}
					else if (!(legend.Format.Fill.Transparency < 0f))
					{
						FillFormat(legend.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, cht.Legend, ref FillColors);
					}
				}
				if (blnBorder)
				{
					LineFormat(legend.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, cht.Legend, ref BorderColors);
				}
				legend = null;
			}
			if (chart.HasDataTable)
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
				DataTable dataTable = chart.DataTable;
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
					FontColors.Add(NewTuple(Conversions.ToInteger(dataTable.Font.Color), RuntimeHelpers.GetObjectValue(parent), shp, cht.DataTable));
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
					FillFormat(dataTable.Format.Fill, RuntimeHelpers.GetObjectValue(parent), shp, cht.DataTable, ref FillColors);
				}
				if (blnBorder)
				{
					LineFormat(dataTable.Format.Line, RuntimeHelpers.GetObjectValue(parent), shp, cht.DataTable, ref BorderColors);
				}
				dataTable = null;
			}
			chart = null;
			return;
		}
	}

	private static void A(TextRange2 A, object B, Microsoft.Office.Interop.PowerPoint.Shape C, object D, ref List<Tuple<int, IndexedObject>> E)
	{
		TextRange2 textRange = A;
		int count = textRange.get_Runs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			TextRange(textRange.get_Runs(i, -1), RuntimeHelpers.GetObjectValue(B), C, textRange.get_Runs(i, -1), ref E);
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
			textRange = null;
			return;
		}
	}

	public static void TextRange(TextRange2 txtRange, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		Font2 font = txtRange.Font;
		if (font.Fill.ForeColor.RGB >= 0)
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
			listColors.Add(NewTuple(font.Fill.ForeColor.RGB, RuntimeHelpers.GetObjectValue(parent), shp, txtRange));
		}
		else
		{
			listColors.Add(NewTuple(0, RuntimeHelpers.GetObjectValue(parent), shp, txtRange));
		}
		AddColorIfNotTransparent(font.Fill.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, txtRange, ref listColors);
		if (font.Highlight.RGB == 0)
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
			if (font.Highlight.Type != MsoColorType.msoColorTypeRGB)
			{
				goto IL_00d8;
			}
		}
		AddColorIfNotTransparent(font.Highlight, RuntimeHelpers.GetObjectValue(parent), shp, txtRange, ref listColors);
		goto IL_00d8;
		IL_00d8:
		if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline && font.UnderlineColor.RGB >= 0)
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
			AddColorIfNotTransparent(font.UnderlineColor, RuntimeHelpers.GetObjectValue(parent), shp, txtRange, ref listColors);
		}
		Microsoft.Office.Core.LineFormat line = font.Line;
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
			if (line.ForeColor.RGB >= 0)
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
				listColors.Add(NewTuple(line.ForeColor.RGB, RuntimeHelpers.GetObjectValue(parent), shp, txtRange));
			}
		}
		line = null;
		font = null;
	}

	public static void FillFormat(Microsoft.Office.Core.FillFormat fill, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		try
		{
			Microsoft.Office.Core.FillFormat fillFormat = fill;
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
				AddColorIfNotTransparent(fillFormat.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
				AddColorIfNotTransparent(fillFormat.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
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

	public static void FillFormat(Microsoft.Office.Interop.PowerPoint.FillFormat fill, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fillFormat = fill;
			if (fillFormat.Visible == MsoTriState.msoTrue)
			{
				AddColorIfNotTransparent(fillFormat.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
				AddColorIfNotTransparent(fillFormat.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
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

	public static void FillFormat(object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, IMsoLegendKey obj, ref List<Tuple<int, IndexedObject>> listColors, bool isMarker)
	{
		try
		{
			A(RuntimeHelpers.GetObjectValue(parent), shp, obj, ref listColors, null, isMarker);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	public static void LineFormat(Microsoft.Office.Core.LineFormat line, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		try
		{
			Microsoft.Office.Core.LineFormat lineFormat = line;
			if (lineFormat.Visible == MsoTriState.msoTrue)
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
				if (lineFormat.Weight > 0f)
				{
					AddColorIfNotTransparent(lineFormat.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
					AddColorIfNotTransparent(lineFormat.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
				}
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

	public static void LineFormat(Microsoft.Office.Interop.PowerPoint.LineFormat line, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat lineFormat = line;
			if (lineFormat.Visible == MsoTriState.msoTrue && lineFormat.Weight > 0f)
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
				AddColorIfNotTransparent(lineFormat.ForeColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
				AddColorIfNotTransparent(lineFormat.BackColor, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors);
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

	public static void AddColorIfNotTransparent(Microsoft.Office.Interop.PowerPoint.ColorFormat cf, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		int rGB = cf.RGB;
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
			listColors.Add(NewTuple(rGB, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj)));
			return;
		}
	}

	public static void AddColorIfNotTransparent(Microsoft.Office.Core.ColorFormat cf, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors)
	{
		A(RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), ref listColors, cf);
	}

	private static void A(object A, Microsoft.Office.Interop.PowerPoint.Shape B, object C, ref List<Tuple<int, IndexedObject>> D, Microsoft.Office.Core.ColorFormat E = null, bool F = false)
	{
		AddColorIntIfNotTransparent(Index.A(RuntimeHelpers.GetObjectValue(C), E, B, F), RuntimeHelpers.GetObjectValue(A), B, RuntimeHelpers.GetObjectValue(C), ref D, F);
	}

	public static void AddColorIntIfNotTransparent(int? c, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, ref List<Tuple<int, IndexedObject>> listColors, bool isMarker = false, bool areRadarLabels = false)
	{
		if (!c.HasValue)
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
			int value = c.Value;
			if (value == Base.TRANSPARENT)
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
				listColors.Add(NewTuple(value, RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), isMarker, areRadarLabels));
				return;
			}
		}
	}

	private static int? A(object A, Microsoft.Office.Core.ColorFormat B, Microsoft.Office.Interop.PowerPoint.Shape C, bool D)
	{
		int? result2;
		try
		{
			if (!(A is IMsoLegendKey))
			{
				int? result;
				if (B == null)
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
					result = null;
				}
				else
				{
					result = B.RGB;
				}
				return result;
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
				IMsoLegendKey msoLegendKey = (IMsoLegendKey)A;
				if (!D)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							return (int?)(clsCharts.UsesLegendLinesForSeriesClrs(C.Chart) ? ((object)msoLegendKey.Format.Line.ForeColor.RGB) : msoLegendKey.Interior.Color);
						}
					}
				}
				if (msoLegendKey.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone && msoLegendKey.MarkerBackgroundColor != Base.TRANSPARENT)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return msoLegendKey.MarkerBackgroundColor;
						}
					}
				}
				result2 = null;
				break;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result2 = null;
			ProjectData.ClearProjectError();
		}
		return result2;
	}

	public static Tuple<int, IndexedObject> NewTuple(int intRGB, object parent, Microsoft.Office.Interop.PowerPoint.Shape shp, object obj, bool isMarker = false, bool areRadarLabels = false)
	{
		return new Tuple<int, IndexedObject>(intRGB, new IndexedObject(RuntimeHelpers.GetObjectValue(parent), shp, RuntimeHelpers.GetObjectValue(obj), isMarker, areRadarLabels));
	}
}
