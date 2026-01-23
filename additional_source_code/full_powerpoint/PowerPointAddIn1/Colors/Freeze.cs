using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.Colors;

public sealed class Freeze
{
	[CompilerGenerated]
	internal sealed class TB
	{
		public IMsoSeries A;

		public TB(TB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal int A()
		{
			return this.A.Format.Fill.ForeColor.RGB;
		}

		[SpecialName]
		internal int B()
		{
			return this.A.Format.Fill.BackColor.RGB;
		}

		[SpecialName]
		internal int C()
		{
			return this.A.Format.Line.ForeColor.RGB;
		}

		[SpecialName]
		internal int D()
		{
			return this.A.Format.Line.BackColor.RGB;
		}
	}

	[CompilerGenerated]
	private static int m_A;

	[CompilerGenerated]
	private static int m_B;

	[CompilerGenerated]
	private static int m_C;

	[CompilerGenerated]
	private static int D;

	private static int ShapesCount
	{
		[CompilerGenerated]
		get
		{
			return Freeze.m_A;
		}
		[CompilerGenerated]
		set
		{
			Freeze.m_A = value;
		}
	}

	private static int SlidesCount
	{
		[CompilerGenerated]
		get
		{
			return Freeze.m_B;
		}
		[CompilerGenerated]
		set
		{
			Freeze.m_B = value;
		}
	}

	private static int LayoutsCount
	{
		[CompilerGenerated]
		get
		{
			return Freeze.m_C;
		}
		[CompilerGenerated]
		set
		{
			Freeze.m_C = value;
		}
	}

	private static int MastersCount
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal static void A(bool A = false)
	{
		if (!Freeze.A())
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
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
				Application application = NG.A.Application;
				B();
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
				try
				{
					activePresentation = application.ActivePresentation;
					try
					{
						application.StartNewUndoEntry();
						try
						{
							enumerator = activePresentation.Slides.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Freeze.A((Slide)enumerator.Current);
								SlidesCount++;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0076;
								}
								continue;
								end_IL_0076:
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
						try
						{
							enumerator2 = activePresentation.Designs.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Design design = (Design)enumerator2.Current;
								try
								{
									enumerator3 = design.SlideMaster.CustomLayouts.GetEnumerator();
									while (enumerator3.MoveNext())
									{
										CustomLayout customLayout = (CustomLayout)enumerator3.Current;
										try
										{
											enumerator4 = customLayout.Shapes.GetEnumerator();
											while (enumerator4.MoveNext())
											{
												Freeze.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current, B: true);
											}
											while (true)
											{
												switch (5)
												{
												case 0:
													break;
												default:
													goto end_IL_0122;
												}
												continue;
												end_IL_0122:
												break;
											}
										}
										finally
										{
											if (enumerator4 is IDisposable)
											{
												while (true)
												{
													switch (3)
													{
													case 0:
														continue;
													}
													(enumerator4 as IDisposable).Dispose();
													break;
												}
											}
										}
										LayoutsCount++;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_0168;
										}
										continue;
										end_IL_0168:
										break;
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in design.SlideMaster.Shapes)
								{
									Freeze.A(shape, B: true);
								}
								MastersCount++;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_01fe;
								}
								continue;
								end_IL_01fe:
								break;
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
						if (!A)
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
							Forms.SuccessMessage(AH.A(13344) + ShapesCount + AH.A(13379) + SlidesCount + AH.A(13402) + LayoutsCount + AH.A(13421) + MastersCount + AH.A(13450));
						}
						C();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Forms.ErrorMessage(ex2.Message);
						clsReporting.LogException(ex2);
						ProjectData.ClearProjectError();
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application = null;
				activePresentation = null;
				return;
			}
		}
	}

	internal static void A()
	{
		if (!A())
		{
			return;
		}
		Application application = NG.A.Application;
		B();
		checked
		{
			Selection selection;
			try
			{
				selection = application.ActiveWindow.Selection;
				try
				{
					switch (selection.Type)
					{
					case PpSelectionType.ppSelectionShapes:
					case PpSelectionType.ppSelectionText:
					{
						application.StartNewUndoEntry();
						IEnumerator enumerator2 = default(IEnumerator);
						try
						{
							enumerator2 = selection.ShapeRange.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, B: false);
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
								break;
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
						Forms.SuccessMessage(AH.A(13344) + ShapesCount + AH.A(13485));
						C();
						break;
					}
					case PpSelectionType.ppSelectionSlides:
						application.StartNewUndoEntry();
						foreach (Slide item in selection.SlideRange)
						{
							A(item);
							SlidesCount++;
						}
						Forms.SuccessMessage(AH.A(13344) + ShapesCount + AH.A(13506) + SlidesCount + AH.A(13533));
						C();
						break;
					default:
						Forms.WarningMessage(AH.A(13552));
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
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
		}
	}

	private static void B()
	{
		ShapesCount = 0;
		SlidesCount = 0;
		LayoutsCount = 0;
		MastersCount = 0;
	}

	private static void A(Slide A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Freeze.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B: true);
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
				return;
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

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, bool B)
	{
		if (Base.A(A))
		{
			return;
		}
		checked
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
				Microsoft.Office.Interop.PowerPoint.Shape shape = A;
				if (shape.Type == MsoShapeType.msoGroup)
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
					int count = shape.GroupItems.Count;
					for (int i = 1; i <= count; i++)
					{
						Freeze.A(shape.GroupItems[i], B);
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
				else
				{
					if (shape.HasTable == MsoTriState.msoTrue)
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
						Freeze.A(A.Table, B);
					}
					else if (shape.HasChart == MsoTriState.msoTrue)
					{
						Freeze.A(A.Chart);
					}
					else if (shape.HasSmartArt == MsoTriState.msoTrue)
					{
						Freeze.A(A.SmartArt);
					}
					else
					{
						Freeze.A(A);
						Freeze.B(A);
						C(A);
					}
					ShapesCount++;
				}
				shape = null;
				return;
			}
		}
	}

	private static void A(Table A, bool B)
	{
		Table table = A;
		int count = table.Rows.Count;
		int count2 = table.Columns.Count;
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				int num2 = count2;
				Cell cell;
				for (int j = 1; j <= num2; cell = null, j++)
				{
					cell = table.Cell(i, j);
					if (!cell.Selected)
					{
						if (!B)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
					}
					Freeze.A(cell.Shape);
					Freeze.B(cell.Shape);
					C(cell.Shape);
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0096;
					}
					continue;
					end_IL_0096:
					break;
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				table = null;
				return;
			}
		}
	}

	private static void A(Chart A)
	{
		Chart chart = A;
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = chart.ChartArea.Format.Fill;
			if (fill.Visible == MsoTriState.msoTrue && fill.ForeColor.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorMixed)
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
				Freeze.A(fill.ForeColor);
			}
			fill = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = chart.PlotArea.Format.Fill;
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
				if (fill2.ForeColor.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorMixed)
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
					Freeze.A(fill2.ForeColor);
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
			Microsoft.Office.Interop.PowerPoint.LineFormat line = chart.ChartArea.Format.Line;
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
				float weight = line.Weight;
				Freeze.A(line.ForeColor);
				if (weight < 0f)
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
					line.Visible = MsoTriState.msoFalse;
				}
				else
				{
					line.Weight = weight;
				}
			}
			line = null;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line2 = chart.PlotArea.Format.Line;
			if (line2.Visible == MsoTriState.msoTrue)
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
				float weight2 = line2.Weight;
				Freeze.A(line2.ForeColor);
				if (weight2 < 0f)
				{
					line2.Visible = MsoTriState.msoFalse;
				}
				else
				{
					line2.Weight = weight2;
				}
			}
			line2 = null;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			TB tB = default(TB);
			IEnumerator enumerator2 = default(IEnumerator);
			int markerForegroundColor = default(int);
			int markerBackgroundColor = default(int);
			IEnumerator enumerator3 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				tB = new TB(tB);
				tB.A = (IMsoSeries)enumerator.Current;
				if (Charts.ImplsAndHasErrorBars(tB.A))
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
						Microsoft.Office.Core.LineFormat line3 = tB.A.ErrorBars.Format.Line;
						Freeze.A(line3.ForeColor);
						Freeze.A(line3.BackColor);
						_ = null;
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						ProjectData.ClearProjectError();
					}
				}
				if (tB.A.HasLeaderLines)
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
						Microsoft.Office.Core.LineFormat line4 = tB.A.LeaderLines.Format.Line;
						Freeze.A(line4.ForeColor);
						Freeze.A(line4.BackColor);
						_ = null;
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
				}
				if (Charts.ImplsTrendLines(tB.A))
				{
					try
					{
						enumerator2 = ((IEnumerable)tB.A.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator2.MoveNext())
						{
							IMsoTrendline msoTrendline = (IMsoTrendline)enumerator2.Current;
							try
							{
								Microsoft.Office.Core.LineFormat line5 = msoTrendline.Format.Line;
								Freeze.A(line5.ForeColor);
								Freeze.A(line5.BackColor);
								_ = null;
							}
							catch (Exception ex13)
							{
								ProjectData.SetProjectError(ex13);
								Exception ex14 = ex13;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0357;
							}
							continue;
							end_IL_0357:
							break;
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
				Dictionary<ChartPoint, int> dictionary = new Dictionary<ChartPoint, int>();
				Dictionary<ChartPoint, int> dictionary2 = new Dictionary<ChartPoint, int>();
				Dictionary<ChartPoint, int> dictionary3 = new Dictionary<ChartPoint, int>();
				Dictionary<ChartPoint, int> dictionary4 = new Dictionary<ChartPoint, int>();
				Dictionary<ChartPoint, int> dictionary5 = new Dictionary<ChartPoint, int>();
				Dictionary<ChartPoint, int> dictionary6 = new Dictionary<ChartPoint, int>();
				int num = Freeze.A(tB.A, -1);
				int num2 = Freeze.A(tB.B, -1);
				int num3 = Freeze.A(tB.C, -1);
				int num4 = Freeze.A(tB.D, -1);
				bool flag;
				try
				{
					if (tB.A.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							markerForegroundColor = tB.A.MarkerForegroundColor;
							markerBackgroundColor = tB.A.MarkerBackgroundColor;
							flag = true;
							break;
						}
					}
					else
					{
						flag = false;
					}
				}
				catch (Exception ex15)
				{
					ProjectData.SetProjectError(ex15);
					Exception ex16 = ex15;
					flag = false;
					ProjectData.ClearProjectError();
				}
				if (Charts.ImplsPoints(tB.A))
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
						enumerator3 = ((IEnumerable)tB.A.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator3.MoveNext())
						{
							ChartPoint chartPoint = (ChartPoint)enumerator3.Current;
							try
							{
								Microsoft.Office.Core.ColorFormat foreColor = chartPoint.Format.Fill.ForeColor;
								if (foreColor.RGB != num)
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
									dictionary3.Add(chartPoint, foreColor.RGB);
								}
								foreColor = null;
							}
							catch (Exception ex17)
							{
								ProjectData.SetProjectError(ex17);
								Exception ex18 = ex17;
								ProjectData.ClearProjectError();
							}
							try
							{
								Microsoft.Office.Core.ColorFormat backColor = chartPoint.Format.Fill.BackColor;
								if (backColor.RGB != num2)
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
									dictionary4.Add(chartPoint, backColor.RGB);
								}
								backColor = null;
							}
							catch (Exception ex19)
							{
								ProjectData.SetProjectError(ex19);
								Exception ex20 = ex19;
								ProjectData.ClearProjectError();
							}
							try
							{
								Microsoft.Office.Core.ColorFormat foreColor2 = chartPoint.Format.Line.ForeColor;
								if (foreColor2.RGB != num3)
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
									dictionary5.Add(chartPoint, foreColor2.RGB);
								}
								foreColor2 = null;
							}
							catch (Exception ex21)
							{
								ProjectData.SetProjectError(ex21);
								Exception ex22 = ex21;
								ProjectData.ClearProjectError();
							}
							try
							{
								Microsoft.Office.Core.ColorFormat backColor2 = chartPoint.Format.Line.BackColor;
								if (backColor2.RGB != num4)
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
									dictionary6.Add(chartPoint, backColor2.RGB);
								}
								backColor2 = null;
							}
							catch (Exception ex23)
							{
								ProjectData.SetProjectError(ex23);
								Exception ex24 = ex23;
								ProjectData.ClearProjectError();
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
								try
								{
									if (chartPoint.MarkerForegroundColor != markerForegroundColor)
									{
										dictionary.Add(chartPoint, chartPoint.MarkerForegroundColor);
									}
								}
								catch (Exception ex25)
								{
									ProjectData.SetProjectError(ex25);
									Exception ex26 = ex25;
									ProjectData.ClearProjectError();
								}
								try
								{
									if (chartPoint.MarkerBackgroundColor != markerBackgroundColor)
									{
										dictionary2.Add(chartPoint, chartPoint.MarkerBackgroundColor);
									}
								}
								catch (Exception ex27)
								{
									ProjectData.SetProjectError(ex27);
									Exception ex28 = ex27;
									ProjectData.ClearProjectError();
								}
							}
							try
							{
								Freeze.A(chartPoint.DataLabel.Font);
							}
							catch (Exception ex29)
							{
								ProjectData.SetProjectError(ex29);
								Exception ex30 = ex29;
								ProjectData.ClearProjectError();
							}
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
				}
				try
				{
					Freeze.A(tB.A.Format.Fill.ForeColor);
					Freeze.A(tB.A.Format.Fill.BackColor);
					using (Dictionary<ChartPoint, int>.Enumerator enumerator4 = dictionary3.GetEnumerator())
					{
						while (enumerator4.MoveNext())
						{
							KeyValuePair<ChartPoint, int> current = enumerator4.Current;
							current.Key.Format.Fill.ForeColor.RGB = current.Value;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0767;
							}
							continue;
							end_IL_0767:
							break;
						}
					}
					using Dictionary<ChartPoint, int>.Enumerator enumerator5 = dictionary4.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						KeyValuePair<ChartPoint, int> current2 = enumerator5.Current;
						current2.Key.Format.Fill.BackColor.RGB = current2.Value;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_07cd;
						}
						continue;
						end_IL_07cd:
						break;
					}
				}
				catch (Exception ex31)
				{
					ProjectData.SetProjectError(ex31);
					Exception ex32 = ex31;
					ProjectData.ClearProjectError();
				}
				dictionary3 = null;
				dictionary4 = null;
				try
				{
					if (tB.A.Format.Line.Visible == MsoTriState.msoTrue)
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
						if (tB.A.Format.Line.Weight > 0f)
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
							Freeze.A(tB.A.Format.Line.ForeColor);
							Freeze.A(tB.A.Format.Line.BackColor);
						}
					}
					using (Dictionary<ChartPoint, int>.Enumerator enumerator6 = dictionary5.GetEnumerator())
					{
						while (enumerator6.MoveNext())
						{
							KeyValuePair<ChartPoint, int> current3 = enumerator6.Current;
							current3.Key.Format.Line.ForeColor.RGB = current3.Value;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_08da;
							}
							continue;
							end_IL_08da:
							break;
						}
					}
					using Dictionary<ChartPoint, int>.Enumerator enumerator7 = dictionary6.GetEnumerator();
					while (enumerator7.MoveNext())
					{
						KeyValuePair<ChartPoint, int> current4 = enumerator7.Current;
						current4.Key.Format.Line.BackColor.RGB = current4.Value;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0942;
						}
						continue;
						end_IL_0942:
						break;
					}
				}
				catch (Exception ex33)
				{
					ProjectData.SetProjectError(ex33);
					Exception ex34 = ex33;
					ProjectData.ClearProjectError();
				}
				dictionary5 = null;
				dictionary6 = null;
				if (!flag)
				{
					continue;
				}
				bool num5 = tB.A.Format.Line.Visible == MsoTriState.msoTrue;
				if (num5)
				{
					tB.A.Format.Line.Visible = MsoTriState.msoFalse;
				}
				tB.A.MarkerForegroundColor = markerForegroundColor;
				tB.A.MarkerBackgroundColor = markerBackgroundColor;
				if (num5)
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
					tB.A.Format.Line.Visible = MsoTriState.msoTrue;
				}
				using (Dictionary<ChartPoint, int>.Enumerator enumerator8 = dictionary.GetEnumerator())
				{
					while (enumerator8.MoveNext())
					{
						KeyValuePair<ChartPoint, int> current5 = enumerator8.Current;
						current5.Key.MarkerForegroundColor = current5.Value;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0a31;
						}
						continue;
						end_IL_0a31:
						break;
					}
				}
				dictionary = null;
				using (Dictionary<ChartPoint, int>.Enumerator enumerator9 = dictionary2.GetEnumerator())
				{
					while (enumerator9.MoveNext())
					{
						KeyValuePair<ChartPoint, int> current6 = enumerator9.Current;
						current6.Key.MarkerBackgroundColor = current6.Value;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0a84;
						}
						continue;
						end_IL_0a84:
						break;
					}
				}
				dictionary2 = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0aaf;
				}
				continue;
				end_IL_0aaf:
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
		try
		{
			foreach (Axis item in modCharts.AxesList(A))
			{
				try
				{
					_ = item.Type;
					Axis axis = item;
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
						try
						{
							Microsoft.Office.Interop.PowerPoint.LineFormat line6 = axisTitle.Format.Line;
							if (line6.Visible == MsoTriState.msoTrue)
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
								if (line6.Weight > 0f)
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
									Freeze.A(line6.ForeColor);
									Freeze.A(line6.BackColor);
								}
							}
							line6 = null;
						}
						catch (Exception ex35)
						{
							ProjectData.SetProjectError(ex35);
							Exception ex36 = ex35;
							ProjectData.ClearProjectError();
						}
						try
						{
							Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = axisTitle.Format.Fill;
							if (fill3.Visible == MsoTriState.msoTrue)
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
								if (fill3.ForeColor.RGB == 0)
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
									if (fill3.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
									{
										goto IL_0c07;
									}
								}
								Freeze.A(fill3.ForeColor);
								Freeze.A(fill3.BackColor);
							}
							goto IL_0c07;
							IL_0c07:
							fill3 = null;
						}
						catch (Exception ex37)
						{
							ProjectData.SetProjectError(ex37);
							Exception ex38 = ex37;
							ProjectData.ClearProjectError();
						}
						Freeze.A(axisTitle.Font);
						axisTitle = null;
					}
					if (axis.HasMajorGridlines)
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
						Freeze.A(axis.MajorGridlines.Format.Line.ForeColor);
					}
					if (axis.HasMinorGridlines)
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
						Freeze.A(axis.MinorGridlines.Format.Line.ForeColor);
					}
					Microsoft.Office.Interop.PowerPoint.LineFormat line7 = axis.Format.Line;
					if (line7.Weight > 0f)
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
						if (line7.Visible == MsoTriState.msoTrue)
						{
							Freeze.A(line7.ForeColor);
						}
					}
					line7 = null;
					axis = null;
				}
				finally
				{
					Axis current7 = null;
				}
			}
		}
		catch (Exception ex39)
		{
			ProjectData.SetProjectError(ex39);
			Exception ex40 = ex39;
			ProjectData.ClearProjectError();
		}
		int count = ((ChartGroups)A.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			ChartGroup chartGroup = (ChartGroup)A.ChartGroups(i);
			if (chartGroup.HasUpDownBars)
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
					Microsoft.Office.Interop.PowerPoint.FillFormat fill4 = chartGroup.UpBars.Format.Fill;
					Freeze.A(fill4.ForeColor);
					Freeze.A(fill4.BackColor);
					_ = null;
					Microsoft.Office.Interop.PowerPoint.LineFormat line8 = chartGroup.UpBars.Format.Line;
					Freeze.A(line8.ForeColor);
					Freeze.A(line8.BackColor);
					_ = null;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill5 = chartGroup.DownBars.Format.Fill;
					Freeze.A(fill5.ForeColor);
					Freeze.A(fill5.BackColor);
					_ = null;
					Microsoft.Office.Interop.PowerPoint.LineFormat line9 = chartGroup.DownBars.Format.Line;
					Freeze.A(line9.ForeColor);
					Freeze.A(line9.BackColor);
					_ = null;
				}
				catch (Exception ex41)
				{
					ProjectData.SetProjectError(ex41);
					Exception ex42 = ex41;
					ProjectData.ClearProjectError();
				}
			}
			if (chartGroup.HasHiLoLines)
			{
				try
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line10 = chartGroup.HiLoLines.Format.Line;
					Freeze.A(line10.ForeColor);
					Freeze.A(line10.BackColor);
					_ = null;
				}
				catch (Exception ex43)
				{
					ProjectData.SetProjectError(ex43);
					Exception ex44 = ex43;
					ProjectData.ClearProjectError();
				}
			}
			if (chartGroup.HasDropLines)
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
					Microsoft.Office.Interop.PowerPoint.LineFormat line11 = chartGroup.DropLines.Format.Line;
					Freeze.A(line11.ForeColor);
					Freeze.A(line11.BackColor);
					_ = null;
				}
				catch (Exception ex45)
				{
					ProjectData.SetProjectError(ex45);
					Exception ex46 = ex45;
					ProjectData.ClearProjectError();
				}
			}
			chartGroup = null;
		}
		IEnumerator enumerator11 = default(IEnumerator);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
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
				ChartTitle chartTitle = chart.ChartTitle;
				try
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line12 = chartTitle.Format.Line;
					if (line12.Visible == MsoTriState.msoTrue && line12.Weight > 0f)
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
						Freeze.A(line12.ForeColor);
						Freeze.A(line12.BackColor);
					}
					line12 = null;
				}
				catch (Exception ex47)
				{
					ProjectData.SetProjectError(ex47);
					Exception ex48 = ex47;
					ProjectData.ClearProjectError();
				}
				try
				{
					Microsoft.Office.Interop.PowerPoint.FillFormat fill6 = chartTitle.Format.Fill;
					if (fill6.Visible == MsoTriState.msoTrue)
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
						if (fill6.ForeColor.RGB == 0)
						{
							if (fill6.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
							{
								goto IL_0fe6;
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
						Freeze.A(fill6.ForeColor);
						Freeze.A(fill6.BackColor);
					}
					goto IL_0fe6;
					IL_0fe6:
					fill6 = null;
				}
				catch (Exception ex49)
				{
					ProjectData.SetProjectError(ex49);
					Exception ex50 = ex49;
					ProjectData.ClearProjectError();
				}
				if (clsCharts.ImplsFont(A.ChartTitle))
				{
					Freeze.A(chartTitle.Font);
				}
				chartTitle = null;
			}
			if (chart.HasLegend)
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
				Legend legend = chart.Legend;
				try
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line13 = legend.Format.Line;
					if (line13.Visible == MsoTriState.msoTrue)
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
						if (line13.Weight > 0f)
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
							Freeze.A(line13.ForeColor);
							Freeze.A(line13.BackColor);
						}
					}
					line13 = null;
				}
				catch (Exception ex51)
				{
					ProjectData.SetProjectError(ex51);
					Exception ex52 = ex51;
					ProjectData.ClearProjectError();
				}
				try
				{
					Microsoft.Office.Interop.PowerPoint.FillFormat fill7 = legend.Format.Fill;
					if (fill7.Visible == MsoTriState.msoTrue)
					{
						if (fill7.ForeColor.RGB == 0)
						{
							if (fill7.ForeColor.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorMixed)
							{
								goto IL_1114;
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
						Freeze.A(fill7.ForeColor);
						Freeze.A(fill7.BackColor);
					}
					goto IL_1114;
					IL_1114:
					fill7 = null;
				}
				catch (Exception ex53)
				{
					ProjectData.SetProjectError(ex53);
					Exception ex54 = ex53;
					ProjectData.ClearProjectError();
				}
				{
					enumerator11 = ((IEnumerable)legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					try
					{
						while (enumerator11.MoveNext())
						{
							Microsoft.Office.Core.LegendEntry legendEntry = (Microsoft.Office.Core.LegendEntry)enumerator11.Current;
							if (Charts.ImplsFont(legendEntry))
							{
								Freeze.A(legendEntry.Font);
							}
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_1180;
							}
							continue;
							end_IL_1180:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator11 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				legend = null;
			}
			if (chart.HasDataTable)
			{
				DataTable dataTable = chart.DataTable;
				Freeze.A(dataTable.Format.Line.ForeColor);
				Freeze.A(dataTable.Font);
				_ = null;
			}
			chart = null;
			return;
		}
	}

	private static int A(Func<int> A, int B)
	{
		int result;
		try
		{
			result = A();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = B;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void A(SmartArt A)
	{
		IEnumerator enumerator = A.Nodes.GetEnumerator();
		try
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				try
				{
					if (smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue)
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
							Freeze.A(smartArtNode.TextFrame2.TextRange);
							break;
						}
					}
					else
					{
						B(smartArtNode.TextFrame2.TextRange);
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				{
					enumerator2 = smartArtNode.Shapes.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator2.Current;
							try
							{
								Microsoft.Office.Core.FillFormat fill = shape.Fill;
								if (fill.Visible == MsoTriState.msoTrue)
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
									Freeze.A(fill.ForeColor);
									Freeze.A(fill.BackColor);
								}
								fill = null;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							try
							{
								Microsoft.Office.Core.LineFormat line = shape.Line;
								if (line.Visible == MsoTriState.msoTrue)
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
									Freeze.A(line.ForeColor);
									Freeze.A(line.BackColor);
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
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_014b;
							}
							continue;
							end_IL_014b:
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		try
		{
			if (shape.HasTextFrame == MsoTriState.msoTrue)
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
					if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							Freeze.A(shape.TextFrame2.TextRange);
							break;
						}
					}
					else
					{
						B(shape.TextFrame2.TextRange);
					}
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
		shape = null;
	}

	private static void A(TextRange2 A)
	{
		TextRange2 textRange = A;
		int count = textRange.get_Runs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			Font2 font = textRange.get_Runs(i, -1).Font;
			Freeze.A(font.Fill.ForeColor);
			Freeze.A(font.Fill.BackColor);
			_ = null;
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
			textRange = null;
			return;
		}
	}

	private static void B(TextRange2 A)
	{
		Font2 font = A.Font;
		Freeze.A(font.Fill.ForeColor);
		Freeze.A(font.Fill.BackColor);
		_ = null;
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = A.Fill;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Freeze.A(fill.ForeColor);
				Freeze.A(fill.BackColor);
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

	private static void C(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
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
				Freeze.A(line.ForeColor);
				Freeze.A(line.BackColor);
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

	private static void A(Microsoft.Office.Interop.PowerPoint.ColorFormat A)
	{
		if (A.RGB == Base.TRANSPARENT)
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
			A.RGB = A.RGB;
			return;
		}
	}

	private static void A(Microsoft.Office.Core.ColorFormat A)
	{
		if (A.RGB == Base.TRANSPARENT)
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
			A.RGB = A.RGB;
			return;
		}
	}

	private static void A(Microsoft.Office.Core.ChartFont A)
	{
		A.Color = RuntimeHelpers.GetObjectValue(A.Color);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.ChartFont A)
	{
		A.Color = RuntimeHelpers.GetObjectValue(A.Color);
	}

	private static void A(ChartColorFormat A)
	{
		NewLateBinding.LateSet(A, null, AH.A(13587), new object[1] { NewLateBinding.LateGet(A, null, AH.A(13587), new object[0], null, null, null) }, null, null);
	}

	private static bool A()
	{
		return Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false);
	}

	private static void C()
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(13598));
	}
}
