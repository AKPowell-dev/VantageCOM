using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Colors;

public sealed class Ribbon
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Point, int> A;

		public static Func<Point, int> B;

		public static Func<Point, int> C;

		public static Func<Point, int> D;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Point A)
		{
			return A.Y;
		}

		[SpecialName]
		internal int B(Point A)
		{
			return A.Y;
		}

		[SpecialName]
		internal int C(Point A)
		{
			return A.X;
		}

		[SpecialName]
		internal int D(Point A)
		{
			return A.X;
		}
	}

	internal static void A(int A)
	{
		try
		{
			Ribbon.A(clsColors.ColorPalette[A].RGB);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			Ribbon.A(color);
			PB.Settings.LastFontColor = color;
			KG.A.InvalidateControl(clsColors.LAST_FONT_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13652));
	}

	internal static void A()
	{
		try
		{
			A(PB.Settings.LastFontColor);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13689));
	}

	private static void A(Color A)
	{
		Application application = NG.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		application.StartNewUndoEntry();
		checked
		{
			if (selection.Type == PpSelectionType.ppSelectionShapes)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = PowerPointAddIn1.Shapes.Base.SelectedShapes().GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						if (shape.HasTable == MsoTriState.msoTrue)
						{
							Table table = shape.Table;
							int count = table.Rows.Count;
							for (int i = 1; i <= count; i++)
							{
								int count2 = table.Columns.Count;
								for (int j = 1; j <= count2; j++)
								{
									Cell cell = table.Cell(i, j);
									if (cell.Selected)
									{
										Ribbon.A(A, cell.Shape);
									}
									cell = null;
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
							table = null;
						}
						else
						{
							Ribbon.A(A, shape);
						}
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_010d;
						}
						continue;
						end_IL_010d:
						break;
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
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			else if (selection.Type == PpSelectionType.ppSelectionText)
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
				selection.TextRange.Font.Color.RGB = ColorTranslator.ToOle(A);
			}
			Ribbon.A(application, A);
			application = null;
			selection = null;
		}
	}

	private static void A(Color A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		try
		{
			B.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void B(int A)
	{
		try
		{
			B(clsColors.ColorPalette[A].RGB);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void B(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			B(color);
			PB.Settings.LastFillColor = color;
			KG.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13724));
	}

	internal static void B()
	{
		try
		{
			B(PB.Settings.LastFillColor);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13761));
	}

	internal static void C()
	{
		try
		{
			B(Color.Transparent);
			PB.Settings.LastFillColor = Color.Transparent;
			KG.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13796));
	}

	private static void B(Color A)
	{
		Application application = NG.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		application.StartNewUndoEntry();
		checked
		{
			if (selection.Type == PpSelectionType.ppSelectionShapes)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = PowerPointAddIn1.Shapes.Base.SelectedShapes().GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						if (shape.HasTable == MsoTriState.msoTrue)
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
							Table table = shape.Table;
							int count = table.Rows.Count;
							for (int i = 1; i <= count; i++)
							{
								int count2 = table.Columns.Count;
								for (int j = 1; j <= count2; j++)
								{
									Cell cell = table.Cell(i, j);
									if (cell.Selected)
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
										Ribbon.A(A, cell.Shape.Fill);
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
								break;
							}
							table = null;
						}
						else
						{
							Ribbon.A(A, shape.Fill);
						}
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0129;
						}
						continue;
						end_IL_0129:
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
			else if (selection.Type == PpSelectionType.ppSelectionText)
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
				Font2 font = selection.TextRange2.Font;
				float size = font.Size;
				string name = font.Name;
				int rGB = font.Fill.ForeColor.RGB;
				font.Highlight.RGB = ColorTranslator.ToOle(A);
				font.Size = size;
				font.Name = name;
				font.Fill.ForeColor.RGB = rGB;
				_ = null;
			}
			Ribbon.A(application, A);
			application = null;
		}
	}

	private static void A(Color A, Microsoft.Office.Interop.PowerPoint.FillFormat B)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fillFormat = B;
			if (A == Color.Transparent)
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
				fillFormat.Visible = MsoTriState.msoFalse;
			}
			else
			{
				fillFormat.Visible = MsoTriState.msoTrue;
				fillFormat.ForeColor.RGB = ColorTranslator.ToOle(A);
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

	internal static void C(int A)
	{
		try
		{
			C(clsColors.ColorPalette[A].RGB);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void C(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			C(color);
			PB.Settings.LastBorderColor = color;
			KG.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13823));
	}

	internal static void D()
	{
		try
		{
			C(PB.Settings.LastBorderColor);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13864));
	}

	internal static void E()
	{
		try
		{
			C(Color.Transparent);
			PB.Settings.LastBorderColor = Color.Transparent;
			KG.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		D(AH.A(13903));
	}

	private static void C(Color A)
	{
		Application application = NG.A.Application;
		application.StartNewUndoEntry();
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = PowerPointAddIn1.Shapes.Base.SelectedShapes().GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					try
					{
						if (shape.HasTable == MsoTriState.msoTrue)
						{
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
								List<Point> list = new List<Point>();
								Table table = shape.Table;
								int count = table.Rows.Count;
								for (int i = 1; i <= count; i++)
								{
									int count2 = table.Columns.Count;
									for (int j = 1; j <= count2; j++)
									{
										if (table.Cell(i, j).Selected)
										{
											list.Add(new Point(j, i));
										}
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_00c3;
										}
										continue;
										end_IL_00c3:
										break;
									}
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									if (list.Count == table.Rows.Count * table.Columns.Count)
									{
										Ribbon.A(table.Rows[1].Cells.Borders[PpBorderType.ppBorderTop], A);
										Ribbon.A(table.Rows[table.Rows.Count].Cells.Borders[PpBorderType.ppBorderBottom], A);
										Ribbon.A(table.Columns[1].Cells.Borders[PpBorderType.ppBorderLeft], A);
										Ribbon.A(table.Columns[table.Columns.Count].Cells.Borders[PpBorderType.ppBorderRight], A);
									}
									else if (list.Count == 1)
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
										Cell cell = table.Cell(list[0].Y, list[0].X);
										Ribbon.A(cell.Borders[PpBorderType.ppBorderTop], A);
										Ribbon.A(cell.Borders[PpBorderType.ppBorderBottom], A);
										Ribbon.A(cell.Borders[PpBorderType.ppBorderLeft], A);
										Ribbon.A(cell.Borders[PpBorderType.ppBorderRight], A);
										_ = null;
									}
									else
									{
										int num = list.Min([SpecialName] (Point point) => point.Y);
										List<Point> source = list;
										Func<Point, int> selector;
										if (_Closure_0024__.B == null)
										{
											selector = (_Closure_0024__.B = [SpecialName] (Point point) => point.Y);
										}
										else
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
											selector = _Closure_0024__.B;
										}
										int num2 = source.Max(selector);
										List<Point> source2 = list;
										Func<Point, int> selector2;
										if (_Closure_0024__.C == null)
										{
											selector2 = (_Closure_0024__.C = [SpecialName] (Point point) => point.X);
										}
										else
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
											selector2 = _Closure_0024__.C;
										}
										int num3 = source2.Min(selector2);
										List<Point> source3 = list;
										Func<Point, int> selector3;
										if (_Closure_0024__.D == null)
										{
											selector3 = (_Closure_0024__.D = [SpecialName] (Point point) => point.X);
										}
										else
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
											selector3 = _Closure_0024__.D;
										}
										int num4 = source3.Max(selector3);
										using List<Point>.Enumerator enumerator2 = list.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											Point current = enumerator2.Current;
											Cell cell2 = table.Cell(current.Y, current.X);
											if (current.Y == num)
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
												Ribbon.A(cell2.Borders[PpBorderType.ppBorderTop], A);
											}
											if (current.Y == num2)
											{
												Ribbon.A(cell2.Borders[PpBorderType.ppBorderBottom], A);
											}
											if (current.X == num3)
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
												Ribbon.A(cell2.Borders[PpBorderType.ppBorderLeft], A);
											}
											if (current.X == num4)
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
												Ribbon.A(cell2.Borders[PpBorderType.ppBorderRight], A);
											}
											cell2 = null;
										}
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_042f;
											}
											continue;
											end_IL_042f:
											break;
										}
									}
									table = null;
									list = null;
									break;
								}
								break;
							}
						}
						else
						{
							Ribbon.A(shape.Line, A);
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
			Ribbon.A(application, A);
			application = null;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.LineFormat A, Color B)
	{
		if (B != Color.Transparent)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.Visible = MsoTriState.msoTrue;
					A.ForeColor.RGB = ColorTranslator.ToOle(B);
					return;
				}
			}
		}
		A.Visible = MsoTriState.msoFalse;
	}

	private static void A(Application A, Color B)
	{
		if (!(B != Color.Transparent))
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
			try
			{
				A.ActivePresentation.ExtraColors.Add(ColorTranslator.ToOle(B));
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
	}

	private static void D(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, A);
	}
}
