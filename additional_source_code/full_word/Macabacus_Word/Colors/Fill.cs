using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using A;
using MacabacusMacros;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Colors;

public sealed class Fill
{
	internal static void A(int A)
	{
		Fill.A(clsColors.ColorPalette[A].RGB);
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			Fill.A(color);
			N.Settings.LastFillColor = color;
			NC.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ButtonColor()
	{
		try
		{
			A(N.Settings.LastFillColor);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void NoFill()
	{
		try
		{
			A(Color.Transparent);
			N.Settings.LastFillColor = Color.Transparent;
			NC.A.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Cycle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		List<Color> list = new List<Color>();
		checked
		{
			try
			{
				list = NC.A.FillColorCycle.Distinct().ToList();
				if (NC.A == null)
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
					NC.A = new clsGlobals();
				}
				clsGlobals a = NC.A;
				if (a.FillColorCycle > list.Count - 1)
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
					a.FillColorCycle = 0;
				}
				A(list[a.FillColorCycle]);
				a.FillColorCycle++;
				a = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			list = null;
		}
	}

	private static void A(Color A)
	{
		Application application = PC.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		UndoRecord undoRecord = application.UndoRecord;
		application.ScreenUpdating = false;
		checked
		{
			try
			{
				switch (selection.Type)
				{
				default:
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
						undoRecord.StartCustomRecord(XC.A(3065));
						if (Conversions.ToBoolean(selection.get_Information(WdInformation.wdWithInTable)))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								Table table = selection.Tables[1];
								int num = 0;
								int count = table.Rows.Count;
								for (int i = 1; i <= count; i++)
								{
									int count2 = table.Columns.Count;
									for (int j = 1; j <= count2; j++)
									{
										Cell cell = table.Cell(i, j);
										if (cell.Range.InRange(selection.Range))
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
											Shading shading = cell.Shading;
											shading.Texture = WdTextureIndex.wdTextureNone;
											shading.BackgroundPatternColor = Helpers.ColorToWdColor(A);
											shading.ForegroundPatternColor = Helpers.ColorToWdColor(A);
											_ = null;
											num++;
										}
										cell = null;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_021e;
										}
										continue;
										end_IL_021e:
										break;
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
									if (num != 0)
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
										Fill.A(selection, A);
										break;
									}
									break;
								}
								break;
							}
						}
						else
						{
							Fill.A(selection, A);
						}
						break;
					}
					break;
				case WdSelectionType.wdSelectionInlineShape:
				{
					undoRecord.StartCustomRecord(XC.A(3044));
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = selection.ChildShapeRange.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Fill.A((Microsoft.Office.Interop.Word.Shape)enumerator2.Current, A);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_009d;
							}
							continue;
							end_IL_009d:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					break;
				}
				case WdSelectionType.wdSelectionShape:
				{
					undoRecord.StartCustomRecord(XC.A(3044));
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = Macabacus_Word.Shapes.Helpers.SelectedShapes(selection).GetEnumerator();
						while (enumerator.MoveNext())
						{
							Fill.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, A);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_010c;
							}
							continue;
							end_IL_010c:
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
			undoRecord.EndCustomRecord();
			application.ScreenUpdating = true;
			undoRecord = null;
			selection = null;
			application = null;
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, Color B)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					try
					{
						Fill.A(B, A.Fill);
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
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Fill.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, B);
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

	private static void A(Color A, Microsoft.Office.Interop.Word.FillFormat B)
	{
		try
		{
			Microsoft.Office.Interop.Word.FillFormat fillFormat = B;
			if (A == Color.Transparent)
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

	private static void A(Selection A, Color B)
	{
		Shading shading = A.Range.Font.Shading;
		shading.Texture = WdTextureIndex.wdTextureNone;
		shading.BackgroundPatternColor = Helpers.ColorToWdColor(B);
		_ = null;
	}
}
