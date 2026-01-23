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
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Colors;

public sealed class Font
{
	internal static void A(int A)
	{
		Font.A(clsColors.ColorPalette[A].RGB);
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			Font.A(color);
			N.Settings.LastFontColor = color;
			NC.A.InvalidateControl(clsColors.LAST_FONT_COLOR_BUTTON);
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
			A(N.Settings.LastFontColor);
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
		checked
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
				List<Color> list = new List<Color>();
				try
				{
					list = NC.A.FontColorCycle.Distinct().ToList();
					if (NC.A == null)
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
						NC.A = new clsGlobals();
					}
					clsGlobals a = NC.A;
					if (a.FontColorCycle > list.Count - 1)
					{
						a.FontColorCycle = 0;
					}
					A(list[a.FontColorCycle]);
					a.FontColorCycle++;
					a = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				list = null;
				return;
			}
		}
	}

	private static void A(Color A)
	{
		Application application = PC.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		UndoRecord undoRecord = application.UndoRecord;
		Helpers.ColorToWdColor(A);
		application.ScreenUpdating = false;
		undoRecord.StartCustomRecord(XC.A(3092));
		try
		{
			switch (selection.Type)
			{
			case WdSelectionType.wdSelectionInlineShape:
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = selection.ChildShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Font.A((Microsoft.Office.Interop.Word.Shape)enumerator2.Current, A);
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
				break;
			}
			case WdSelectionType.wdSelectionShape:
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = Macabacus_Word.Shapes.Helpers.SelectedShapes(selection).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Font.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, A);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00f6;
						}
						continue;
						end_IL_00f6:
						break;
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
				break;
			}
			default:
				selection.Range.Font.Color = Helpers.ColorToWdColor(A);
				break;
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

	private static void A(Microsoft.Office.Interop.Word.Shape A, Color B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (1)
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
						A.TextFrame.TextRange.Font.TextColor.RGB = Information.RGB(B.R, B.G, B.B);
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
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Font.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, B);
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
