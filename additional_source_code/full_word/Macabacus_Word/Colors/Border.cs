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

public sealed class Border
{
	internal static void A(int A)
	{
		Border.A(clsColors.ColorPalette[A].RGB);
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			Border.A(color);
			N.Settings.LastBorderColor = color;
			NC.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
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
			A(N.Settings.LastBorderColor);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void NoBorder()
	{
		try
		{
			A(Color.Transparent);
			N.Settings.LastBorderColor = Color.Transparent;
			NC.A.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
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
				list = NC.A.BorderColorCycle.Distinct().ToList();
				if (NC.A == null)
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
					NC.A = new clsGlobals();
				}
				clsGlobals a = NC.A;
				if (a.BorderColorCycle > list.Count - 1)
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
					a.BorderColorCycle = 0;
				}
				A(list[a.BorderColorCycle]);
				a.BorderColorCycle++;
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
		Helpers.ColorToWdColor(A);
		application.ScreenUpdating = false;
		undoRecord.StartCustomRecord(XC.A(3019));
		try
		{
			switch (selection.Type)
			{
			case WdSelectionType.wdSelectionInlineShape:
				{
					IEnumerator enumerator2 = selection.ChildShapeRange.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							Border.A((Microsoft.Office.Interop.Word.Shape)enumerator2.Current, A);
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
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				break;
			case WdSelectionType.wdSelectionShape:
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = Macabacus_Word.Shapes.Helpers.SelectedShapes(selection).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Border.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, A);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00ec;
						}
						continue;
						end_IL_00ec:
						break;
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
				break;
			}
			default:
			{
				Borders borders = selection.Range.Borders;
				borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
				borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
				borders.OutsideColor = Helpers.ColorToWdColor(A);
				_ = null;
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
						Microsoft.Office.Interop.Word.LineFormat line = A.Line;
						if (B == Color.Transparent)
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
							line.Visible = MsoTriState.msoTrue;
							line.ForeColor.RGB = ColorTranslator.ToOle(B);
						}
						line = null;
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
				Border.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, B);
			}
			while (true)
			{
				switch (7)
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

	private static void A(InlineShape A, Color B)
	{
		try
		{
			Microsoft.Office.Interop.Word.LineFormat line = A.Line;
			if (B == Color.Transparent)
			{
				line.Visible = MsoTriState.msoFalse;
			}
			else
			{
				line.Visible = MsoTriState.msoTrue;
				line.ForeColor.RGB = ColorTranslator.ToOle(B);
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
}
