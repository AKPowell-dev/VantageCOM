using System;
using System.Drawing;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsRibbon
{
	public static void InvalidateLinkedItemControls()
	{
		clsRibbon.InvalidateLinkedItemControls(NC.A);
	}

	public static Bitmap RecolorColorButton(string id)
	{
		Bitmap bitmap = default(Bitmap);
		Color color = default(Color);
		if (Operators.CompareString(id, clsColors.LAST_FONT_COLOR_BUTTON, TextCompare: false) == 0)
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
			bitmap = new Bitmap(M.FontColorPicker);
			color = N.Settings.LastFontColor;
		}
		else if (Operators.CompareString(id, clsColors.LAST_FILL_COLOR_BUTTON, TextCompare: false) == 0)
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
			Selection selection = PC.A.Application.Selection;
			if (selection != null)
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
				WdSelectionType type = selection.Type;
				if ((uint)(type - 7) <= 1u)
				{
					bitmap = ((N.Settings.LastFillColor == Color.Transparent) ? new Bitmap(M.NoFill) : new Bitmap(M.FillColorPicker));
				}
				else
				{
					Bitmap bitmap2;
					if (!(N.Settings.LastFillColor == Color.Transparent))
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
						bitmap2 = new Bitmap(M.HighlightColorPicker);
					}
					else
					{
						bitmap2 = new Bitmap(M.NoHighlight);
					}
					bitmap = bitmap2;
				}
				selection = null;
			}
			else
			{
				bitmap = new Bitmap(M.HighlightColorPicker);
			}
			color = N.Settings.LastFillColor;
		}
		else if (Operators.CompareString(id, clsColors.LAST_BORDER_COLOR_BUTTON, TextCompare: false) == 0)
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
			Bitmap bitmap3;
			if (!(N.Settings.LastBorderColor == Color.Transparent))
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
				bitmap3 = new Bitmap(M.BorderColorPicker);
			}
			else
			{
				bitmap3 = new Bitmap(M.NoBorder);
			}
			bitmap = bitmap3;
			color = N.Settings.LastBorderColor;
		}
		return clsColors.RecolorColorButton(bitmap, color);
	}

	public static void Application_WindowSelectionChange(Selection Sel)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					break;
				case 49:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
							goto end_IL_0000_3;
						}
						goto default;
					}
					end_IL_0000_2:
					break;
				}
				num2 = 2;
				clsRibbon.InvalidateLinkedItemControls(NC.A);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 49;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	public static void InvalidateOpenDocumentRequiredControls()
	{
		NC.A.InvalidateControl(XC.A(21848));
		_ = null;
	}
}
