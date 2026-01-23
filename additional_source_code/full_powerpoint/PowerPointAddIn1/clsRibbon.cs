using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.Aiwa.UI;
using MacabacusMacros.LogoLibrary.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Pagination;
using PowerPointAddIn1.Publishing.Share;
using PowerPointAddIn1.Shapes.Arrange;
using PowerPointAddIn1.Shapes.SelectMatch;

namespace PowerPointAddIn1;

public sealed class clsRibbon
{
	public static bool CallbackView(IRibbonControl control)
	{
		bool result = default(bool);
		try
		{
			DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
			if (activeWindow.View.Type == PpViewType.ppViewNormal)
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
				PpViewType viewType = activeWindow.Panes[2].ViewType;
				if (viewType != PpViewType.ppViewSlide)
				{
					if (viewType != PpViewType.ppViewNormal)
					{
						goto IL_0097;
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
				}
				result = true;
			}
			else
			{
				if (!((activeWindow.View.Type == PpViewType.ppViewSlideSorter) & (Operators.CompareString(control.Id, AH.A(152081), TextCompare: false) == 0)))
				{
					goto IL_0097;
				}
				result = true;
			}
			goto end_IL_0000;
			IL_0097:
			activeWindow = null;
			end_IL_0000:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool CallbackSlideView(bool ShowWarning = false)
	{
		return clsPowerPoint.IsNormalView(NG.A.Application, ShowWarning);
	}

	public static bool CallbackIsGrouped()
	{
		return HasShapeType(MsoShapeType.msoGroup);
	}

	public static bool CallbackNotGrouped()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		bool result = default(bool);
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
					goto IL_0007;
				case 105:
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
							goto IL_0007;
						case 3:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0007:
					num2 = 2;
					if (NG.A.Application.ActiveWindow.Selection.ShapeRange.Count <= 1)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 3;
				result = true;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 105;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool HasShapeType(MsoShapeType vType)
	{
		bool result = default(bool);
		try
		{
			IEnumerator enumerator = NG.A.Application.ActiveWindow.Selection.ShapeRange.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					if (((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current).Type != vType)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						result = true;
						return result;
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_005b;
					}
					continue;
					end_IL_005b:
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public void GroupControlsReset()
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
					goto IL_0007;
				case 87:
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
							goto IL_0007;
						case 3:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0007:
					num2 = 2;
					KG.A.InvalidateControl(AH.A(152100));
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 3;
				KG.A.InvalidateControl(AH.A(152115));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 87;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public void InvalidateTab()
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
				case 61:
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
				KG.A.InvalidateControl(AH.A(152130));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 61;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void InvalidateLinkedItemControls()
	{
		clsRibbon.InvalidateLinkedItemControls(KG.A);
	}

	public static Bitmap RecolorColorButton(string id)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Bitmap bitmap2 = default(Bitmap);
		string left = default(string);
		Color color = default(Color);
		Bitmap result = default(Bitmap);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				Bitmap bitmap;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 408:
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
							goto IL_0007;
						case 4:
							goto IL_000c;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_0040;
						case 8:
							goto IL_0057;
						case 9:
							goto IL_0074;
						case 10:
							goto IL_00bb;
						case 12:
							goto IL_00d0;
						case 13:
							goto IL_00e4;
						case 14:
							goto IL_0121;
						case 3:
						case 7:
						case 11:
						case 15:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 16:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0074:
					num2 = 9;
					if (!(PB.Settings.LastFillColor == Color.Transparent))
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
						bitmap = new Bitmap(OB.FillColorPicker);
					}
					else
					{
						bitmap = new Bitmap(OB.NoFill);
					}
					bitmap2 = bitmap;
					goto IL_00bb;
					IL_0007:
					num2 = 2;
					left = id;
					goto IL_000c;
					IL_000c:
					num2 = 4;
					if (Operators.CompareString(left, clsColors.LAST_FONT_COLOR_BUTTON, TextCompare: false) == 0)
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
						goto IL_0030;
					}
					goto IL_0057;
					IL_00d0:
					num2 = 12;
					if (Operators.CompareString(left, clsColors.LAST_BORDER_COLOR_BUTTON, TextCompare: false) != 0)
					{
						break;
					}
					goto IL_00e4;
					IL_0121:
					num2 = 14;
					color = PB.Settings.LastBorderColor;
					break;
					IL_00e4:
					num2 = 13;
					bitmap2 = ((PB.Settings.LastBorderColor == Color.Transparent) ? new Bitmap(OB.NoBorder) : new Bitmap(OB.BorderColorPicker));
					goto IL_0121;
					IL_0030:
					num2 = 5;
					bitmap2 = new Bitmap(OB.FontColorPicker);
					goto IL_0040;
					IL_0040:
					num2 = 6;
					color = PB.Settings.LastFontColor;
					break;
					IL_0057:
					num2 = 8;
					if (Operators.CompareString(left, clsColors.LAST_FILL_COLOR_BUTTON, TextCompare: false) == 0)
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
						goto IL_0074;
					}
					goto IL_00d0;
					IL_00bb:
					num2 = 10;
					color = PB.Settings.LastFillColor;
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 15;
				result = clsColors.RecolorColorButton(bitmap2, color);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 408;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static string GenerateLabel(string strName, List<string> listUsedKeytips)
	{
		string result = strName;
		checked
		{
			int num = strName.Length - 1;
			int num2 = 0;
			while (true)
			{
				if (num2 <= num)
				{
					string text = strName.Substring(num2, 1).ToUpper();
					if (Regex.IsMatch(text, AH.A(152149)))
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
						if (!listUsedKeytips.Contains(text))
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
							result = clsRibbon.FixAmpersand(strName.Substring(0, num2)) + AH.A(82543) + clsRibbon.FixAmpersand(strName.Substring(num2));
							listUsedKeytips.Add(text);
							break;
						}
					}
					num2++;
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
				break;
			}
			return result;
		}
	}

	public static void Application_WindowSelectionChange(Selection Sel)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		IRibbonUI ribbonUI = default(IRibbonUI);
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
					goto IL_0007;
				case 705:
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
							goto IL_0007;
						case 3:
							goto IL_000f;
						case 4:
							goto IL_001c;
						case 5:
							goto IL_0029;
						case 6:
							goto IL_0036;
						case 7:
							goto IL_004a;
						case 8:
							goto IL_005e;
						case 9:
							goto IL_0072;
						case 10:
							goto IL_0087;
						case 11:
							goto IL_009a;
						case 12:
							goto IL_00ad;
						case 13:
							goto IL_00c2;
						case 14:
							goto IL_00d7;
						case 15:
							goto IL_00ec;
						case 16:
							goto IL_0101;
						case 17:
							goto IL_0116;
						case 18:
							goto IL_0129;
						case 19:
							goto IL_013e;
						case 20:
							goto IL_0151;
						case 21:
							goto IL_0166;
						case 22:
							goto IL_0179;
						case 23:
							goto IL_018c;
						case 24:
							goto IL_01a1;
						case 25:
							goto IL_01b6;
						case 26:
							goto IL_01cb;
						case 27:
							goto IL_01e0;
						case 28:
							goto IL_01e2;
						case 29:
							goto IL_01ea;
						case 30:
							goto IL_020e;
						case 31:
							goto IL_0217;
						case 32:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 33:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_020e:
					num2 = 30;
					Text.SelectionChange(Sel);
					goto IL_0217;
					IL_0007:
					num2 = 2;
					ribbonUI = KG.A;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					ribbonUI.InvalidateControl(clsColors.LAST_FONT_COLOR_BUTTON);
					goto IL_001c;
					IL_001c:
					num2 = 4;
					ribbonUI.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
					goto IL_0029;
					IL_0029:
					num2 = 5;
					ribbonUI.InvalidateControl(clsColors.LAST_BORDER_COLOR_BUTTON);
					goto IL_0036;
					IL_0036:
					num2 = 6;
					ribbonUI.InvalidateControl(AH.A(152160));
					goto IL_004a;
					IL_004a:
					num2 = 7;
					ribbonUI.InvalidateControl(AH.A(152193));
					goto IL_005e;
					IL_005e:
					num2 = 8;
					ribbonUI.InvalidateControl(AH.A(152226));
					goto IL_0072;
					IL_0072:
					num2 = 9;
					ribbonUI.InvalidateControl(AH.A(152263));
					goto IL_0087;
					IL_0087:
					num2 = 10;
					ribbonUI.InvalidateControl(AH.A(152290));
					goto IL_009a;
					IL_009a:
					num2 = 11;
					ribbonUI.InvalidateControl(AH.A(152317));
					goto IL_00ad;
					IL_00ad:
					num2 = 12;
					ribbonUI.InvalidateControl(AH.A(152348));
					goto IL_00c2;
					IL_00c2:
					num2 = 13;
					ribbonUI.InvalidateControl(AH.A(152369));
					goto IL_00d7;
					IL_00d7:
					num2 = 14;
					ribbonUI.InvalidateControl(AH.A(152406));
					goto IL_00ec;
					IL_00ec:
					num2 = 15;
					ribbonUI.InvalidateControl(AH.A(75110));
					goto IL_0101;
					IL_0101:
					num2 = 16;
					ribbonUI.InvalidateControl(AH.A(68483));
					goto IL_0116;
					IL_0116:
					num2 = 17;
					ribbonUI.InvalidateControl(AH.A(152445));
					goto IL_0129;
					IL_0129:
					num2 = 18;
					ribbonUI.InvalidateControl(AH.A(152486));
					goto IL_013e;
					IL_013e:
					num2 = 19;
					ribbonUI.InvalidateControl(AH.A(152523));
					goto IL_0151;
					IL_0151:
					num2 = 20;
					ribbonUI.InvalidateControl(AH.A(152560));
					goto IL_0166;
					IL_0166:
					num2 = 21;
					ribbonUI.InvalidateControl(AH.A(152603));
					goto IL_0179;
					IL_0179:
					num2 = 22;
					ribbonUI.InvalidateControl(AH.A(152640));
					goto IL_018c;
					IL_018c:
					num2 = 23;
					ribbonUI.InvalidateControl(AH.A(152675));
					goto IL_01a1;
					IL_01a1:
					num2 = 24;
					ribbonUI.InvalidateControl(AH.A(152700));
					goto IL_01b6;
					IL_01b6:
					num2 = 25;
					ribbonUI.InvalidateControl(AH.A(152729));
					goto IL_01cb;
					IL_01cb:
					num2 = 26;
					ribbonUI.InvalidateControl(AH.A(152750));
					goto IL_01e0;
					IL_01e0:
					ribbonUI = null;
					goto IL_01e2;
					IL_01e2:
					num2 = 28;
					Ribbon.ResetSelectionType();
					goto IL_01ea;
					IL_01ea:
					num2 = 29;
					if (KG.A.TextLinkCompatibilityMode)
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
						goto IL_020e;
					}
					goto IL_0217;
					IL_0217:
					num2 = 31;
					PowerPointAddIn1.Links.Hyperlinks.SelectionChange(Sel);
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 32;
				InvalidateLinkedItemControls();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 705;
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

	public static void Application_PresentationCloseFinal(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		InvalidateOpenPresentationRequiredControls();
	}

	public static void Application_WindowActivate(Microsoft.Office.Interop.PowerPoint.Presentation Pres, DocumentWindow Wn)
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
				case 61:
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
				KG.A.InvalidateControl(AH.A(152773));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 61;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void InvalidateOpenPresentationRequiredControls()
	{
		IRibbonUI a = KG.A;
		a.InvalidateControl(AH.A(70518));
		a.InvalidateControl(PowerPointAddIn1.Pagination.Pane.RIBBON_CONTROL);
		a.InvalidateControl(PowerPointAddIn1.Publishing.Share.Pane.A);
		a.InvalidateControl(Pane.RIBBON_CONTROL);
		a.InvalidateControl(Pane.RIBBON_CONTROL);
		a.InvalidateControl(PowerPointAddIn1.Shapes.SelectMatch.Pane.A);
		a.InvalidateControl(PowerPointAddIn1.Shapes.Arrange.Pane.A);
		a.InvalidateControl(AH.A(152816));
		a.InvalidateControl(AH.A(152837));
		a.InvalidateControl(AH.A(152858));
		a.InvalidateControl(AH.A(152879));
		a.InvalidateControl(AH.A(152906));
		a.InvalidateControl(AH.A(152939));
		a.InvalidateControl(AH.A(94212));
		a.InvalidateControl(AH.A(152970));
		a.InvalidateControl(AH.A(152993));
		a.InvalidateControl(AH.A(153026));
		a.InvalidateControl(AH.A(153051));
		a.InvalidateControl(AH.A(153074));
		a.InvalidateControl(AH.A(153099));
		a.InvalidateControl(AH.A(153120));
		a.InvalidateControl(AH.A(153141));
		_ = null;
	}
}
