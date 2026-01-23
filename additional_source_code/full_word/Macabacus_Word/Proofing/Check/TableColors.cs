using System;
using System.Linq;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class TableColors
{
	private enum BC
	{
		A = 0,
		B = 255,
		C = 128,
		D = 208,
		E = 223
	}

	private struct CC
	{
		public BC A;

		public WdThemeColorIndex A;

		public double A;

		public int A;
	}

	private struct DC
	{
		public double A;

		public double B;

		public double C;
	}

	public static int Colours2(WdColor wdColor)
	{
		CC cC = A((int)wdColor);
		BC a = cC.A;
		if (a <= BC.C)
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
			if (a != BC.A)
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
				if (a != BC.C)
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
					goto IL_01a3;
				}
				Interaction.MsgBox(XC.A(23050));
			}
			else
			{
				Interaction.MsgBox(XC.A(22856));
			}
		}
		else if (a != BC.D)
		{
			if (a != BC.B)
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
				goto IL_01a3;
			}
			Interaction.MsgBox(XC.A(22957));
		}
		else
		{
			Interaction.MsgBox(XC.A(23171) + A(cC.A, MsoLanguageID.msoLanguageIDEnglishUS) + A(cC.A) + XC.A(23248) + Conversions.ToString((long)cC.A & 0xFFL) + XC.A(22698) + Conversions.ToString((double)((long)cC.A & 0xFF00L) / 256.0) + XC.A(22698) + Conversions.ToString((double)(cC.A & 0xFF0000) / 65536.0) + XC.A(20696));
		}
		goto IL_01b9;
		IL_01b9:
		return cC.A;
		IL_01a3:
		Interaction.MsgBox(XC.A(23305));
		goto IL_01b9;
	}

	private static CC A(int A)
	{
		CC result = default(CC);
		string text = Strings.Right(new string('0', 7) + Conversion.Hex(A), 8);
		byte b = Conversions.ToByte(XC.A(23418) + Strings.Left(text, 2));
		Interaction.MsgBox(b);
		byte b2 = b;
		if (b2 == 0)
		{
			result.A = BC.A;
		}
		else if (b2 == byte.MaxValue)
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
			result.A = BC.B;
		}
		else if (b2 == 128)
		{
			result.A = BC.C;
		}
		else
		{
			if ((uint)b2 >= 208u && (uint)b2 <= 223u)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return TableColors.A(b, text);
					}
				}
			}
			result.A = (BC)b;
		}
		return result;
	}

	private static CC A(byte A, string B)
	{
		CC result = default(CC);
		byte b = Conversions.ToByte(XC.A(23418) + Strings.Mid(B, 7, 2));
		byte b2 = Conversions.ToByte(XC.A(23418) + Strings.Mid(B, 5, 2));
		result.A = (BC)(A & 0xF0);
		result.A = (WdThemeColorIndex)(A & 0xF);
		if (b2 != byte.MaxValue)
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
			result.A = Math.Round(-1.0 + (double)(int)b2 / 255.0, 2);
		}
		if (b != byte.MaxValue)
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
			result.A = Math.Round(1.0 - (double)(int)b / 255.0, 2);
		}
		result.A = Conversions.ToInteger(TableColors.A(result.A, result.A));
		return result;
	}

	private static string A(WdThemeColorIndex A, double B)
	{
		MsoThemeColorSchemeIndex index = TableColors.A(A);
		DC a = TableColors.A(Conversions.ToInteger(NewLateBinding.LateGet(PC.A.Application.ActiveDocument.DocumentTheme.ThemeColorScheme.Cast<object>().ElementAtOrDefault((int)index), null, XC.A(22408), new object[0], null, null, null)));
		a.C = a.C * Math.Abs(B) + (double)(0 - ((B > 0.0) ? 1 : 0)) * (1.0 - B);
		return Conversions.ToString(TableColors.B(a));
	}

	private static MsoThemeColorSchemeIndex A(WdThemeColorIndex A)
	{
		MsoThemeColorSchemeIndex msoThemeColorSchemeIndex = default(MsoThemeColorSchemeIndex);
		return A switch
		{
			WdThemeColorIndex.wdThemeColorMainDark1 => MsoThemeColorSchemeIndex.msoThemeDark1, 
			WdThemeColorIndex.wdThemeColorMainLight1 => MsoThemeColorSchemeIndex.msoThemeLight1, 
			WdThemeColorIndex.wdThemeColorMainDark2 => MsoThemeColorSchemeIndex.msoThemeDark2, 
			WdThemeColorIndex.wdThemeColorMainLight2 => MsoThemeColorSchemeIndex.msoThemeLight2, 
			WdThemeColorIndex.wdThemeColorAccent1 => MsoThemeColorSchemeIndex.msoThemeAccent1, 
			WdThemeColorIndex.wdThemeColorAccent2 => MsoThemeColorSchemeIndex.msoThemeAccent2, 
			WdThemeColorIndex.wdThemeColorAccent3 => MsoThemeColorSchemeIndex.msoThemeAccent3, 
			WdThemeColorIndex.wdThemeColorAccent4 => MsoThemeColorSchemeIndex.msoThemeAccent4, 
			WdThemeColorIndex.wdThemeColorAccent5 => MsoThemeColorSchemeIndex.msoThemeAccent5, 
			WdThemeColorIndex.wdThemeColorAccent6 => MsoThemeColorSchemeIndex.msoThemeAccent6, 
			WdThemeColorIndex.wdThemeColorHyperlink => MsoThemeColorSchemeIndex.msoThemeHyperlink, 
			WdThemeColorIndex.wdThemeColorHyperlinkFollowed => MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink, 
			WdThemeColorIndex.wdThemeColorBackground1 => MsoThemeColorSchemeIndex.msoThemeLight1, 
			WdThemeColorIndex.wdThemeColorText1 => MsoThemeColorSchemeIndex.msoThemeDark1, 
			WdThemeColorIndex.wdThemeColorBackground2 => MsoThemeColorSchemeIndex.msoThemeLight2, 
			WdThemeColorIndex.wdThemeColorText2 => MsoThemeColorSchemeIndex.msoThemeDark2, 
			_ => msoThemeColorSchemeIndex, 
		};
	}

	private static DC A(int A)
	{
		DC result = default(DC);
		string str = Strings.Right(new string('0', 7) + Conversion.Hex(A), 8);
		double num = Conversions.ToDouble(XC.A(23418) + Strings.Mid(str, 7, 2)) / 255.0;
		double num2 = Conversions.ToDouble(XC.A(23418) + Strings.Mid(str, 5, 2)) / 255.0;
		double num3 = Conversions.ToDouble(XC.A(23418) + Strings.Mid(str, 3, 2)) / 255.0;
		double num4 = num;
		if (num2 > num4)
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
			num4 = num2;
		}
		if (num3 > num4)
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
			num4 = num3;
		}
		double num5 = num;
		if (num2 < num5)
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
			num5 = num2;
		}
		if (num3 < num5)
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
			num5 = num3;
		}
		double num6 = num4 - num5;
		result.C = (num4 + num5) / 2.0;
		if (num6 == 0.0)
		{
			result.B = 0.0;
			result.A = 0.0;
		}
		else
		{
			double num7 = num4;
			if (num7 == num)
			{
				result.A = 1.0 / 6.0 * (num2 - num3) / num6 - (double)(0 - ((num3 > num2) ? 1 : 0));
			}
			else if (num7 == num2)
			{
				result.A = 1.0 / 6.0 * (num3 - num) / num6 + 1.0 / 3.0;
			}
			else if (num7 == num3)
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
				result.A = 1.0 / 6.0 * (num - num2) / num6 + 2.0 / 3.0;
			}
			if (result.C < 0.5)
			{
				result.B = num6 / (2.0 * result.C);
			}
			else
			{
				result.B = num6 / (2.0 - 2.0 * result.C);
			}
		}
		return result;
	}

	private static int B(DC A)
	{
		double num;
		double num2;
		double num3;
		if (A.B == 0.0)
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
			num = A.C;
			num2 = A.C;
			num3 = A.C;
		}
		else
		{
			double num4;
			if (A.C < 0.5)
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
				num4 = A.C * (1.0 + A.B);
			}
			else
			{
				num4 = A.C + A.B - A.C * A.B;
			}
			double b = 2.0 * A.C - num4;
			num = TableColors.A(num4, b, Conversions.ToDouble(Interaction.IIf(A.A > 2.0 / 3.0, A.A - 2.0 / 3.0, A.A + 1.0 / 3.0)));
			num2 = TableColors.A(num4, b, A.A);
			num3 = TableColors.A(num4, b, Conversions.ToDouble(Interaction.IIf(A.A < 1.0 / 3.0, A.A + 2.0 / 3.0, A.A - 1.0 / 3.0)));
		}
		return checked((int)Conversions.ToLong(XC.A(23423) + Strings.Right(XC.A(23432) + Conversion.Hex(Math.Round(num3 * 255.0)), 2) + Strings.Right(XC.A(23432) + Conversion.Hex(Math.Round(num2 * 255.0)), 2) + Strings.Right(XC.A(23432) + Conversion.Hex(Math.Round(num * 255.0)), 2)));
	}

	private static double A(double A, double B, double C)
	{
		if (C < 1.0 / 6.0)
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
					return B + (A - B) * 6.0 * C;
				}
			}
		}
		if (C < 0.5)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return A;
				}
			}
		}
		if (C < 2.0 / 3.0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return B + (A - B) * (2.0 / 3.0 - C) * 6.0;
				}
			}
		}
		return B;
	}

	private static string A(WdThemeColorIndex A, MsoLanguageID B = MsoLanguageID.msoLanguageIDEnglishUS)
	{
		if (B == MsoLanguageID.msoLanguageIDNone)
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
			B = (MsoLanguageID)PC.A.Application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
		}
		MsoLanguageID msoLanguageID = B;
		if (msoLanguageID != MsoLanguageID.msoLanguageIDFrench)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (msoLanguageID != MsoLanguageID.msoLanguageIDDutch)
					{
						return A switch
						{
							WdThemeColorIndex.wdThemeColorMainDark1 => XC.A(24179), 
							WdThemeColorIndex.wdThemeColorMainLight1 => XC.A(24192), 
							WdThemeColorIndex.wdThemeColorMainDark2 => XC.A(24207), 
							WdThemeColorIndex.wdThemeColorMainLight2 => XC.A(24220), 
							WdThemeColorIndex.wdThemeColorAccent1 => XC.A(23499), 
							WdThemeColorIndex.wdThemeColorAccent2 => XC.A(23516), 
							WdThemeColorIndex.wdThemeColorAccent3 => XC.A(23533), 
							WdThemeColorIndex.wdThemeColorAccent4 => XC.A(23550), 
							WdThemeColorIndex.wdThemeColorAccent5 => XC.A(23567), 
							WdThemeColorIndex.wdThemeColorAccent6 => XC.A(23584), 
							WdThemeColorIndex.wdThemeColorHyperlink => XC.A(23601), 
							WdThemeColorIndex.wdThemeColorHyperlinkFollowed => XC.A(24235), 
							WdThemeColorIndex.wdThemeColorBackground1 => XC.A(24272), 
							WdThemeColorIndex.wdThemeColorText1 => XC.A(24297), 
							WdThemeColorIndex.wdThemeColorBackground2 => XC.A(24310), 
							WdThemeColorIndex.wdThemeColorText2 => XC.A(24335), 
							_ => XC.A(24348) + Conversions.ToString((int)A), 
						};
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							return A switch
							{
								WdThemeColorIndex.wdThemeColorMainDark1 => XC.A(23435), 
								WdThemeColorIndex.wdThemeColorMainLight1 => XC.A(23452), 
								WdThemeColorIndex.wdThemeColorMainDark2 => XC.A(23467), 
								WdThemeColorIndex.wdThemeColorMainLight2 => XC.A(23484), 
								WdThemeColorIndex.wdThemeColorAccent1 => XC.A(23499), 
								WdThemeColorIndex.wdThemeColorAccent2 => XC.A(23516), 
								WdThemeColorIndex.wdThemeColorAccent3 => XC.A(23533), 
								WdThemeColorIndex.wdThemeColorAccent4 => XC.A(23550), 
								WdThemeColorIndex.wdThemeColorAccent5 => XC.A(23567), 
								WdThemeColorIndex.wdThemeColorAccent6 => XC.A(23584), 
								WdThemeColorIndex.wdThemeColorHyperlink => XC.A(23601), 
								WdThemeColorIndex.wdThemeColorHyperlinkFollowed => XC.A(23620), 
								WdThemeColorIndex.wdThemeColorBackground1 => XC.A(23657), 
								WdThemeColorIndex.wdThemeColorText1 => XC.A(23684), 
								WdThemeColorIndex.wdThemeColorBackground2 => XC.A(23699), 
								WdThemeColorIndex.wdThemeColorText2 => XC.A(23726), 
								_ => XC.A(23741) + Conversions.ToString((int)A), 
							};
						}
					}
				}
			}
		}
		return A switch
		{
			WdThemeColorIndex.wdThemeColorMainDark1 => XC.A(23760), 
			WdThemeColorIndex.wdThemeColorMainLight1 => XC.A(23777), 
			WdThemeColorIndex.wdThemeColorMainDark2 => XC.A(23792), 
			WdThemeColorIndex.wdThemeColorMainLight2 => XC.A(23809), 
			WdThemeColorIndex.wdThemeColorAccent1 => XC.A(23824), 
			WdThemeColorIndex.wdThemeColorAccent2 => XC.A(23853), 
			WdThemeColorIndex.wdThemeColorAccent3 => XC.A(23882), 
			WdThemeColorIndex.wdThemeColorAccent4 => XC.A(23911), 
			WdThemeColorIndex.wdThemeColorAccent5 => XC.A(23940), 
			WdThemeColorIndex.wdThemeColorAccent6 => XC.A(23969), 
			WdThemeColorIndex.wdThemeColorHyperlink => XC.A(23998), 
			WdThemeColorIndex.wdThemeColorHyperlinkFollowed => XC.A(24029), 
			WdThemeColorIndex.wdThemeColorBackground1 => XC.A(24074), 
			WdThemeColorIndex.wdThemeColorText1 => XC.A(24103), 
			WdThemeColorIndex.wdThemeColorBackground2 => XC.A(24118), 
			WdThemeColorIndex.wdThemeColorText2 => XC.A(24147), 
			_ => XC.A(24162) + Conversions.ToString((int)A), 
		};
	}

	private static string A(double A, MsoLanguageID B = MsoLanguageID.msoLanguageIDEnglishUS)
	{
		if (B == MsoLanguageID.msoLanguageIDNone)
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
			B = (MsoLanguageID)PC.A.Application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
		}
		if (A == 0.0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return "";
				}
			}
		}
		if (A > 0.0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					MsoLanguageID msoLanguageID = B;
					string text;
					if (msoLanguageID != MsoLanguageID.msoLanguageIDFrench)
					{
						if (msoLanguageID == MsoLanguageID.msoLanguageIDDutch)
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
							text = XC.A(24365);
						}
						else
						{
							text = XC.A(24413);
						}
					}
					else
					{
						text = XC.A(24386);
					}
					return text + Conversions.ToString(A * 100.0) + XC.A(20080);
				}
				}
			}
		}
		if (A < 0.0)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					MsoLanguageID msoLanguageID2 = B;
					string text;
					if (msoLanguageID2 != MsoLanguageID.msoLanguageIDFrench)
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
						if (msoLanguageID2 == MsoLanguageID.msoLanguageIDDutch)
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
							text = XC.A(24434);
						}
						else
						{
							text = XC.A(24488);
						}
					}
					else
					{
						text = XC.A(24459);
					}
					return text + Conversions.ToString(A * -100.0) + XC.A(20080);
				}
				}
			}
		}
		string result = default(string);
		return result;
	}
}
