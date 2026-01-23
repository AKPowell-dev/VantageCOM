using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class DependencyDensity
{
	public static void Apply()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		long num5 = default(long);
		int num6 = default(int);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		List<long> list = default(List<long>);
		Range range = default(Range);
		Range range2 = default(Range);
		Range range3 = default(Range);
		IEnumerator enumerator = default(IEnumerator);
		bool flag = default(bool);
		int num7 = default(int);
		bool flag2 = default(bool);
		int num8 = default(int);
		bool flag3 = default(bool);
		IEnumerator enumerator2 = default(IEnumerator);
		double c = default(double);
		Color c2 = default(Color);
		Interior interior = default(Interior);
		bool flag4 = default(bool);
		Range range4 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						if (!Licensing.AllowVisualizationOperation())
						{
							goto end_IL_0000;
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_001f;
					case 1688:
						{
							num = num2;
							switch (num3)
							{
							case 2:
								break;
							case 1:
								goto IL_0538;
							default:
								goto end_IL_0000_2;
							}
							goto IL_043f;
						}
						IL_0538:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_001f;
						case 4:
							goto IL_0030;
						case 5:
							goto IL_0039;
						case 6:
							goto IL_0040;
						case 8:
							goto IL_005e;
						case 9:
							goto IL_006a;
						case 10:
							goto IL_0095;
						case 11:
							goto IL_00a9;
						case 12:
							goto IL_00d5;
						case 14:
							goto IL_00e5;
						case 15:
							goto IL_00ef;
						case 16:
							goto IL_00f9;
						case 17:
							goto IL_0106;
						case 18:
							goto IL_0125;
						case 19:
							goto IL_013e;
						case 20:
							goto IL_0144;
						case 21:
							goto IL_014a;
						case 22:
						case 23:
						case 24:
						case 34:
						case 40:
							goto IL_0150;
						case 25:
							goto IL_0157;
						case 26:
							goto IL_0176;
						case 28:
							goto IL_0195;
						case 29:
							goto IL_019c;
						case 31:
							goto IL_0217;
						case 32:
							goto IL_021d;
						case 33:
							goto IL_0223;
						case 27:
						case 30:
						case 35:
							goto IL_0231;
						case 37:
							goto IL_0242;
						case 38:
							goto IL_0248;
						case 39:
							goto IL_024e;
						case 36:
						case 41:
							goto IL_025c;
						case 42:
							goto IL_026d;
						case 43:
							goto IL_027c;
						case 44:
							goto IL_028d;
						case 45:
							goto IL_02af;
						case 46:
							goto IL_02d7;
						case 47:
							goto IL_02e1;
						case 48:
							goto IL_02ec;
						case 49:
							goto IL_02f2;
						case 50:
							goto IL_02fc;
						case 51:
							goto IL_030a;
						case 52:
							goto IL_032d;
						case 53:
							goto IL_0342;
						case 54:
							goto IL_036c;
						case 55:
							goto IL_0388;
						case 56:
							goto IL_0396;
						case 57:
							goto IL_03aa;
						case 58:
							goto IL_03c0;
						case 60:
							goto IL_03c5;
						case 61:
							goto IL_03d2;
						case 59:
						case 62:
							goto IL_03d4;
						case 63:
							goto IL_03dd;
						case 64:
							goto IL_03f8;
						case 65:
							goto IL_041a;
						case 66:
							goto IL_0421;
						case 68:
							goto IL_0439;
						case 67:
						case 69:
							goto IL_043f;
						case 70:
							goto IL_0446;
						case 71:
							goto IL_044d;
						case 72:
							goto IL_0493;
						case 73:
							goto IL_04be;
						case 74:
							goto IL_04c9;
						case 75:
							goto IL_04cc;
						case 76:
							goto IL_04d7;
						case 77:
							goto IL_04e1;
						case 78:
							goto IL_04eb;
						case 79:
							goto IL_04fc;
						case 13:
						case 80:
							goto IL_0510;
						case 81:
							goto IL_0515;
						case 82:
							goto IL_051b;
						case 83:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 7:
						case 84:
							goto end_IL_0000;
						}
						goto default;
						IL_0176:
						num2 = 26;
						if (Information.Err().Number <= 0)
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
							goto IL_0195;
						}
						goto IL_0231;
						IL_025c:
						num2 = 41;
						num5 = Math.Max(num5, num6);
						goto IL_026d;
						IL_0195:
						ProjectData.ClearProjectError();
						num3 = 0;
						goto IL_019c;
						IL_001f:
						num2 = 3;
						application = MH.A.Application;
						goto IL_0030;
						IL_0030:
						num2 = 4;
						list = new List<long>();
						goto IL_0039;
						IL_0039:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_0040;
						IL_0040:
						num2 = 6;
						if (!(application.Selection is Range))
						{
							goto end_IL_0000;
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
						goto IL_005e;
						IL_019c:
						num2 = 29;
						if (Operators.CompareString(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), application.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
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
							goto IL_0217;
						}
						goto IL_0231;
						IL_005e:
						num2 = 8;
						range2 = JH.A((Range)null);
						goto IL_006a;
						IL_006a:
						num2 = 9;
						if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
							goto IL_0095;
						}
						goto IL_00a9;
						IL_026d:
						num2 = 42;
						list.Add(num6 - 1);
						goto IL_027c;
						IL_0095:
						num2 = 10;
						range2 = application.ActiveCell.CurrentRegion;
						goto IL_00a9;
						IL_00a9:
						num2 = 11;
						if (Operators.ConditionalCompareObjectGreater(range2.Cells.CountLarge, KH.A.UndoMaxCells, TextCompare: false))
						{
							goto IL_00d5;
						}
						goto IL_00e5;
						IL_00d5:
						num2 = 12;
						if (Common.A() != DialogResult.Cancel)
						{
							goto IL_00e5;
						}
						goto IL_0510;
						IL_00e5:
						num2 = 14;
						application.ScreenUpdating = false;
						goto IL_00ef;
						IL_00ef:
						num2 = 15;
						application.EnableEvents = false;
						goto IL_00f9;
						IL_00f9:
						num2 = 16;
						range3 = application.ActiveCell;
						goto IL_0106;
						IL_0106:
						num2 = 17;
						enumerator = range2.GetEnumerator();
						goto IL_027f;
						IL_027f:
						if (enumerator.MoveNext())
						{
							range = (Range)enumerator.Current;
							goto IL_0125;
						}
						goto IL_028d;
						IL_028d:
						num2 = 44;
						if (enumerator is IDisposable)
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
							(enumerator as IDisposable).Dispose();
						}
						goto IL_02af;
						IL_021d:
						num2 = 32;
						flag = false;
						goto IL_0223;
						IL_0223:
						num2 = 33;
						num7++;
						goto IL_0150;
						IL_02af:
						num2 = 45;
						NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
						goto IL_02d7;
						IL_02d7:
						num2 = 46;
						if (flag2)
						{
							goto IL_02e1;
						}
						goto IL_0439;
						IL_02e1:
						num2 = 47;
						Common.A(range2, Common.VisualizationType.DependencyDensity);
						goto IL_02ec;
						IL_02ec:
						num2 = 48;
						num8 = 0;
						goto IL_02f2;
						IL_02f2:
						num2 = 49;
						num5--;
						goto IL_02fc;
						IL_02fc:
						num2 = 50;
						flag3 = JH.A(range2);
						goto IL_030a;
						IL_030a:
						num2 = 51;
						enumerator2 = range2.GetEnumerator();
						goto IL_03e0;
						IL_03e0:
						if (enumerator2.MoveNext())
						{
							range = (Range)enumerator2.Current;
							goto IL_032d;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						goto IL_03f8;
						IL_0231:
						num2 = 35;
						if (!flag)
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
							goto IL_0242;
						}
						goto IL_025c;
						IL_03f8:
						num2 = 64;
						if (enumerator2 is IDisposable)
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
							(enumerator2 as IDisposable).Dispose();
						}
						goto IL_041a;
						IL_0242:
						num2 = 37;
						num7 = 1;
						goto IL_0248;
						IL_0248:
						num2 = 38;
						flag = true;
						goto IL_024e;
						IL_041a:
						num2 = 65;
						if (flag3)
						{
							goto IL_0421;
						}
						goto IL_043f;
						IL_0421:
						num2 = 66;
						JH.A(range2, VH.A(51125));
						goto IL_043f;
						IL_027c:
						num2 = 43;
						goto IL_027f;
						IL_032d:
						num2 = 52;
						if (list[num8] > 0)
						{
							goto IL_0342;
						}
						goto IL_03c5;
						IL_0342:
						num2 = 53;
						c = Math.Round((double)(100 * list[num8]) / (double)num5) / 100.0;
						goto IL_036c;
						IL_036c:
						num2 = 54;
						c2 = Common.A(Color.White, Color.DarkOrange, c);
						goto IL_0388;
						IL_0388:
						num2 = 55;
						interior = range.Interior;
						goto IL_0396;
						IL_0396:
						num2 = 56;
						interior.Pattern = XlPattern.xlPatternGray75;
						goto IL_03aa;
						IL_03aa:
						num2 = 57;
						interior.PatternColor = ColorTranslator.ToOle(c2);
						goto IL_03c0;
						IL_03c0:
						interior = null;
						goto IL_03d4;
						IL_03c5:
						num2 = 60;
						_ = range.Interior;
						goto IL_03d2;
						IL_03d2:
						_ = null;
						goto IL_03d4;
						IL_03d4:
						num2 = 62;
						num8++;
						goto IL_03dd;
						IL_03dd:
						num2 = 63;
						goto IL_03e0;
						IL_0439:
						num2 = 68;
						flag4 = true;
						goto IL_043f;
						IL_043f:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0446;
						IL_0446:
						num2 = 70;
						range4 = range2;
						goto IL_044d;
						IL_044d:
						num2 = 71;
						NewLateBinding.LateCall(NewLateBinding.LateGet(range4.Parent, null, VH.A(8701), new object[0], null, null, null), null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
						goto IL_0493;
						IL_0493:
						num2 = 72;
						NewLateBinding.LateCall(range4.Parent, null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
						goto IL_04be;
						IL_04be:
						num2 = 73;
						range4.Select();
						goto IL_04c9;
						IL_04c9:
						range4 = null;
						goto IL_04cc;
						IL_04cc:
						num2 = 75;
						range3.Activate();
						goto IL_04d7;
						IL_04d7:
						num2 = 76;
						application.ScreenUpdating = true;
						goto IL_04e1;
						IL_04e1:
						num2 = 77;
						application.EnableEvents = true;
						goto IL_04eb;
						IL_04eb:
						num2 = 78;
						if (flag4)
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
							goto IL_04fc;
						}
						goto IL_0510;
						IL_024e:
						num2 = 39;
						num6++;
						goto IL_0150;
						IL_04fc:
						num2 = 79;
						Forms.InfoMessage(VH.A(40093));
						goto IL_0510;
						IL_0510:
						num2 = 80;
						application = null;
						goto IL_0515;
						IL_0515:
						num2 = 81;
						range2 = null;
						goto IL_051b;
						IL_051b:
						num2 = 82;
						range3 = null;
						break;
						IL_0217:
						num2 = 31;
						flag2 = true;
						goto IL_021d;
						IL_0125:
						num2 = 18;
						range.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_013e;
						IL_013e:
						num2 = 19;
						num6 = 1;
						goto IL_0144;
						IL_0144:
						num2 = 20;
						num7 = 1;
						goto IL_014a;
						IL_014a:
						num2 = 21;
						flag = true;
						goto IL_0150;
						IL_0150:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0157;
						IL_0157:
						num2 = 25;
						range.NavigateArrow(false, num6, num7);
						goto IL_0176;
						end_IL_0000_3:
						break;
					}
					num2 = 83;
					Common.A(VH.A(51125));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1688;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}
}
