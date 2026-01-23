using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Format;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class FunctionalMap
{
	public static void Apply()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Range range = default(Range);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		double clr = default(double);
		Range range2 = default(Range);
		bool bSet = default(bool);
		bool flag = default(bool);
		IEnumerator enumerator = default(IEnumerator);
		Interior interior = default(Interior);
		Range range3 = default(Range);
		string input = default(string);
		List<string>.Enumerator enumerator2 = default(List<string>.Enumerator);
		string current = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num5;
				bool num6;
				int num7;
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
					goto IL_0021;
				case 1501:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0032;
						case 5:
							goto IL_003f;
						case 6:
							goto IL_0046;
						case 8:
							goto IL_0062;
						case 9:
							goto IL_006e;
						case 10:
							goto IL_00a6;
						case 12:
							goto IL_00b6;
						case 13:
							goto IL_00c1;
						case 14:
							goto IL_00cb;
						case 15:
							goto IL_00d7;
						case 16:
							goto IL_00fa;
						case 17:
							goto IL_0123;
						case 18:
							goto IL_0129;
						case 19:
							goto IL_0137;
						case 20:
							goto IL_019d;
						case 21:
							goto IL_01b0;
						case 22:
							goto IL_01d5;
						case 23:
							goto IL_020f;
						case 25:
							goto IL_021f;
						case 26:
							goto IL_0235;
						case 27:
							goto IL_0245;
						case 28:
							goto IL_028d;
						case 30:
							goto IL_029d;
						case 31:
							goto IL_02ab;
						case 32:
							goto IL_02b9;
						case 33:
							goto IL_02c0;
						case 34:
							goto IL_02d4;
						case 35:
							goto IL_02f3;
						case 36:
							goto IL_0323;
						case 38:
							goto IL_0333;
						case 39:
							goto IL_03ce;
						case 41:
							goto IL_03db;
						case 42:
							goto IL_03e9;
						case 43:
							goto IL_03fd;
						case 44:
							goto IL_0418;
						case 46:
							goto IL_041d;
						case 47:
							goto IL_0438;
						case 37:
						case 40:
						case 45:
						case 48:
							goto IL_0450;
						case 49:
							goto IL_0453;
						case 50:
							goto IL_046e;
						case 51:
							goto IL_0490;
						case 52:
							goto IL_04a1;
						case 24:
						case 29:
						case 53:
							goto IL_04b7;
						case 11:
						case 54:
							goto IL_04c1;
						case 55:
							goto IL_04c6;
						case 56:
							goto IL_04c9;
						case 57:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 7:
						case 58:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_028d:
					num2 = 28;
					A(range, 5);
					goto IL_04b7;
					IL_04b7:
					num2 = 53;
					application.ScreenUpdating = true;
					goto IL_04c1;
					IL_04c1:
					num2 = 54;
					application = null;
					goto IL_04c6;
					IL_04c6:
					num2 = 55;
					goto IL_04c9;
					IL_0021:
					num2 = 3;
					application = MH.A.Application;
					goto IL_0032;
					IL_0032:
					num2 = 4;
					clr = 0.0;
					goto IL_003f;
					IL_003f:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0046;
					IL_0046:
					num2 = 6;
					if (!(application.Selection is Range))
					{
						goto end_IL_0000;
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
					goto IL_0062;
					IL_04c9:
					num2 = 56;
					break;
					IL_0062:
					num2 = 8;
					range2 = JH.A((Range)null);
					goto IL_006e;
					IL_006e:
					num2 = 9;
					if (Operators.ConditionalCompareObjectGreater(range2.Cells.CountLarge, KH.A.UndoMaxCells, TextCompare: false))
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
						goto IL_00a6;
					}
					goto IL_00b6;
					IL_029d:
					num2 = 30;
					AutoColor.CheckForOffSheetReferences(range, ref clr, ref bSet);
					goto IL_02ab;
					IL_00a6:
					num2 = 10;
					if (Common.A() != DialogResult.Cancel)
					{
						goto IL_00b6;
					}
					goto IL_04c1;
					IL_00b6:
					num2 = 12;
					Common.A(range2, Common.VisualizationType.FunctionalMap);
					goto IL_00c1;
					IL_00c1:
					num2 = 13;
					application.ScreenUpdating = false;
					goto IL_00cb;
					IL_00cb:
					num2 = 14;
					flag = JH.A(range2);
					goto IL_00d7;
					IL_00d7:
					num2 = 15;
					enumerator = range2.GetEnumerator();
					goto IL_0456;
					IL_0456:
					if (enumerator.MoveNext())
					{
						range = (Range)enumerator.Current;
						goto IL_00fa;
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
					goto IL_046e;
					IL_02ab:
					num2 = 31;
					AutoColor.CheckForPartialInputs(range, ref clr, ref bSet);
					goto IL_02b9;
					IL_046e:
					num2 = 50;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0490;
					IL_02c0:
					num2 = 33;
					if (!bSet)
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
						goto IL_02d4;
					}
					goto IL_041d;
					IL_03e9:
					num2 = 42;
					interior.Pattern = XlPattern.xlPatternGray75;
					goto IL_03fd;
					IL_0490:
					num2 = 51;
					if (flag)
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
						goto IL_04a1;
					}
					goto IL_04b7;
					IL_02d4:
					num2 = 34;
					if (Conversions.ToBoolean(range3.HasFormula))
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
						goto IL_02f3;
					}
					goto IL_0333;
					IL_04a1:
					num2 = 52;
					JH.A(range2, VH.A(51449));
					goto IL_04b7;
					IL_02b9:
					num2 = 32;
					range3 = range;
					goto IL_02c0;
					IL_00fa:
					num2 = 16;
					if (!string.IsNullOrEmpty(Conversions.ToString(range.Text)))
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
						goto IL_0123;
					}
					goto IL_0453;
					IL_03fd:
					num2 = 43;
					interior.PatternColor = ColorTranslator.ToOle(Color.White);
					goto IL_0418;
					IL_0123:
					num2 = 17;
					bSet = false;
					goto IL_0129;
					IL_0129:
					num2 = 18;
					clr = 0.0;
					goto IL_0137;
					IL_0137:
					num2 = 19;
					if (Conversions.ToBoolean(range.HasFormula))
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
						num5 = ((Operators.CompareString(KH.A.AutoColors[6].ToString(), "", TextCompare: false) != 0) ? 1 : 0);
					}
					else
					{
						num5 = 0;
					}
					if (Conversions.ToBoolean((byte)num5 != 0))
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
						goto IL_019d;
					}
					goto IL_0245;
					IL_0323:
					num2 = 36;
					A(range, 2);
					goto IL_0450;
					IL_041d:
					num2 = 46;
					range3.Interior.Pattern = XlPattern.xlPatternGray50;
					goto IL_0438;
					IL_0333:
					num2 = 38;
					num6 = (Operators.CompareString(KH.A.AutoColors[0].ToString(), "", TextCompare: false) != 0) & Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range3.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
					if (KH.A.AutoColorDates)
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
						num7 = (Information.IsDate(RuntimeHelpers.GetObjectValue(range3.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))) ? 1 : 0);
					}
					else
					{
						num7 = 0;
					}
					if (((num6 ? 1u : 0u) | (uint)num7) != 0)
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
						goto IL_03ce;
					}
					goto IL_03db;
					IL_02f3:
					num2 = 35;
					if (Operators.CompareString(KH.A.AutoColors[2].ToString(), "", TextCompare: false) != 0)
					{
						goto IL_0323;
					}
					goto IL_0450;
					IL_0450:
					range3 = null;
					goto IL_0453;
					IL_019d:
					num2 = 20;
					input = Conversions.ToString(range.Formula);
					goto IL_01b0;
					IL_01b0:
					num2 = 21;
					enumerator2 = KH.A.DataFunctions.GetEnumerator();
					goto IL_0222;
					IL_0222:
					if (enumerator2.MoveNext())
					{
						current = enumerator2.Current;
						goto IL_01d5;
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
					goto IL_0235;
					IL_0438:
					num2 = 47;
					range3.Interior.PatternColor = clr;
					goto IL_0450;
					IL_0235:
					num2 = 26;
					((IDisposable)enumerator2/*cast due to .constrained prefix*/).Dispose();
					goto IL_0245;
					IL_0418:
					interior = null;
					goto IL_0450;
					IL_01d5:
					num2 = 22;
					if (Regex.IsMatch(input, VH.A(51423) + current + VH.A(51444), RegexOptions.IgnoreCase))
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
						goto IL_020f;
					}
					goto IL_021f;
					IL_0453:
					num2 = 49;
					goto IL_0456;
					IL_020f:
					num2 = 23;
					A(range, 6);
					goto IL_04b7;
					IL_021f:
					num2 = 25;
					goto IL_0222;
					IL_0245:
					num2 = 27;
					if (Operators.CompareString(KH.A.AutoColors[5].ToString(), "", TextCompare: false) != 0)
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
						if (range.Hyperlinks.Count > 0)
						{
							goto IL_028d;
						}
					}
					goto IL_029d;
					IL_03ce:
					num2 = 39;
					A(range, 0);
					goto IL_0450;
					IL_03db:
					num2 = 41;
					interior = range3.Interior;
					goto IL_03e9;
					end_IL_0000_3:
					break;
				}
				num2 = 57;
				Common.A(VH.A(51449));
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1501;
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
			switch (2)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static void A(Range A, int B)
	{
		Interior interior = A.Interior;
		interior.Pattern = XlPattern.xlPatternGray50;
		interior.PatternColor = clsColors.RGB2Ole(KH.A.AutoColors[B]);
		_ = null;
	}
}
