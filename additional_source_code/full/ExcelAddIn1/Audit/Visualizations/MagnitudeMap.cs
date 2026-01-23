using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class MagnitudeMap
{
	public static void Apply()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		Range range = default(Range);
		bool flag = default(bool);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		Range range2 = default(Range);
		wpfMagnitudeMap wpfMagnitudeMap2 = default(wpfMagnitudeMap);
		Range range3 = default(Range);
		Range rng = default(Range);
		IEnumerator enumerator2 = default(IEnumerator);
		Range rng2 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
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
				case 1013:
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
							goto IL_0039;
						case 6:
							goto IL_0057;
						case 7:
							goto IL_0063;
						case 8:
							goto IL_008b;
						case 9:
							goto IL_009e;
						case 10:
							goto IL_00ac;
						case 11:
							goto IL_00c0;
						case 12:
							goto IL_00ec;
						case 14:
							goto IL_00fc;
						case 15:
							goto IL_0107;
						case 16:
							goto IL_0111;
						case 17:
							goto IL_012d;
						case 18:
							goto IL_0137;
						case 19:
							goto IL_0143;
						case 20:
							goto IL_016d;
						case 22:
							goto IL_017c;
						case 23:
							goto IL_01a5;
						case 24:
							goto IL_01ca;
						case 25:
							goto IL_01d6;
						case 26:
							goto IL_01dd;
						case 27:
							goto IL_01e7;
						case 28:
							goto IL_01ed;
						case 29:
							goto IL_0205;
						case 31:
							goto IL_022f;
						case 32:
							goto IL_0254;
						case 33:
							goto IL_0262;
						case 34:
							goto IL_0273;
						case 35:
							goto IL_027d;
						case 36:
							goto IL_0283;
						case 37:
							goto IL_0299;
						case 21:
						case 30:
						case 38:
							goto IL_02b1;
						case 39:
							goto IL_02c2;
						case 40:
							goto IL_02d8;
						case 41:
							goto IL_02e2;
						case 42:
							goto IL_02f6;
						case 43:
							goto IL_02f9;
						case 45:
							goto IL_0301;
						case 13:
						case 44:
						case 46:
							goto IL_0315;
						case 47:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 48:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0299:
					num2 = 37;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_02b1;
					IL_02f9:
					num2 = 43;
					range = null;
					goto IL_0315;
					IL_02b1:
					num2 = 38;
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
						goto IL_02c2;
					}
					goto IL_02d8;
					IL_0301:
					num2 = 45;
					Forms.InfoMessage(VH.A(51505));
					goto IL_0315;
					IL_0021:
					num2 = 3;
					application = MH.A.Application;
					goto IL_0032;
					IL_0032:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0039;
					IL_0039:
					num2 = 5;
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
					goto IL_0057;
					IL_02c2:
					num2 = 39;
					JH.A(range, VH.A(51478));
					goto IL_02d8;
					IL_0057:
					num2 = 6;
					range2 = JH.A((Range)null);
					goto IL_0063;
					IL_0063:
					num2 = 7;
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
						goto IL_008b;
					}
					goto IL_009e;
					IL_02d8:
					num2 = 40;
					application.ScreenUpdating = true;
					goto IL_02e2;
					IL_008b:
					num2 = 8;
					range2 = application.ActiveCell.CurrentRegion;
					goto IL_009e;
					IL_009e:
					num2 = 9;
					range = ExcelAddIn1.Formulas.Helpers.CellsWithNumbers(range2);
					goto IL_00ac;
					IL_00ac:
					num2 = 10;
					if (range != null)
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
						goto IL_00c0;
					}
					goto IL_0301;
					IL_02e2:
					num2 = 41;
					Common.A(VH.A(51478));
					goto IL_02f6;
					IL_00c0:
					num2 = 11;
					if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, KH.A.UndoMaxCells, TextCompare: false))
					{
						goto IL_00ec;
					}
					goto IL_00fc;
					IL_00ec:
					num2 = 12;
					if (Common.A() != DialogResult.Cancel)
					{
						goto IL_00fc;
					}
					goto IL_0315;
					IL_00fc:
					num2 = 14;
					Common.A(range2, Common.VisualizationType.MagnitudeHeatmap);
					goto IL_0107;
					IL_0107:
					num2 = 15;
					wpfMagnitudeMap2 = new wpfMagnitudeMap();
					goto IL_0111;
					IL_0111:
					num2 = 16;
					if (wpfMagnitudeMap2.ShowDialog() == true)
					{
						goto IL_012d;
					}
					goto IL_02f6;
					IL_012d:
					num2 = 17;
					application.ScreenUpdating = false;
					goto IL_0137;
					IL_0137:
					num2 = 18;
					flag = JH.A(range);
					goto IL_0143;
					IL_0143:
					num2 = 19;
					if (wpfMagnitudeMap2.radAll.IsChecked == true)
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
						goto IL_016d;
					}
					goto IL_017c;
					IL_02f6:
					wpfMagnitudeMap2 = null;
					goto IL_02f9;
					IL_016d:
					num2 = 20;
					A(range);
					goto IL_02b1;
					IL_017c:
					num2 = 22;
					if (wpfMagnitudeMap2.radRows.IsChecked == true)
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
						goto IL_01a5;
					}
					goto IL_022f;
					IL_0254:
					num2 = 32;
					range3 = ExcelAddIn1.Formulas.Helpers.CellsWithNumbers(rng);
					goto IL_0262;
					IL_01a5:
					num2 = 23;
					enumerator2 = range2.Rows.GetEnumerator();
					goto IL_01f0;
					IL_01f0:
					if (enumerator2.MoveNext())
					{
						rng2 = (Range)enumerator2.Current;
						goto IL_01ca;
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
					goto IL_0205;
					IL_0262:
					num2 = 33;
					if (range3 != null)
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
						goto IL_0273;
					}
					goto IL_0283;
					IL_0205:
					num2 = 29;
					if (enumerator2 is IDisposable)
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
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_02b1;
					IL_0273:
					num2 = 34;
					A(range3);
					goto IL_027d;
					IL_027d:
					num2 = 35;
					range3 = null;
					goto IL_0283;
					IL_0315:
					num2 = 46;
					application = null;
					break;
					IL_01ca:
					num2 = 24;
					range3 = ExcelAddIn1.Formulas.Helpers.CellsWithNumbers(rng2);
					goto IL_01d6;
					IL_01d6:
					num2 = 25;
					if (range3 != null)
					{
						goto IL_01dd;
					}
					goto IL_01ed;
					IL_01dd:
					num2 = 26;
					A(range3);
					goto IL_01e7;
					IL_01e7:
					num2 = 27;
					range3 = null;
					goto IL_01ed;
					IL_01ed:
					num2 = 28;
					goto IL_01f0;
					IL_022f:
					num2 = 31;
					enumerator = range2.Columns.GetEnumerator();
					goto IL_0286;
					IL_0286:
					if (enumerator.MoveNext())
					{
						rng = (Range)enumerator.Current;
						goto IL_0254;
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
					goto IL_0299;
					IL_0283:
					num2 = 36;
					goto IL_0286;
					end_IL_0000_3:
					break;
				}
				num2 = 47;
				range2 = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1013;
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

	private static void A(Range A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		double num5 = default(double);
		double num6 = default(double);
		double num7 = default(double);
		IEnumerator enumerator = default(IEnumerator);
		Range range = default(Range);
		double num8 = default(double);
		bool flag = default(bool);
		bool flag2 = default(bool);
		IEnumerator enumerator2 = default(IEnumerator);
		Range range2 = default(Range);
		Interior interior = default(Interior);
		double c = default(double);
		Color c2 = default(Color);
		Interior interior2 = default(Interior);
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
				case 912:
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
							goto IL_0014;
						case 4:
							goto IL_0021;
						case 5:
							goto IL_003e;
						case 6:
							goto IL_0051;
						case 7:
							goto IL_0066;
						case 8:
							goto IL_007b;
						case 9:
							goto IL_009c;
						case 10:
							goto IL_00a2;
						case 11:
							goto IL_00bb;
						case 12:
							goto IL_00c1;
						case 13:
							goto IL_00dc;
						case 14:
							goto IL_00f4;
						case 15:
							goto IL_00fe;
						case 16:
							goto IL_0113;
						case 17:
							goto IL_0121;
						case 18:
							goto IL_0143;
						case 19:
							goto IL_0155;
						case 20:
							goto IL_016e;
						case 21:
							goto IL_017c;
						case 22:
							goto IL_0190;
						case 23:
							goto IL_01ab;
						case 25:
							goto IL_01b3;
						case 26:
							goto IL_01c3;
						case 28:
							goto IL_01d3;
						case 27:
						case 29:
							goto IL_0201;
						case 30:
							goto IL_0210;
						case 32:
							goto IL_0237;
						case 33:
							goto IL_0246;
						case 31:
						case 34:
							goto IL_026b;
						case 35:
							goto IL_0279;
						case 36:
							goto IL_028d;
						case 37:
							goto IL_02a5;
						case 24:
						case 38:
							goto IL_02a8;
						case 39:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 40:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0051:
					num2 = 6;
					num5 = Math.Min(Math.Abs(num6), num5);
					goto IL_0066;
					IL_0007:
					num2 = 2;
					num5 = 10000000000.0;
					goto IL_0014;
					IL_0014:
					num2 = 3;
					num7 = -10000000000.0;
					goto IL_0021;
					IL_0021:
					num2 = 4;
					enumerator = A.GetEnumerator();
					goto IL_00c4;
					IL_00c4:
					if (enumerator.MoveNext())
					{
						range = (Range)enumerator.Current;
						goto IL_003e;
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
					goto IL_00dc;
					IL_003e:
					num2 = 5;
					num6 = Conversions.ToDouble(range.Value2);
					goto IL_0051;
					IL_00dc:
					num2 = 13;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00f4;
					IL_0066:
					num2 = 7;
					num7 = Math.Max(Math.Abs(num6), num7);
					goto IL_007b;
					IL_00f4:
					num2 = 14;
					num8 = num7 - num5;
					goto IL_00fe;
					IL_00fe:
					num2 = 15;
					if (flag)
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
						if (flag2)
						{
							goto IL_0113;
						}
					}
					goto IL_0121;
					IL_007b:
					num2 = 8;
					if (num6 > 0.0)
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
						goto IL_009c;
					}
					goto IL_00a2;
					IL_00a2:
					num2 = 10;
					if (num6 < 0.0)
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
						goto IL_00bb;
					}
					goto IL_00c1;
					IL_0113:
					num2 = 16;
					num5 = 0.0;
					goto IL_0121;
					IL_0121:
					num2 = 17;
					enumerator2 = A.GetEnumerator();
					goto IL_02ab;
					IL_02ab:
					if (enumerator2.MoveNext())
					{
						range2 = (Range)enumerator2.Current;
						goto IL_0143;
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
					break;
					IL_00c1:
					num2 = 12;
					goto IL_00c4;
					IL_00bb:
					num2 = 11;
					flag2 = true;
					goto IL_00c1;
					IL_0143:
					num2 = 18;
					num6 = Conversions.ToDouble(range2.Value2);
					goto IL_0155;
					IL_0155:
					num2 = 19;
					if (num6 == 0.0)
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
						goto IL_016e;
					}
					goto IL_01b3;
					IL_009c:
					num2 = 9;
					flag = true;
					goto IL_00a2;
					IL_016e:
					num2 = 20;
					interior = range2.Interior;
					goto IL_017c;
					IL_017c:
					num2 = 21;
					interior.Pattern = XlPattern.xlPatternGray75;
					goto IL_0190;
					IL_0190:
					num2 = 22;
					interior.PatternColor = ColorTranslator.ToOle(Color.White);
					goto IL_01ab;
					IL_01ab:
					interior = null;
					goto IL_02a8;
					IL_01b3:
					num2 = 25;
					if (num8 == 0.0)
					{
						goto IL_01c3;
					}
					goto IL_01d3;
					IL_01c3:
					num2 = 26;
					c = 0.5;
					goto IL_0201;
					IL_01d3:
					num2 = 28;
					c = Math.Round(100.0 * (Math.Abs(num6) - num5) / num8) / 100.0;
					goto IL_0201;
					IL_0201:
					num2 = 29;
					if (num6 > 0.0)
					{
						goto IL_0210;
					}
					goto IL_0237;
					IL_0210:
					num2 = 30;
					c2 = Common.A(Color.White, Color.FromArgb(74, 184, 103), c);
					goto IL_026b;
					IL_0237:
					num2 = 32;
					if (num6 < 0.0)
					{
						goto IL_0246;
					}
					goto IL_026b;
					IL_0246:
					num2 = 33;
					c2 = Common.A(Color.White, Color.FromArgb(238, 82, 85), c);
					goto IL_026b;
					IL_026b:
					num2 = 34;
					interior2 = range2.Interior;
					goto IL_0279;
					IL_0279:
					num2 = 35;
					interior2.Pattern = XlPattern.xlPatternGray75;
					goto IL_028d;
					IL_028d:
					num2 = 36;
					interior2.PatternColor = ColorTranslator.ToOle(c2);
					goto IL_02a5;
					IL_02a5:
					interior2 = null;
					goto IL_02a8;
					IL_02a8:
					num2 = 38;
					goto IL_02ab;
					end_IL_0000_2:
					break;
				}
				num2 = 39;
				if (enumerator2 is IDisposable)
				{
					(enumerator2 as IDisposable).Dispose();
				}
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 912;
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
			switch (3)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}
