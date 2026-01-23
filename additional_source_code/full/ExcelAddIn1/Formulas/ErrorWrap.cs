using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class ErrorWrap
{
	[CompilerGenerated]
	private static string m_A;

	private static string LastPromptValue
	{
		[CompilerGenerated]
		get
		{
			return ErrorWrap.m_A;
		}
		[CompilerGenerated]
		set
		{
			ErrorWrap.m_A = value;
		}
	} = VH.A(118884);

	public static void Toggle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
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
			Regex regex = new Regex(VH.A(160122), RegexOptions.IgnoreCase);
			bool flag = false;
			bool flag2 = false;
			Range range = null;
			Range range2 = JH.A((Range)null);
			if (range2 != null)
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
				if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
					if (Conversions.ToBoolean(range2.HasFormula))
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
						range = range2;
					}
				}
				else
				{
					try
					{
						range = range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
			}
			if (range == null)
			{
				range2 = null;
				return;
			}
			Application application = MH.A.Application;
			Application application2 = application;
			application2.ScreenUpdating = false;
			application2.EnableEvents = false;
			application2.DisplayAlerts = false;
			XlCalculation calculation = application2.Calculation;
			application2.Calculation = XlCalculation.xlCalculationManual;
			_ = null;
			bool flag3 = JH.A(range);
			try
			{
				enumerator = range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range3 = (Range)enumerator.Current;
					if (range3.Formula.ToString().Contains(VH.A(160189)))
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
						range3.Formula = regex.Replace(range3.Formula.ToString(), VH.A(160210));
					}
					range3 = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_018e;
					}
					continue;
					end_IL_018e:
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
			try
			{
				bool flag4 = Conversions.ToBoolean(application.ActiveCell.HasFormula);
				if (flag4)
				{
					goto IL_02dc;
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
				flag2 = !NewLateBinding.LateGet(range.Cells[1, 1], null, VH.A(68956), new object[0], null, null, null).ToString().Contains(VH.A(118889));
				wpfErrorWrap wpfErrorWrap2 = new wpfErrorWrap();
				wpfErrorWrap2.optAdd.IsChecked = flag2;
				wpfErrorWrap2.optRemove.IsChecked = !flag2;
				wpfErrorWrap2.ShowDialog();
				if (wpfErrorWrap2.DialogResult.HasValue)
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
					if (wpfErrorWrap2.DialogResult.Value)
					{
						flag2 = wpfErrorWrap2.optAdd.IsChecked.Value;
						wpfErrorWrap2 = null;
						goto IL_02dc;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_02ad;
						}
						continue;
						end_IL_02ad:
						break;
					}
				}
				goto end_IL_01bb;
				IL_0327:
				string text;
				if (!KH.A.ErrorValuePrompt)
				{
					text = KH.A.DefaultErrorValue;
					if (Operators.CompareString(text, null, TextCompare: false) == 0)
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
						text = "";
					}
					goto IL_03a6;
				}
				text = Forms.InputBox(VH.A(118768), VH.A(118803), LastPromptValue);
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
				{
					LastPromptValue = text;
					goto IL_03a6;
				}
				goto end_IL_01bb;
				IL_03a6:
				if (!Versioned.IsNumeric(text))
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
					text = VH.A(39830) + text + VH.A(39830);
					flag = true;
				}
				Regex regex2 = new Regex(VH.A(118904));
				try
				{
					enumerator2 = range.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Range range4 = (Range)enumerator2.Current;
						try
						{
							if (!NewLateBinding.LateGet(range4, null, VH.A(1998), new object[0], null, null, null).ToString().Contains(VH.A(118889)))
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									string text2 = regex2.Match(Conversions.ToString(NewLateBinding.LateGet(range4, null, VH.A(1998), new object[0], null, null, null))).ToString();
									string text3 = regex2.Replace(Conversions.ToString(NewLateBinding.LateGet(range4, null, VH.A(1998), new object[0], null, null, null)), "");
									NewLateBinding.LateSet(range4, null, VH.A(1998), new object[1] { text2 + VH.A(79125) + text3 + VH.A(2378) + text + VH.A(39904) }, null, null);
									if (flag)
									{
										range4.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
									}
									break;
								}
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							if (!range4.Formula.ToString().Contains(VH.A(118889)))
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
								string text2 = regex2.Match(Conversions.ToString(range4.Formula)).ToString();
								string text3 = regex2.Replace(Conversions.ToString(range4.Formula), "");
								range4.Formula = text2 + VH.A(79125) + text3 + VH.A(2378) + text + VH.A(39904);
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
									range4.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
								}
							}
							ProjectData.ClearProjectError();
						}
						range4 = null;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				regex2 = null;
				if (flag3)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						JH.A(range, VH.A(160239));
						break;
					}
				}
				goto end_IL_01bb;
				IL_02dc:
				if (flag2)
				{
					goto IL_0327;
				}
				if (flag4)
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
					if (!application.ActiveCell.Formula.ToString().Contains(VH.A(118889)))
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
						goto IL_0327;
					}
				}
				enumerator3 = range.GetEnumerator();
				try
				{
					while (enumerator3.MoveNext())
					{
						Range range5 = (Range)enumerator3.Current;
						try
						{
							if (NewLateBinding.LateGet(range5, null, VH.A(1998), new object[0], null, null, null).ToString().Contains(VH.A(118889)))
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									string text3 = Conversions.ToString(NewLateBinding.LateGet(range5, null, VH.A(1998), new object[0], null, null, null));
									while (text3.Contains(VH.A(79125)))
									{
										text3 = A(text3);
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										NewLateBinding.LateSet(range5, null, VH.A(1998), new object[1] { text3 }, null, null);
										break;
									}
									break;
								}
							}
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							if (range5.Formula.ToString().Contains(VH.A(118889)))
							{
								string text3 = Conversions.ToString(range5.Formula);
								while (text3.Contains(VH.A(79125)))
								{
									text3 = A(text3);
								}
								range5.Formula = text3;
							}
							ProjectData.ClearProjectError();
						}
						range5 = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_07d4;
						}
						continue;
						end_IL_07d4:
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator3 as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
				if (flag3)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						JH.A(range, VH.A(160239));
						break;
					}
				}
				end_IL_01bb:;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			Common.RefreshLiveVisualizations(range2);
			Application application3 = application;
			application3.DisplayAlerts = true;
			application3.Calculation = calculation;
			application3.EnableEvents = true;
			application3.ScreenUpdating = true;
			_ = null;
			application = null;
			range2 = null;
			range = null;
			regex = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(160239));
			return;
		}
	}

	private static string A(string A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string text = default(string);
		string text2 = default(string);
		int num5 = default(int);
		int num6 = default(int);
		int num7 = default(int);
		int length = default(int);
		int num9 = default(int);
		int num10 = default(int);
		int num11 = default(int);
		bool flag = default(bool);
		bool flag2 = default(bool);
		bool flag3 = default(bool);
		string result = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num8;
					char c;
					switch (try0000_dispatch)
					{
					default:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0007;
					case 640:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000;
							}
							int num4 = unchecked(num + 1);
							num = 0;
							switch (num4)
							{
							case 1:
								break;
							case 2:
								goto IL_0007;
							case 3:
								goto IL_000c;
							case 4:
								goto IL_0015;
							case 5:
								goto IL_002b;
							case 6:
								goto IL_0038;
							case 7:
								goto IL_0047;
							case 9:
								goto IL_00a7;
							case 10:
								goto IL_00d7;
							case 11:
								goto IL_00de;
							case 12:
								goto IL_00e9;
							case 13:
								goto IL_00ef;
							case 14:
								goto IL_00f5;
							case 17:
								goto IL_0100;
							case 19:
								goto IL_010b;
							case 20:
								goto IL_0114;
							case 21:
								goto IL_0133;
							case 24:
								goto IL_013f;
							case 26:
								goto IL_014b;
							case 27:
								goto IL_0165;
							case 8:
							case 16:
							case 18:
							case 23:
							case 25:
							case 28:
							case 29:
								goto IL_016b;
							case 30:
								goto IL_0180;
							case 15:
							case 31:
								goto IL_019c;
							case 22:
							case 32:
								goto IL_01b8;
							case 33:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 34:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_01b8:
						num2 = 32;
						text = Strings.Replace(text2, Strings.Mid(text2, num5, num6 - num5 + 1), text);
						break;
						IL_0007:
						num2 = 2;
						text2 = A;
						goto IL_000c;
						IL_000c:
						num2 = 3;
						text = "";
						goto IL_0015;
						IL_0015:
						num2 = 4;
						num7 = Strings.InStr(text2, VH.A(79125));
						goto IL_002b;
						IL_002b:
						num2 = 5;
						length = text2.Length;
						goto IL_0038;
						IL_0038:
						num2 = 6;
						num8 = num7;
						num9 = length;
						num10 = num8;
						goto IL_01a5;
						IL_01a5:
						if (num10 <= num9)
						{
							goto IL_0047;
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
						goto IL_01b8;
						IL_0100:
						num2 = 17;
						num11++;
						goto IL_016b;
						IL_0047:
						num2 = 7;
						c = Strings.GetChar(text2, num10);
						if (c != '"')
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
							switch (c)
							{
							case 'I':
								break;
							case '(':
								goto IL_0100;
							case ')':
								goto IL_010b;
							case '\'':
								goto IL_013f;
							case ',':
								goto IL_014b;
							default:
								goto IL_016b;
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
							goto IL_00a7;
						}
						goto IL_013f;
						IL_00de:
						num2 = 11;
						num10 = num10 + 8 - 1;
						goto IL_00e9;
						IL_00a7:
						num2 = 9;
						if (Operators.CompareString(Strings.Mid(text2, num10, 8), VH.A(79125), TextCompare: false) == 0)
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
							goto IL_00d7;
						}
						goto IL_016b;
						IL_00d7:
						num2 = 10;
						num5 = num10;
						goto IL_00de;
						IL_00e9:
						num2 = 12;
						flag = true;
						goto IL_00ef;
						IL_014b:
						num2 = 26;
						if (flag)
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
							if (num11 == 1 && !flag2)
							{
								goto IL_0165;
							}
						}
						goto IL_016b;
						IL_00ef:
						num2 = 13;
						flag3 = false;
						goto IL_00f5;
						IL_00f5:
						num2 = 14;
						num11 = 1;
						goto IL_019c;
						IL_013f:
						num2 = 24;
						flag2 = !flag2;
						goto IL_016b;
						IL_0165:
						num2 = 27;
						flag3 = true;
						goto IL_016b;
						IL_010b:
						num2 = 19;
						num11--;
						goto IL_0114;
						IL_0114:
						num2 = 20;
						if (flag3)
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
							if (num11 == 0)
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
								goto IL_0133;
							}
						}
						goto IL_016b;
						IL_016b:
						num2 = 29;
						if (flag)
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
							if (!flag3)
							{
								goto IL_0180;
							}
						}
						goto IL_019c;
						IL_0180:
						num2 = 30;
						text += Conversions.ToString(Strings.GetChar(text2, num10));
						goto IL_019c;
						IL_019c:
						num2 = 31;
						num10++;
						goto IL_01a5;
						IL_0133:
						num2 = 21;
						num6 = num10;
						goto IL_01b8;
						end_IL_0000_2:
						break;
					}
					num2 = 33;
					result = text;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 640;
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
				switch (5)
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
}
