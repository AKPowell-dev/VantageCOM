using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class AutoColor
{
	private static int m_A = 0;

	public static void Selection()
	{
		Range range;
		try
		{
			range = JH.A((Range)null);
			bool num = JH.A(range);
			Range(range);
			if (num)
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
					if (KH.A.UndoFont)
					{
						JH.A(range, VH.A(60635));
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
		range = null;
	}

	public static void Worksheet()
	{
		Application application = MH.A.Application;
		if (application.ActiveSheet is Worksheet)
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
			application.ScreenUpdating = false;
			Range usedRange;
			try
			{
				usedRange = ((Worksheet)application.ActiveSheet).UsedRange;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = usedRange.GetEnumerator();
					while (enumerator.MoveNext())
					{
						AutoColorIfNotEmpty((Range)enumerator.Current);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_006d;
						}
						continue;
						end_IL_006d:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
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
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			usedRange = null;
			application.ScreenUpdating = true;
		}
		application = null;
	}

	public static void Range(Range rng)
	{
		Application application = rng.Application;
		checked
		{
			Range range;
			Range activeCell;
			try
			{
				application.ScreenUpdating = false;
				AutoColor.m_A++;
				application.EnableEvents = false;
				range = (Range)application.Selection;
				activeCell = application.ActiveCell;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = rng.GetEnumerator();
					while (enumerator.MoveNext())
					{
						AutoColorIfNotEmpty((Range)enumerator.Current);
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				range.Select();
				activeCell.Activate();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range = null;
			activeCell = null;
			AutoColor.m_A--;
			if (AutoColor.m_A < 1)
			{
				application.ScreenUpdating = true;
			}
			application.EnableEvents = true;
			application = null;
		}
	}

	public static void Workbook()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
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
				Application application = MH.A.Application;
				application.ScreenUpdating = false;
				AutoColor.m_A++;
				try
				{
					try
					{
						enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Worksheet worksheet = (Worksheet)enumerator.Current;
							try
							{
								enumerator2 = worksheet.UsedRange.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									AutoColorIfNotEmpty((Range)enumerator2.Current);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0091;
									}
									continue;
									end_IL_0091:
									break;
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										(enumerator2 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00c5;
							}
							continue;
							end_IL_00c5:
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
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				AutoColor.m_A--;
				if (AutoColor.m_A < 1)
				{
					application.ScreenUpdating = true;
				}
				application = null;
				return;
			}
		}
	}

	public static void AutoColorIfNotEmpty(Range rng)
	{
		if (rng.Value2 == null)
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
			A(rng);
			return;
		}
	}

	private static void A(Range A)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		double clr = default(double);
		int num = default(int);
		int num3 = default(int);
		Range range = default(Range);
		bool bSet = default(bool);
		List<string>.Enumerator enumerator = default(List<string>.Enumerator);
		string current = default(string);
		string input = default(string);
		List<string>.Enumerator enumerator2 = default(List<string>.Enumerator);
		string current2 = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				bool num5;
				int num6;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					clr = 0.0;
					goto IL_000d;
				case 1065:
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
							goto IL_000d;
						case 3:
							goto IL_0014;
						case 5:
							goto IL_0053;
						case 6:
							goto IL_0075;
						case 8:
							goto IL_00a1;
						case 9:
							goto IL_00b8;
						case 10:
							goto IL_00c8;
						case 11:
							goto IL_0119;
						case 12:
							goto IL_012c;
						case 13:
							goto IL_0151;
						case 14:
							goto IL_0184;
						case 16:
							goto IL_0193;
						case 17:
							goto IL_01ab;
						case 18:
							goto IL_01bb;
						case 19:
							goto IL_0208;
						case 21:
							goto IL_0217;
						case 22:
							goto IL_0224;
						case 23:
							goto IL_0231;
						case 24:
							goto IL_0237;
						case 25:
							goto IL_024b;
						case 26:
							goto IL_0266;
						case 27:
							goto IL_029e;
						case 29:
							goto IL_02ad;
						case 30:
							goto IL_0344;
						case 32:
							goto IL_0350;
						case 34:
							goto IL_0369;
						case 28:
						case 31:
						case 33:
						case 35:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
						case 7:
						case 15:
						case 20:
						case 36:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0350:
					num2 = 32;
					range.Font.Color = 0;
					break;
					IL_000d:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0014;
					IL_0014:
					num2 = 3;
					if (!KH.A.AutoColorText)
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
						if (A.Application.WorksheetFunction.IsText(A))
						{
							goto end_IL_0000_3;
						}
					}
					goto IL_0053;
					IL_0208:
					num2 = 19;
					SetColor(A, 5);
					goto end_IL_0000_3;
					IL_0217:
					num2 = 21;
					CheckForOffSheetReferences(A, ref clr, ref bSet);
					goto IL_0224;
					IL_0231:
					num2 = 23;
					range = A;
					goto IL_0237;
					IL_0224:
					num2 = 22;
					CheckForPartialInputs(A, ref clr, ref bSet);
					goto IL_0231;
					IL_0053:
					num2 = 5;
					enumerator = KH.A.NoAutoColorCycle.GetEnumerator();
					goto IL_00a3;
					IL_00a3:
					if (enumerator.MoveNext())
					{
						current = enumerator.Current;
						goto IL_0075;
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
					goto IL_00b8;
					IL_0237:
					num2 = 24;
					if (!bSet)
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
						goto IL_024b;
					}
					goto IL_0369;
					IL_00b8:
					num2 = 9;
					((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
					goto IL_00c8;
					IL_00c8:
					num2 = 10;
					if (Conversions.ToBoolean(Conversions.ToBoolean(A.HasFormula) && Operators.CompareString(KH.A.AutoColors[6].ToString(), "", TextCompare: false) != 0))
					{
						goto IL_0119;
					}
					goto IL_01bb;
					IL_00a1:
					num2 = 8;
					goto IL_00a3;
					IL_0266:
					num2 = 26;
					if (Operators.CompareString(KH.A.AutoColors[2].ToString(), "", TextCompare: false) == 0)
					{
						break;
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
					goto IL_029e;
					IL_024b:
					num2 = 25;
					if (Conversions.ToBoolean(range.HasFormula))
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
						goto IL_0266;
					}
					goto IL_02ad;
					IL_0119:
					num2 = 11;
					input = Conversions.ToString(A.Formula);
					goto IL_012c;
					IL_012c:
					num2 = 12;
					enumerator2 = KH.A.DataFunctions.GetEnumerator();
					goto IL_0196;
					IL_0196:
					if (enumerator2.MoveNext())
					{
						current2 = enumerator2.Current;
						goto IL_0151;
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
					goto IL_01ab;
					IL_0344:
					num2 = 30;
					SetColor(A, 0);
					break;
					IL_01ab:
					num2 = 17;
					((IDisposable)enumerator2/*cast due to .constrained prefix*/).Dispose();
					goto IL_01bb;
					IL_029e:
					num2 = 27;
					SetColor(A, 2);
					break;
					IL_0151:
					num2 = 13;
					if (Regex.IsMatch(input, VH.A(51423) + current2 + VH.A(51444), RegexOptions.IgnoreCase))
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
						goto IL_0184;
					}
					goto IL_0193;
					IL_02ad:
					num2 = 29;
					num5 = (Operators.CompareString(KH.A.AutoColors[0].ToString(), "", TextCompare: false) != 0) & Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))));
					if (KH.A.AutoColorDates)
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
						num6 = (Information.IsDate(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))) ? 1 : 0);
					}
					else
					{
						num6 = 0;
					}
					if (((num5 ? 1u : 0u) | (uint)num6) != 0)
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
						goto IL_0344;
					}
					goto IL_0350;
					IL_0184:
					num2 = 14;
					SetColor(A, 6);
					goto end_IL_0000_3;
					IL_0193:
					num2 = 16;
					goto IL_0196;
					IL_01bb:
					num2 = 18;
					if (Operators.CompareString(KH.A.AutoColors[5].ToString(), "", TextCompare: false) != 0)
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
						if (A.Hyperlinks.Count > 0)
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
							goto IL_0208;
						}
					}
					goto IL_0217;
					IL_0369:
					num2 = 34;
					range.Font.Color = clr;
					break;
					IL_0075:
					num2 = 6;
					if (Operators.ConditionalCompareObjectEqual(clsColors.RGB2Ole(current), A.Font.Color, TextCompare: false))
					{
						goto end_IL_0000_3;
					}
					goto IL_00a1;
					end_IL_0000_2:
					break;
				}
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1065;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void CheckForOffSheetReferences(Range cell, ref double clr, ref bool bSet)
	{
		if (!Conversions.ToBoolean(cell.HasFormula))
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
			string text = Conversions.ToString(cell.Formula);
			if (KH.A.AutoColors[4].Length > 0)
			{
				int num = Strings.InStr(text, VH.A(7120));
				if (num > 1)
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
					int num2 = Strings.InStr(text, VH.A(43340));
					if (num2 > num && Strings.InStr(text, VH.A(7827)) > num2)
					{
						bSet = true;
					}
				}
				if (!bSet)
				{
					bSet = Regex.IsMatch(text, VH.A(148482));
					if (!bSet)
					{
						bSet = Regex.IsMatch(text, VH.A(148507));
					}
				}
				if (bSet)
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
					clr = clsColors.RGB2Ole(KH.A.AutoColors[4]);
				}
			}
			if (bSet)
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
				if (KH.A.AutoColors[3].Length <= 0)
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
					text = Helpers.RemoveExtraneousSheetName(text, cell.Worksheet.Name);
					if (Strings.InStr(text, VH.A(7827)) > 1)
					{
						clr = clsColors.RGB2Ole(KH.A.AutoColors[3]);
						bSet = true;
					}
					return;
				}
			}
		}
	}

	public static void CheckForPartialInputs(Range cell, ref double clr, ref bool bSet)
	{
		int num;
		if (Operators.CompareString(KH.A.AutoColors[1].ToString(), "", TextCompare: false) != 0)
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
			num = (Conversions.ToBoolean(cell.HasFormula) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		if (!Conversions.ToBoolean((byte)num != 0))
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (Helpers.ContainsPartialInput(cell))
			{
				clr = clsColors.RGB2Ole(KH.A.AutoColors[1]);
				bSet = true;
			}
			return;
		}
	}

	public static void SetColor(Range rng, int i)
	{
		rng.Font.Color = clsColors.RGB2Ole(KH.A.AutoColors[i]);
	}
}
