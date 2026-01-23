using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Footnotes
{
	[CompilerGenerated]
	private static int m_A;

	private static readonly string m_A = VH.A(147328);

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Footnotes.m_A;
		}
		[CompilerGenerated]
		set
		{
			Footnotes.m_A = value;
		}
	}

	public static void Toggle()
	{
		Application application = MH.A.Application;
		Range range = null;
		Regex regex = FootnoteRegex();
		MatchCollection matchCollection;
		Range range2;
		try
		{
			if (application.Selection is Range)
			{
				application.ScreenUpdating = false;
				range2 = A(application);
				{
					IEnumerator enumerator = range2.GetEnumerator();
					try
					{
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Range range3 = (Range)enumerator.Current;
								if (!regex.IsMatch(Conversions.ToString(range3.Text)))
								{
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
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								if (!application.WorksheetFunction.IsText(range3))
								{
									continue;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									range = range3;
									break;
								}
								break;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_00a7;
								}
								continue;
								end_IL_00a7:
								break;
							}
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
				if (range != null)
				{
					IEnumerator enumerator3 = default(IEnumerator);
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						bool flag = JH.A(range2);
						if (Operators.ConditionalCompareObjectEqual(range.get_Characters((object)Strings.Len(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), (object)1).Font.Superscript, true, TextCompare: false))
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
							range2.Font.Superscript = false;
						}
						else
						{
							foreach (Range item in range2)
							{
								if (!application.WorksheetFunction.IsText(item))
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
									break;
								}
								Range range4 = item;
								float num = Conversions.ToSingle(range4.Height);
								range4.Font.Superscript = false;
								string text = Conversions.ToString(range4.Text);
								matchCollection = regex.Matches(text);
								try
								{
									enumerator3 = matchCollection.GetEnumerator();
									while (enumerator3.MoveNext())
									{
										string text2 = NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(enumerator3.Current), null, VH.A(146974), new object[1] { 1 }, null, null, null).ToString();
										range4.get_Characters((object)checked(text.Length - text2.Length + 1), (object)text2.Length).Font.Superscript = true;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_0270;
										}
										continue;
										end_IL_0270:
										break;
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								range4.RowHeight = num;
								range4 = null;
							}
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
								if (KH.A.UndoFont)
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
									JH.A(range2, VH.A(146987));
								}
							}
						}
						A(VH.A(147004));
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
		application.ScreenUpdating = true;
		application = null;
		matchCollection = null;
		range = null;
		range2 = null;
		regex = null;
	}

	public static void Show()
	{
		A(A: true);
		A(VH.A(147035));
	}

	public static void Hide()
	{
		A(A: false);
		A(VH.A(147064));
	}

	private static void A(bool A)
	{
		Application application = MH.A.Application;
		Regex regex = FootnoteRegex();
		if (application.Selection is Range)
		{
			Range range = Footnotes.A(application);
			application.ScreenUpdating = false;
			try
			{
				if (A)
				{
					IEnumerator enumerator = default(IEnumerator);
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						try
						{
							enumerator = range.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Range range2 = (Range)enumerator.Current;
								if (!CellContainsFootnote(range2, regex))
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
									break;
								}
								range2.Font.Color = RuntimeHelpers.GetObjectValue(range2.get_Characters((object)1, (object)1).Font.Color);
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
						break;
					}
				}
				else
				{
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = range.GetEnumerator();
						IEnumerator enumerator3 = default(IEnumerator);
						while (enumerator2.MoveNext())
						{
							Range range3 = (Range)enumerator2.Current;
							if (!CellContainsFootnote(range3, regex))
							{
								continue;
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
							object objectValue = RuntimeHelpers.GetObjectValue(range3.Interior.Color);
							string text = range3.Text.ToString();
							MatchCollection matchCollection = regex.Matches(text);
							try
							{
								enumerator3 = matchCollection.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									int length = ((Match)enumerator3.Current).Groups[1].ToString().Length;
									range3.get_Characters((object)checked(text.Length - length + 1), (object)length).Font.Color = RuntimeHelpers.GetObjectValue(objectValue);
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										(enumerator3 as IDisposable).Dispose();
										break;
									}
								}
							}
							matchCollection = null;
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
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			range = null;
		}
		application = null;
		regex = null;
	}

	private static Range A(Application A)
	{
		Range range = (Range)A.Selection;
		if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
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
			range = JH.A((Range)null).SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		return range;
	}

	public static bool CellContainsFootnote(Range rng, Regex re)
	{
		int num;
		if (re.IsMatch(rng.Text.ToString()))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (rng.Application.WorksheetFunction.IsText(rng))
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
				num = (Conversions.ToBoolean(Operators.CompareObjectEqual(rng.get_Characters((object)rng.Value2.ToString().Length, (object)1).Font.Superscript, true, TextCompare: false)) ? 1 : 0);
				goto IL_009b;
			}
		}
		num = 0;
		goto IL_009b;
		IL_009b:
		return Conversions.ToBoolean((byte)num != 0);
	}

	public static Regex FootnoteRegex()
	{
		return new Regex(VH.A(147093), RegexOptions.IgnoreCase);
	}

	public static void CheckSequence()
	{
		if (!A())
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
			new wpfFootnoteSequence().ShowDialog();
			_ = null;
			A(VH.A(147231));
			return;
		}
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, A);
	}

	public static void Cycle()
	{
		if (!A())
		{
			return;
		}
		Application application = MH.A.Application;
		checked
		{
			if (application.Selection is Range)
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
				if (CycleIndex == 0)
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
					if (Regex.IsMatch(application.ActiveCell.NumberFormat.ToString(), Footnotes.m_A))
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
						A(0);
						A(VH.A(147266));
						goto IL_00e8;
					}
				}
				if (CycleIndex == 9)
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
					CycleIndex = 0;
				}
				else
				{
					CycleIndex++;
				}
				A(CycleIndex);
				if (CycleIndex == 1)
				{
					A(VH.A(147266));
				}
			}
			goto IL_00e8;
		}
		IL_00e8:
		application = null;
	}

	public static void DoFootnote(IRibbonControl control)
	{
		A(Conversions.ToInteger(control.Tag));
	}

	private static void A(int A)
	{
		Application application = MH.A.Application;
		string text = "";
		try
		{
			if (application.Selection is Range)
			{
				Range range = Footnotes.A(application);
				if (!Base.IsWorksheetProtected(range.Worksheet))
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
					switch (A)
					{
					case 1:
						text = VH.A(147295);
						break;
					case 2:
						text = VH.A(147298);
						break;
					case 3:
						text = VH.A(147301);
						break;
					case 4:
						text = VH.A(147304);
						break;
					case 5:
						text = VH.A(147307);
						break;
					case 6:
						text = VH.A(147310);
						break;
					case 7:
						text = VH.A(147313);
						break;
					case 8:
						text = VH.A(147316);
						break;
					case 9:
						text = VH.A(147319);
						break;
					}
					if (A > 0)
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
						text = VH.A(147322) + text + VH.A(147325);
					}
					bool flag = JH.A(range);
					application.ScreenUpdating = false;
					try
					{
						IEnumerator enumerator = range.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Range range2 = (Range)enumerator.Current;
								string text2 = Conversions.ToString(range2.NumberFormat);
								if (A > 0)
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
									if (Operators.CompareString(text2, VH.A(20593), TextCompare: false) != 0)
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
										if (Operators.CompareString(text2, "", TextCompare: false) != 0)
										{
											text2 = Footnotes.A(text2);
											string[] array = text2.Split(';');
											List<string> list = new List<string>();
											string[] array2 = array;
											foreach (string text3 in array2)
											{
												list.Add(text3 + text);
											}
											range2.NumberFormat = string.Join(VH.A(77635), list.ToArray());
											list = null;
											continue;
										}
									}
									if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.Value2)))
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
										range2.NumberFormat = VH.A(64024) + text;
									}
									else
									{
										range2.NumberFormat = VH.A(48146) + text;
									}
								}
								else
								{
									range2.NumberFormat = Footnotes.A(text2);
								}
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_02a2;
								}
								continue;
								end_IL_02a2:
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
						NumberFormat.HandleException(ex2);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
					if (flag)
					{
						JH.A(range, VH.A(147266));
					}
				}
				range = null;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static string A(string A)
	{
		return Regex.Replace(A, Footnotes.m_A, "");
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
