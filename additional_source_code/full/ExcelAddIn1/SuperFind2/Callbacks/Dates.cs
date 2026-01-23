using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Dates
{
	internal static void A(WorksheetItem A, Range B)
	{
		B = RangeHelpers.H(B);
		if (B == null)
		{
			return;
		}
		Range A2 = null;
		IEnumerator enumerator = B.Areas.GetEnumerator();
		checked
		{
			try
			{
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
					if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
						try
						{
							if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
								RangeHelpers.A(ref A2, range);
								break;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						continue;
					}
					int num = Information.LBound((Array)objectValue);
					int num2 = Information.UBound((Array)objectValue);
					for (int i = num; i <= num2; i++)
					{
						int num3 = Information.LBound((Array)objectValue, 2);
						int num4 = Information.UBound((Array)objectValue, 2);
						for (int j = num3; j <= num4; j++)
						{
							try
							{
								if (!Information.IsDate(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null))))
								{
									continue;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
									break;
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
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
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_01ac;
					}
					continue;
					end_IL_01ac:
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
			U(A, A2);
			A2 = null;
		}
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, bool>)Dates.A);
	}

	private static bool A(DateTime A, DateTime B)
	{
		return DateTime.Compare(A, B) == 0;
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, bool>)Dates.B);
	}

	private static bool B(DateTime A, DateTime B)
	{
		return DateTime.Compare(A, B) == -1;
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, bool>)C);
	}

	private static bool C(DateTime A, DateTime B)
	{
		if (DateTime.Compare(A, B) != -1)
		{
			return DateTime.Compare(A, B) == 0;
		}
		return true;
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, bool>)D);
	}

	private static bool D(DateTime A, DateTime B)
	{
		return DateTime.Compare(A, B) == 1;
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, bool>)E);
	}

	private static bool E(DateTime A, DateTime B)
	{
		if (DateTime.Compare(A, B) != 1)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return DateTime.Compare(A, B) == 0;
				}
			}
		}
		return true;
	}

	private static void A(WorksheetItem A, Range B, Func<DateTime, DateTime, bool> C)
	{
		DateTime arg;
		try
		{
			arg = DateTime.Parse(Props.SearchForm.Input1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103097));
			ProjectData.ClearProjectError();
			return;
		}
		B = RangeHelpers.H(B);
		if (B == null)
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
				Range A2 = null;
				enumerator = B.Areas.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
						{
							try
							{
								if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
									if (!C(Conversions.ToDate(objectValue), arg))
									{
										break;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										RangeHelpers.A(ref A2, range);
										break;
									}
									break;
								}
								continue;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							continue;
						}
						int num = Information.LBound((Array)objectValue);
						int num2 = Information.UBound((Array)objectValue);
						for (int i = num; i <= num2; i++)
						{
							int num3 = Information.LBound((Array)objectValue, 2);
							int num4 = Information.UBound((Array)objectValue, 2);
							for (int j = num3; j <= num4; j++)
							{
								object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null));
								try
								{
									if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue2)) || !C(Conversions.ToDate(objectValue2), arg))
									{
										continue;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
										break;
									}
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_01fd;
								}
								continue;
								end_IL_01fd:
								break;
							}
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0224;
						}
						continue;
						end_IL_0224:
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
				U(A, A2);
				A2 = null;
				return;
			}
		}
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, DateTime, bool>)Dates.A);
	}

	private static bool A(DateTime A, DateTime B, DateTime C)
	{
		if (DateTime.Compare(A, B) == 1)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return DateTime.Compare(A, C) == -1;
				}
			}
		}
		return false;
	}

	internal static void H(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, DateTime, DateTime, bool>)Dates.B);
	}

	private static bool B(DateTime A, DateTime B, DateTime C)
	{
		if (DateTime.Compare(A, B) != -1)
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
			if (DateTime.Compare(A, B) != 0)
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
				if (DateTime.Compare(A, C) != 1)
				{
					return DateTime.Compare(A, C) == 0;
				}
			}
		}
		return true;
	}

	private static void A(WorksheetItem A, Range B, Func<DateTime, DateTime, DateTime, bool> C)
	{
		DateTime arg;
		DateTime arg2;
		try
		{
			arg = DateTime.Parse(Props.SearchForm.Input1);
			arg2 = DateTime.Parse(Props.SearchForm.Input2);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103097));
			ProjectData.ClearProjectError();
			return;
		}
		B = RangeHelpers.H(B);
		if (B == null)
		{
			return;
		}
		checked
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
				Range A2 = null;
				try
				{
					enumerator = B.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
							try
							{
								if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
									if (C(Conversions.ToDate(objectValue), arg, arg2))
									{
										RangeHelpers.A(ref A2, range);
									}
									break;
								}
								continue;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							continue;
						}
						int num = Information.LBound((Array)objectValue);
						int num2 = Information.UBound((Array)objectValue);
						for (int i = num; i <= num2; i++)
						{
							int num3 = Information.LBound((Array)objectValue, 2);
							int num4 = Information.UBound((Array)objectValue, 2);
							for (int j = num3; j <= num4; j++)
							{
								object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null));
								try
								{
									if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue2)))
									{
										continue;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										if (!C(Conversions.ToDate(objectValue2), arg, arg2))
										{
											break;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
											break;
										}
										break;
									}
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
							}
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
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0243;
						}
						continue;
						end_IL_0243:
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
				U(A, A2);
				A2 = null;
				return;
			}
		}
	}

	internal static void I(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, bool>)C);
	}

	private static bool C(DateTime A)
	{
		return DateTime.Compare(A, DateTime.Now) > 0;
	}

	internal static void J(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, bool>)D);
	}

	private static bool D(DateTime A)
	{
		return DateTime.Compare(A, DateTime.Now) < 0;
	}

	internal static void K(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, bool>)E);
	}

	private static bool E(DateTime A)
	{
		bool flag = default(bool);
		return DateAndTime.Weekday(A) switch
		{
			1 => false, 
			2 => true, 
			3 => true, 
			4 => true, 
			5 => true, 
			6 => true, 
			7 => false, 
			_ => flag, 
		};
	}

	internal static void L(WorksheetItem A, Range B)
	{
		Dates.A(A, B, F);
	}

	private static bool F(DateTime A)
	{
		return !E(A);
	}

	internal static void M(WorksheetItem A, Range B)
	{
		Dates.A(A, B, G);
	}

	private static bool G(DateTime A)
	{
		return DateAndTime.Weekday(A) == 2;
	}

	internal static void N(WorksheetItem A, Range B)
	{
		Dates.A(A, B, H);
	}

	private static bool H(DateTime A)
	{
		return DateAndTime.Weekday(A) == 3;
	}

	internal static void O(WorksheetItem A, Range B)
	{
		Dates.A(A, B, I);
	}

	private static bool I(DateTime A)
	{
		return DateAndTime.Weekday(A) == 4;
	}

	internal static void P(WorksheetItem A, Range B)
	{
		Dates.A(A, B, J);
	}

	private static bool J(DateTime A)
	{
		return DateAndTime.Weekday(A) == 5;
	}

	internal static void Q(WorksheetItem A, Range B)
	{
		Dates.A(A, B, K);
	}

	private static bool K(DateTime A)
	{
		return DateAndTime.Weekday(A) == 6;
	}

	private static void A(WorksheetItem A, Range B, Func<DateTime, bool> C)
	{
		B = RangeHelpers.H(B);
		if (B == null)
		{
			return;
		}
		checked
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
				Range A2 = null;
				try
				{
					enumerator = B.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
							try
							{
								if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
									if (!C(Conversions.ToDate(objectValue)))
									{
										break;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										RangeHelpers.A(ref A2, range);
										break;
									}
									break;
								}
								continue;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							continue;
						}
						int num = Information.LBound((Array)objectValue);
						int num2 = Information.UBound((Array)objectValue);
						for (int i = num; i <= num2; i++)
						{
							int num3 = Information.LBound((Array)objectValue, 2);
							int num4 = Information.UBound((Array)objectValue, 2);
							for (int j = num3; j <= num4; j++)
							{
								object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null));
								try
								{
									if (Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue2)) && C(Conversions.ToDate(objectValue2)))
									{
										RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
								}
							}
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
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_01e7;
						}
						continue;
						end_IL_01e7:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				U(A, A2);
				A2 = null;
				return;
			}
		}
	}

	internal static void R(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, int, bool>)Dates.A);
	}

	private static bool A(DateTime A, int B)
	{
		return A.Year == B;
	}

	internal static void S(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, int, bool>)Dates.B);
	}

	private static bool B(DateTime A, int B)
	{
		return A.Month == B;
	}

	internal static void T(WorksheetItem A, Range B)
	{
		Dates.A(A, B, (Func<DateTime, int, bool>)C);
	}

	private static bool C(DateTime A, int B)
	{
		int num = default(int);
		switch (A.Month)
		{
		case 1:
		case 2:
		case 3:
			num = 1;
			break;
		case 4:
		case 5:
		case 6:
			num = 2;
			break;
		case 7:
		case 8:
		case 9:
			num = 3;
			break;
		case 10:
		case 11:
		case 12:
			num = 4;
			break;
		}
		return num == B;
	}

	private static void A(WorksheetItem A, Range B, Func<DateTime, int, bool> C)
	{
		B = RangeHelpers.H(B);
		if (B == null)
		{
			return;
		}
		Range A2 = null;
		string[] array = Props.SearchForm.Input1.Split(',');
		checked
		{
			int num = array.Length - 1;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = B.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
					int result;
					if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
					{
						try
						{
							if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								int num2 = num;
								for (int i = 0; i <= num2; i++)
								{
									if (!int.TryParse(array[i].Trim(), out result))
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
										break;
									}
									if (C(Conversions.ToDate(objectValue), result))
									{
										RangeHelpers.A(ref A2, range);
									}
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_010f;
									}
									continue;
									end_IL_010f:
									break;
								}
								break;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						continue;
					}
					int num3 = Information.LBound((Array)objectValue);
					int num4 = Information.UBound((Array)objectValue);
					for (int j = num3; j <= num4; j++)
					{
						int num5 = Information.LBound((Array)objectValue, 2);
						int num6 = Information.UBound((Array)objectValue, 2);
						for (int k = num5; k <= num6; k++)
						{
							object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { j, k }, null));
							try
							{
								if (!Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue2)))
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
									int num7 = num;
									for (int l = 0; l <= num7; l++)
									{
										if (!int.TryParse(array[l].Trim(), out result))
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
										if (!C(Conversions.ToDate(objectValue2), result))
										{
											continue;
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
										RangeHelpers.A(ref A2, (Range)range.Cells[j, k]);
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_023f;
										}
										continue;
										end_IL_023f:
										break;
									}
									break;
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0269;
							}
							continue;
							end_IL_0269:
							break;
						}
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_028e;
					}
					continue;
					end_IL_028e:
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
			U(A, A2);
			A2 = null;
		}
	}

	private static void U(WorksheetItem A, Range B)
	{
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = B.Rows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range a = (Range)enumerator.Current;
					A.C(a);
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
						switch (1)
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
	}

	private static void V(WorksheetItem A, Range B)
	{
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = B.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range a = (Range)enumerator.Current;
					A.C(a);
				}
				return;
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
	}
}
