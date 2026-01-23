using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.TraceDialogs;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Helpers
{
	public static Range SpecialCellsFormulas(Range rngSelection = null)
	{
		Range range = null;
		if (rngSelection == null)
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
			Application application = MH.A.Application;
			if (application.Selection is Range)
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
				rngSelection = (Range)application.Selection;
			}
			application = null;
		}
		Range result = default(Range);
		if (rngSelection != null)
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
			if (Operators.ConditionalCompareObjectEqual(rngSelection.Cells.CountLarge, 1, TextCompare: false))
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
				if (Conversions.ToBoolean(Operators.NotObject(rngSelection.HasFormula)))
				{
					goto IL_00cf;
				}
				range = rngSelection;
			}
			else
			{
				try
				{
					range = rngSelection.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					result = null;
					ProjectData.ClearProjectError();
					goto IL_00cf;
				}
			}
		}
		result = range;
		goto IL_00cf;
		IL_00cf:
		return result;
	}

	public static Range SpecialCellsValues(Range rng)
	{
		int num;
		if (Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
		{
			if (!string.IsNullOrEmpty(Conversions.ToString(rng.Text)))
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
				if (!Conversions.ToBoolean(rng.HasFormula))
				{
					num = (Conversions.ToBoolean(rng.HasArray) ? 1 : 0);
					goto IL_0069;
				}
			}
			num = 1;
			goto IL_0069;
		}
		Range range;
		Range result;
		try
		{
			range = rng.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
			goto IL_00b4;
		}
		goto IL_00b2;
		IL_00b4:
		return result;
		IL_00b2:
		result = range;
		goto IL_00b4;
		IL_0069:
		if (Conversions.ToBoolean((byte)num != 0))
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
			result = null;
			goto IL_00b4;
		}
		range = rng;
		goto IL_00b2;
	}

	public static Range CellsWithNumbers(Range rng)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		Range range2 = default(Range);
		Range result = default(Range);
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
				case 724:
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
							goto IL_000c;
						case 4:
							goto IL_0011;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_0052;
						case 7:
							goto IL_0076;
						case 9:
							goto IL_007d;
						case 11:
							goto IL_0085;
						case 12:
							goto IL_0099;
						case 8:
						case 10:
						case 13:
							goto IL_00b1;
						case 14:
							goto IL_00ce;
						case 16:
							goto IL_022f;
						case 17:
							goto IL_0236;
						case 19:
							goto IL_023e;
						case 20:
							goto IL_024f;
						case 15:
						case 18:
						case 21:
							goto IL_0257;
						case 22:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 23:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_022f:
					num2 = 16;
					if (range != null)
					{
						goto IL_0236;
					}
					goto IL_023e;
					IL_0007:
					num2 = 2;
					range2 = null;
					goto IL_000c;
					IL_000c:
					num2 = 3;
					range = null;
					goto IL_0011;
					IL_0011:
					num2 = 4;
					if (Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
					{
						goto IL_0030;
					}
					goto IL_0085;
					IL_0030:
					num2 = 5;
					if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(rng.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
					{
						goto IL_0052;
					}
					goto IL_00b1;
					IL_0052:
					num2 = 6;
					if (Conversions.ToBoolean(rng.HasFormula))
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
						goto IL_0076;
					}
					goto IL_007d;
					IL_0236:
					num2 = 17;
					result = range;
					goto end_IL_0000_3;
					IL_023e:
					num2 = 19;
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
						goto IL_024f;
					}
					goto IL_0257;
					IL_0257:
					num2 = 21;
					range2 = null;
					break;
					IL_0076:
					num2 = 7;
					range = rng;
					goto IL_00b1;
					IL_007d:
					num2 = 9;
					range2 = rng;
					goto IL_00b1;
					IL_0085:
					num2 = 11;
					range2 = rng.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
					goto IL_0099;
					IL_0099:
					num2 = 12;
					range = rng.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlNumbers);
					goto IL_00b1;
					IL_00b1:
					num2 = 13;
					if (range2 != null && range != null)
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
						goto IL_00ce;
					}
					goto IL_022f;
					IL_024f:
					num2 = 20;
					result = range2;
					goto end_IL_0000_3;
					IL_00ce:
					num2 = 14;
					result = rng.Application.Union(range2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto end_IL_0000_3;
					end_IL_0000_2:
					break;
				}
				num2 = 22;
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 724;
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

	public static string RemoveExtraneousSheetName(string strFormula, string strSheet)
	{
		if (strSheet.Contains(VH.A(39851)))
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
			strSheet = strSheet.Replace(VH.A(39851), VH.A(39854));
		}
		if (strFormula.Contains(strSheet))
		{
			if (Regex.IsMatch(strSheet, VH.A(152919)))
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
				strSheet = VH.A(39851) + strSheet + VH.A(39851);
			}
			strSheet = strSheet.Replace(VH.A(39848), VH.A(51444)).Replace(VH.A(39904), VH.A(152984)).Replace(VH.A(41262), VH.A(48931))
				.Replace(VH.A(41312), VH.A(152989))
				.Replace(VH.A(54459), VH.A(152994))
				.Replace(VH.A(43203), VH.A(152999));
			strFormula = Regex.Replace(strFormula, VH.A(153004) + strSheet + VH.A(7827), VH.A(152356));
		}
		return strFormula;
	}

	public static Range GetLabelCell(Range rng)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		Range result = default(Range);
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
				case 171:
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
						case 4:
							num2 = 4;
							range = null;
							goto end_IL_0000_3;
						default:
							goto end_IL_0000;
						case 5:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0007:
					num2 = 2;
					range = ((_Application)rng.Application).get_Range(RuntimeHelpers.GetObjectValue(rng.Worksheet.Cells[rng.Row, 1]), (object)rng).SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlTextValues);
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 3;
				result = (Range)range.get_Item((object)range.Count, RuntimeHelpers.GetObjectValue(Missing.Value));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 171;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool ContainsMergedCells(Range rng)
	{
		if (!Information.IsDBNull(RuntimeHelpers.GetObjectValue(rng.MergeCells)))
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
					return Conversions.ToBoolean(rng.MergeCells);
				}
			}
		}
		return true;
	}

	public static bool ContainsPartialInput(Range cell)
	{
		string input = Regex.Replace(A(cell), Base.CELL_REF_PATTERN, "");
		IEnumerator enumerator = Regex.Matches(input, VH.A(153039)).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				int num = Conversions.ToInteger(((Match)enumerator.Current).Groups[1].Value);
				if (num <= 10000)
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
					if (num <= 10)
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
						if ((uint)num <= 1u)
						{
							continue;
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
						if (num == 10)
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
					}
					else
					{
						if (num == 100)
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
						if (num == 1000 || num == 10000)
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
					}
				}
				else if (num <= 1000000)
				{
					if (num == 100000 || num == 1000000)
					{
						continue;
					}
				}
				else
				{
					if (num == 10000000)
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
					if (num == 100000000)
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
						break;
					}
					if (num == 1000000000)
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
				}
				return true;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0149;
				}
				continue;
				end_IL_0149:
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
		return false;
	}

	private static string A(Range A)
	{
		string text = A.Formula.ToString();
		if (text.Contains(VH.A(7827)))
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
			text = text.Replace(VH.A(39854), "");
			text = Regex.Replace(text, VH.A(153060), "");
		}
		return text;
	}

	public static string RemoveUnnecessaryEquals(Range rng, string strFormula)
	{
		string text = Regex.Replace(strFormula, VH.A(153079), VH.A(152356));
		if (KH.A.AutoColorOnEntry)
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
			if (strFormula.Length != text.Length)
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
				rng.Font.Color = clsColors.RGB2Ole(KH.A.AutoColors[0]);
			}
		}
		return text;
	}

	public static bool IsFunctionMatch(Range cell, string strFunction)
	{
		Type typeFromHandle = typeof(Helpers);
		string memberName = VH.A(153096);
		object[] array = new object[2];
		array[0] = NewLateBinding.LateGet(cell, null, VH.A(1998), new object[0], null, null, null);
		array[1] = strFunction;
		object[] array2 = array;
		bool[] obj = new bool[2] { true, true };
		bool[] array3 = obj;
		object value = NewLateBinding.LateGet(null, typeFromHandle, memberName, array, null, null, obj);
		if (array3[0])
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
			NewLateBinding.LateSetComplex(cell, null, VH.A(1998), new object[1] { array2[0] }, null, null, OptimisticSet: true, RValueBase: false);
		}
		if (array3[1])
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
			strFunction = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(string));
		}
		return Conversions.ToBoolean(value);
	}

	public static bool IsFunctionMatch(string strFormula, string strFunction)
	{
		return A(strFunction).IsMatch(strFormula);
	}

	public static MatchCollection FunctionMatches(Range cell, string strFunction)
	{
		return A(strFunction).Matches(Conversions.ToString(NewLateBinding.LateGet(cell, null, VH.A(1998), new object[0], null, null, null)));
	}

	private static Regex A(string A)
	{
		return new Regex(VH.A(4544) + A + VH.A(51444), RegexOptions.IgnoreCase);
	}

	internal static string B(Range A)
	{
		string result;
		try
		{
			result = Conversions.ToString(Operators.ConcatenateObject(NewLateBinding.LateGet(A, null, VH.A(1998), new object[0], null, null, null), ""));
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = Conversions.ToString(Operators.ConcatenateObject(A.Formula, ""));
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static void A(Range A, string B)
	{
		try
		{
			string memberName = VH.A(1998);
			object[] array = new object[1];
			object obj = B;
			if (obj == null)
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
				obj = "";
			}
			array[0] = obj;
			NewLateBinding.LateSet(A, null, memberName, array, null, null);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			object obj2 = B;
			if (obj2 == null)
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
				obj2 = "";
			}
			A.Formula = obj2;
			ProjectData.ClearProjectError();
		}
	}

	internal static bool A(Range A)
	{
		try
		{
			Range range = SpecialCellsFormulas(A);
			int num;
			if (range == null)
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
				num = 0;
			}
			else
			{
				num = range.Count;
			}
			return num > 0;
		}
		finally
		{
			Range range = null;
		}
	}
}
