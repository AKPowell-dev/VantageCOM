using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using ExcelAddIn1.Format;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Formatting
{
	internal static void A(Analysis A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		foreach (Range item in RangeHelpers.A(C, A, VH.A(2512)))
		{
			A.Observations.Add(new MergedCells(B, item));
		}
	}

	internal static void A(ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(C.NumberFormat.ToString(), VH.A(2545), TextCompare: false) == 0)
			{
				A.TripleSemicolons.Add(C);
			}
			return;
		}
	}

	internal static void B(Analysis A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		List<Observation> observations = A.Observations;
		string text = KH.A.AutoColors[0];
		if (text.Length <= 0)
		{
			return;
		}
		Range range = RangeHelpers.B(C);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			int num = clsColors.RGB2Ole(text);
			A.ActionStarted(VH.A(2552), range.Areas.Count);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					if (A.ItemCancelled())
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00b0;
							}
							continue;
							end_IL_00b0:
							break;
						}
						break;
					}
					if (Operators.ConditionalCompareObjectEqual(range2.CountLarge, 1, TextCompare: false))
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
						Formatting.A(observations, B, range2, num);
					}
					else if (Information.IsDBNull(RuntimeHelpers.GetObjectValue(range2.Font.Color)))
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
						A.ActionStarted(VH.A(2611), Conversions.ToLong(range2.CountLarge));
						try
						{
							enumerator2 = range2.GetEnumerator();
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									Range c = (Range)enumerator2.Current;
									if (A.ItemCancelled())
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_0162;
											}
											continue;
											end_IL_0162:
											break;
										}
										break;
									}
									Formatting.A(observations, B, c, num);
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0181;
									}
									continue;
									end_IL_0181:
									break;
								}
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						A.ActionEnded();
					}
					else if (Operators.ConditionalCompareObjectNotEqual(range2.Font.Color, num, TextCompare: false))
					{
						observations.Add(new UncoloredInput(B, range2));
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			A.ActionEnded();
			range = null;
			return;
		}
	}

	private static void A(List<Observation> A, Severity B, Range C, int D)
	{
		if (Operators.ConditionalCompareObjectNotEqual(C.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					throw new NotImplementedException();
				}
			}
		}
		if (Information.IsDate(RuntimeHelpers.GetObjectValue(C.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
			if (!KH.A.AutoColorDates)
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
				break;
			}
		}
		if (Operators.ConditionalCompareObjectNotEqual(C.Font.Color, D, TextCompare: false))
		{
			A.Add(new UncoloredInput(B, C));
		}
	}

	internal static void A(ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C, List<int> D)
	{
		if (B == Severity.Ignore || !FillColor.HasFill(C.Interior))
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
			if (!D.Contains(Conversions.ToInteger(C.Interior.Color)))
			{
				A.CellFillColors.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor(C, Conversions.ToInteger(C.Interior.Color)));
			}
			return;
		}
	}

	internal static void B(ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C, List<int> D)
	{
		if (B == Severity.Ignore)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			XlBordersIndex[] array = new XlBordersIndex[4]
			{
				XlBordersIndex.xlEdgeTop,
				XlBordersIndex.xlEdgeBottom,
				XlBordersIndex.xlEdgeLeft,
				XlBordersIndex.xlEdgeRight
			};
			int num = 0;
			Border border;
			while (true)
			{
				if (num < array.Length)
				{
					XlBordersIndex index = array[num];
					border = C.Borders[index];
					if (ExcelAddIn1.Format.Borders.HasBorder(border))
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
						if (!D.Contains(Conversions.ToInteger(border.Color)))
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
							A.CellBorderColors.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor(C, Conversions.ToInteger(border.Color)));
							break;
						}
					}
					num = checked(num + 1);
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
				break;
			}
			border = null;
			return;
		}
	}

	internal static void A(Analysis A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
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
			using List<Range>.Enumerator enumerator = RangeHelpers.A(C, A, VH.A(2512)).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				A.Observations.Add(new ConditionalFormat(B, current));
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}
}
