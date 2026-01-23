using System;
using System.Collections;
using System.Collections.Generic;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class BestPractices
{
	internal static void A(ref List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		if (B == Severity.Ignore)
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
			Range usedRange = ((Microsoft.Office.Interop.Excel.Worksheet)C.Worksheets[1]).UsedRange;
			long num = 0L;
			long num2 = 0L;
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			bool flag4 = false;
			Range visibleRange = C.Windows[1].VisibleRange;
			if (usedRange.Columns.Count < visibleRange.Columns.Count)
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
				if (usedRange.Rows.Count < checked(visibleRange.Rows.Count * 2))
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
					flag2 = true;
				}
			}
			visibleRange = null;
			if (flag2)
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
				Range range = RangeHelpers.A(usedRange);
				if (range != null)
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
					num = range.Count;
					flag = num < 5;
					range = null;
				}
				else
				{
					flag = true;
				}
			}
			int num4;
			Range range2;
			if (flag)
			{
				try
				{
					range2 = usedRange.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlTextValues);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					range2 = null;
					ProjectData.ClearProjectError();
				}
				try
				{
					num2 = Conversions.ToLong(usedRange.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers).CountLarge);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
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
					long num3 = Conversions.ToLong(range2.Cells.CountLarge);
					if (num3 >= 2)
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
						if (num3 > num)
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
							num4 = ((num3 > num2) ? 1 : 0);
							goto IL_01a9;
						}
					}
					num4 = 0;
					goto IL_01a9;
				}
			}
			goto IL_0275;
			IL_0275:
			if (flag2)
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
					if (flag3)
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
						if (flag4)
						{
							goto IL_02b0;
						}
					}
				}
			}
			A.Add(new CoverMissing(B));
			goto IL_02b0;
			IL_02b0:
			usedRange = null;
			return;
			IL_01a9:
			flag3 = (byte)num4 != 0;
			if (flag3)
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
				if (range2 != null)
				{
					double num5 = Conversions.ToDouble(C.Styles[VH.A(2084)].Font.Size);
					try
					{
						enumerator = range2.GetEnumerator();
						while (true)
						{
							if (enumerator.MoveNext())
							{
								if (!(Conversions.ToDouble(((Range)enumerator.Current).Font.Size) >= num5 + 2.0))
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
									flag4 = true;
									break;
								}
								break;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0247;
								}
								continue;
								end_IL_0247:
								break;
							}
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
					range2 = null;
				}
			}
			goto IL_0275;
		}
	}
}
