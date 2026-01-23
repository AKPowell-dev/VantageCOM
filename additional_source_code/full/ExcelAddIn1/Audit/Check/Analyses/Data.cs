using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Data
{
	internal static void A(ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = RangeHelpers.E(C);
			if (range == null)
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
				try
				{
					enumerator = range.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						if (!range2.Validation.Value)
						{
							A.DataValidationFailed.Add(range2);
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0073;
						}
						continue;
						end_IL_0073:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				range = null;
				return;
			}
		}
	}

	internal static void A(Analysis A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		List<Observation> observations = A.Observations;
		Range range = null;
		try
		{
			range = C.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range == null)
		{
			return;
		}
		A.ActionStarted(VH.A(2381), range.Areas.Count);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = range.Areas.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					if (A.ItemCancelled())
					{
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
							break;
						}
						break;
					}
					if (Conversions.ToLong(range2.Cells.CountLarge) <= 3)
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
					Range range3;
					try
					{
						range3 = range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						range3 = null;
						ProjectData.ClearProjectError();
					}
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
						try
						{
							enumerator2 = range2.Rows.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range c = (Range)enumerator2.Current;
								Data.A(observations, B, c);
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
						range3 = null;
					}
					else
					{
						Data.A(observations, B, range2);
					}
					continue;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_016c;
					}
					continue;
					end_IL_016c:
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
		A.ActionEnded();
		range = null;
	}

	private static void A(List<Observation> A, Severity B, Range C)
	{
		WorksheetFunction worksheetFunction = C.Application.WorksheetFunction;
		double num;
		double num2;
		try
		{
			num = worksheetFunction.Quartile_Inc(C, 1.0);
			num2 = worksheetFunction.Quartile_Inc(C, 3.0);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
			return;
		}
		worksheetFunction = null;
		double num3 = num2 - num;
		double num4 = num - 1.5 * num3;
		double num5 = num2 + 1.5 * num3;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = C.GetEnumerator();
			while (enumerator.MoveNext())
			{
				if (!Information.IsDate(RuntimeHelpers.GetObjectValue(((Range)enumerator.Current).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_00ca;
				}
				continue;
				end_IL_00ca:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = C.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Range range = (Range)enumerator2.Current;
				double num6;
				try
				{
					num6 = Convert.ToDouble(RuntimeHelpers.GetObjectValue(range.Value2));
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
					continue;
				}
				if (!(num6 < num4))
				{
					if (!(num6 > num5))
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
				A.Add(new DataOutlier(B, range, C));
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (1)
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

	internal static void A(Analysis A, ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Severity C, Range D)
	{
		if (C == Severity.Ignore)
		{
			return;
		}
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
			Range range = null;
			try
			{
				range = D.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlTextValues);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (range == null)
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
				A.ActionStarted(VH.A(2436), Conversions.ToLong(range.Cells.CountLarge));
				try
				{
					enumerator = range.Cells.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
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
										goto end_IL_0096;
									}
									continue;
									end_IL_0096:
									break;
								}
								break;
							}
							if (!range2.Errors.get_Item((object)XlErrorChecks.xlNumberAsText).Value)
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
							B.NumbersStoredAsText.Add(range2);
							continue;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00e1;
							}
							continue;
							end_IL_00e1:
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
				return;
			}
		}
	}
}
