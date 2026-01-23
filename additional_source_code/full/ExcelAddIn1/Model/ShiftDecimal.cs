using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class ShiftDecimal
{
	public static void Left()
	{
		A(VH.A(94462));
	}

	public static void Right()
	{
		A(VH.A(94471));
	}

	private static void A(string A)
	{
		if (!Licensing.AllowRestrictedMode())
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
			Range range = null;
			bool flag = false;
			Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				Action<Range> action;
				if (Operators.CompareString(A, VH.A(94462), TextCompare: false) == 0)
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
					action = ShiftDecimal.A;
				}
				else
				{
					action = B;
				}
				Range range2;
				try
				{
					range2 = JH.A((Range)null);
					bool flag2 = JH.A(range2);
					if (Operators.ConditionalCompareObjectGreater(range2.Cells.CountLarge, 1, TextCompare: false))
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
							range = range2.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (range != null)
						{
							application.ScreenUpdating = false;
							application.EnableEvents = false;
							XlCalculation calculation = application.Calculation;
							application.Calculation = XlCalculation.xlCalculationManual;
							try
							{
								try
								{
									enumerator = range.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Range range3 = (Range)enumerator.Current;
										if (!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range3.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
										action(range3);
										flag = true;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0167;
										}
										continue;
										end_IL_0167:
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
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							application.Calculation = calculation;
							application.EnableEvents = true;
							application.ScreenUpdating = true;
							range = null;
						}
						goto IL_023f;
					}
					int num;
					if (Conversions.ToBoolean(Operators.NotObject(range2.HasFormula)))
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
						if (!string.IsNullOrEmpty(Conversions.ToString(range2.Text)))
						{
							num = (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))) ? 1 : 0);
							goto IL_021d;
						}
					}
					num = 0;
					goto IL_021d;
					IL_021d:
					if (Conversions.ToBoolean((byte)num != 0))
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
						action(range2);
						flag = true;
					}
					goto IL_023f;
					IL_023f:
					if (flag2)
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
						JH.A(range2, VH.A(94482));
					}
					if (!flag)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(VH.A(94509));
							break;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				range2 = null;
				action = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(94598));
			}
			application = null;
			return;
		}
	}

	private static void A(Range A)
	{
		A.Value2 = Operators.DivideObject(A.Value2, 10);
	}

	private static void B(Range A)
	{
		A.Value2 = Operators.MultiplyObject(A.Value2, 10);
	}
}
