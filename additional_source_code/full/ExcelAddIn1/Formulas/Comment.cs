using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Comment
{
	public static void Cells()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
			Application application = MH.A.Application;
			Range range = null;
			Range range2 = null;
			Range range3 = null;
			Range range4 = null;
			bool flag;
			Worksheet worksheet;
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
				worksheet = application.ActiveCell.Worksheet;
				if (application.ActiveWindow.SelectedSheets.Count == 1)
				{
					ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
				}
				if (!worksheet.ProtectContents)
				{
					flag = Operators.CompareString(application.ActiveCell.PrefixCharacter.ToString(), VH.A(39851), TextCompare: false) == 0;
					range = (Range)application.Selection;
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
							if (Conversions.ToBoolean(Operators.NotObject(range.HasFormula)))
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
								range2 = range;
							}
						}
						else
						{
							try
							{
								range2 = range.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
					}
					else if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
					{
						int num;
						if (!Conversions.ToBoolean(range.HasFormula))
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
							num = (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)) ? 1 : 0);
						}
						else
						{
							num = 1;
						}
						if (Conversions.ToBoolean((byte)num != 0))
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
							range2 = range;
						}
					}
					else
					{
						try
						{
							range3 = range.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						try
						{
							range4 = range.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
						if (range3 != null)
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
							if (range4 != null)
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
								range2 = application.Union(range3, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								goto IL_039e;
							}
						}
						if (range3 != null)
						{
							range2 = range3;
						}
						else if (range4 != null)
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
							range2 = range4;
						}
					}
					goto IL_039e;
				}
				Forms.ErrorMessage(VH.A(152392));
				goto IL_0597;
			}
			goto IL_0599;
			IL_0599:
			application = null;
			return;
			IL_0597:
			worksheet = null;
			goto IL_0599;
			IL_039e:
			if (range2 != null)
			{
				Application application2 = application;
				application2.ScreenUpdating = false;
				application2.EnableEvents = false;
				XlCalculation calculation = application2.Calculation;
				application2.Calculation = XlCalculation.xlCalculationManual;
				_ = null;
				bool flag2 = JH.A(range);
				try
				{
					enumerator = range2.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range5 = (Range)enumerator.Current;
						try
						{
							Range range6 = range5;
							if (Conversions.ToBoolean(Operators.NotObject(range6.HasArray)))
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
								if (!flag)
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
									try
									{
										range6.Value2 = VH.A(39851) + NewLateBinding.LateGet(range6, null, VH.A(1998), new object[0], null, null, null).ToString();
									}
									catch (Exception ex7)
									{
										ProjectData.SetProjectError(ex7);
										Exception ex8 = ex7;
										range6.Value2 = VH.A(39851) + range6.Formula.ToString();
										ProjectData.ClearProjectError();
									}
								}
								else
								{
									try
									{
										NewLateBinding.LateSet(range6, null, VH.A(1998), new object[1] { range6.Value2.ToString() }, null, null);
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										range6.Formula = range6.Value2.ToString();
										ProjectData.ClearProjectError();
									}
								}
							}
							range6 = null;
						}
						catch (Exception ex11)
						{
							ProjectData.SetProjectError(ex11);
							Exception ex12 = ex11;
							ProjectData.ClearProjectError();
						}
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
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
					JH.A(range, VH.A(152361));
				}
				Application application3 = application;
				application3.Calculation = calculation;
				application3.EnableEvents = true;
				application3.ScreenUpdating = true;
				_ = null;
				range2 = null;
			}
			range = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(152361));
			goto IL_0597;
		}
	}
}
