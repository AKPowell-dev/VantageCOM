using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class FlipSign
{
	public static void Go()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		bool flag3 = default(bool);
		IEnumerator enumerator = default(IEnumerator);
		bool flag4 = default(bool);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			bool flag = true;
			bool flag2 = false;
			if (EditMode.IsEditMode(application))
			{
				application = null;
				return;
			}
			Range range;
			try
			{
				range = (Range)application.Selection;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				application = null;
				ProjectData.ClearProjectError();
				return;
			}
			Range range2 = null;
			Range range3 = null;
			Range range4 = null;
			if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
				if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
				{
					if (Conversions.ToBoolean(range.HasFormula))
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
						range2 = range;
					}
					else
					{
						range3 = range;
					}
				}
			}
			else
			{
				try
				{
					range2 = range.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlNumbers);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				try
				{
					range3 = range.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			if (range2 != null)
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
				DialogResult dialogResult = MessageBox.Show(VH.A(152439), VH.A(40448), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				if (dialogResult == DialogResult.Cancel)
				{
					goto IL_06d9;
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
				if (dialogResult == DialogResult.No)
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
					flag = false;
				}
			}
			if (range3 != null && range2 != null && flag)
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
				application.Union(range3, range2, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			else
			{
				if (range2 != null)
				{
					if (flag)
					{
						goto IL_0304;
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
				if (range3 == null)
				{
					goto IL_06d9;
				}
			}
			goto IL_0304;
			IL_06d9:
			try
			{
				range.Select();
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			application.CutCopyMode = (XlCutCopyMode)0;
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			application = null;
			range3 = null;
			range2 = null;
			range4 = null;
			range = null;
			return;
			IL_0304:
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			try
			{
				flag3 = JH.A(range);
				range4 = Ranges.FirstBlankCell(range.Worksheet);
				range4.Value2 = -1;
				range4.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
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
					range3.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationMultiply, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				if (range2 != null && flag)
				{
					Regex regex = new Regex(VH.A(152601));
					Regex regex2 = new Regex(VH.A(152624));
					Regex regex3 = new Regex(VH.A(152637));
					XlCalculation calculation = application.Calculation;
					application.Calculation = XlCalculation.xlCalculationManual;
					try
					{
						enumerator = range2.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range range5 = (Range)enumerator.Current;
							try
							{
								if (Conversions.ToBoolean(range5.HasArray))
								{
									flag4 = true;
								}
								else
								{
									try
									{
										if (regex.IsMatch(Conversions.ToString(NewLateBinding.LateGet(range5, null, VH.A(1998), new object[0], null, null, null))))
										{
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												string input = VH.A(48936) + regex2.Replace(Conversions.ToString(NewLateBinding.LateGet(range5, null, VH.A(1998), new object[0], null, null, null)), "");
												NewLateBinding.LateSet(range5, null, VH.A(1998), new object[1] { regex3.Replace(input, "") }, null, null);
												break;
											}
										}
										else
										{
											NewLateBinding.LateSet(range5, null, VH.A(1998), new object[1] { VH.A(152644) + Strings.Mid(Conversions.ToString(NewLateBinding.LateGet(range5, null, VH.A(1998), new object[0], null, null, null)), 2) + VH.A(39904) }, null, null);
										}
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										if (regex.IsMatch(Conversions.ToString(range5.Formula)))
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
											string input = VH.A(48936) + regex2.Replace(Conversions.ToString(range5.Formula), "");
											range5.Formula = regex3.Replace(input, "");
										}
										else
										{
											range5.Formula = VH.A(152644) + Strings.Mid(Conversions.ToString(range5.Formula), 2) + VH.A(39904);
										}
										ProjectData.ClearProjectError();
									}
								}
							}
							catch (Exception ex11)
							{
								ProjectData.SetProjectError(ex11);
								Exception ex12 = ex11;
								flag2 = true;
								ProjectData.ClearProjectError();
							}
							range5 = null;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0601;
							}
							continue;
							end_IL_0601:
							break;
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
					application.Calculation = calculation;
					regex = null;
					regex2 = null;
					regex3 = null;
					if (flag2)
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
						Forms.ErrorMessage(VH.A(152651));
					}
					else if (flag4)
					{
						Forms.WarningMessage(VH.A(152748));
					}
				}
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(152900));
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				clsReporting.LogException(ex14);
				Forms.ErrorMessage(ex14.Message);
				ProjectData.ClearProjectError();
			}
			finally
			{
				try
				{
					range4.ClearContents();
				}
				catch (Exception ex15)
				{
					ProjectData.SetProjectError(ex15);
					Exception ex16 = ex15;
					ProjectData.ClearProjectError();
				}
				if (flag3)
				{
					JH.A(range, VH.A(152900));
				}
			}
			goto IL_06d9;
		}
	}
}
