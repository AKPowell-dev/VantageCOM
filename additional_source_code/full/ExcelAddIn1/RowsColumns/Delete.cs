using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Delete
{
	public static void Row()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
		{
			application = null;
			return;
		}
		Microsoft.Office.Interop.Excel.Application application2 = application;
		int num = default(int);
		Range range = default(Range);
		try
		{
			if (application2.Windows.Count > 0)
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
					if (!(application2.Selection is Range))
					{
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						num = application2.ActiveWindow.SelectedSheets.Count;
						if (num > 1)
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
							if (!Core.ConfirmMultipleSheets())
							{
								num = 0;
								break;
							}
						}
						range = (Range)application2.Selection;
						if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, range.Worksheet.Rows.CountLarge, TextCompare: false))
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
							if (MessageBox.Show(VH.A(170358), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
							{
								num = 0;
								break;
							}
						}
						string cell = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
						application2.ScreenUpdating = false;
						application2.EnableEvents = false;
						if (num > 1)
						{
							application2.Calculation = XlCalculation.xlCalculationManual;
						}
						try
						{
							enumerator = application2.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (objectValue is Worksheet)
								{
									((_Worksheet)(Worksheet)objectValue).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireRow.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
								}
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_01c3;
								}
								continue;
								end_IL_01c3:
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
						Core.LogActivity(VH.A(170504));
						break;
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (num > 1)
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
				XlCalculation calculation = default(XlCalculation);
				application2.Calculation = calculation;
			}
			application2.ScreenUpdating = true;
			application2.EnableEvents = true;
		}
		application2 = null;
		JH.A((object)range);
		application = null;
	}

	public static void Column()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
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
					application = null;
					return;
				}
			}
		}
		Microsoft.Office.Interop.Excel.Application application2 = application;
		int num = default(int);
		Range range = default(Range);
		try
		{
			if (application2.Windows.Count > 0)
			{
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (!(application2.Selection is Range))
					{
						break;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						num = application2.ActiveWindow.SelectedSheets.Count;
						if (num > 1)
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
							if (!Core.ConfirmMultipleSheets())
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									num = 0;
									break;
								}
								break;
							}
						}
						range = (Range)application2.Selection;
						if (Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, range.Worksheet.Columns.CountLarge, TextCompare: false))
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
							if (MessageBox.Show(VH.A(170525), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
							{
								num = 0;
								break;
							}
						}
						string cell = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
						application2.ScreenUpdating = false;
						application2.EnableEvents = false;
						if (num > 1)
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
							application2.Calculation = XlCalculation.xlCalculationManual;
						}
						try
						{
							enumerator = application2.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue is Worksheet))
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
								((_Worksheet)(Worksheet)objectValue).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_01f1;
								}
								continue;
								end_IL_01f1:
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
						Core.LogActivity(VH.A(170677));
						break;
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			if (num > 1)
			{
				XlCalculation calculation = default(XlCalculation);
				application2.Calculation = calculation;
			}
			application2.ScreenUpdating = true;
			application2.EnableEvents = true;
		}
		application2 = null;
		JH.A((object)range);
		application = null;
	}

	public static void BlankRows()
	{
		if (!A())
		{
			return;
		}
		checked
		{
			Range range = default(Range);
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
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				long num = 0L;
				bool flag = false;
				try
				{
					if (application.Selection is Range)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							range = (Range)application.Selection;
							if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
								range = range.Worksheet.Cells;
								range.Select();
							}
							if (MessageBox.Show(VH.A(170704), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
							{
								break;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								XlCalculation calculation = application.Calculation;
								application.Calculation = XlCalculation.xlCalculationManual;
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								Range range2 = null;
								try
								{
									range2 = ((Range)range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).SpecialCells(XlCellType.xlCellTypeBlanks, RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
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
									try
									{
										foreach (Range area in range2.Areas)
										{
											for (long num2 = Conversions.ToLong(area.Rows.CountLarge); num2 >= 1; num2 += -1)
											{
												object instance = area.Rows[num2, RuntimeHelpers.GetObjectValue(Missing.Value)];
												if (application.WorksheetFunction.CountA(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, VH.A(51236), new object[0], null, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == 0.0)
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
													NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, VH.A(51236), new object[0], null, null, null), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
													num++;
												}
												instance = null;
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
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										Forms.ErrorMessage(ex4.Message);
										flag = true;
										ProjectData.ClearProjectError();
									}
									range2 = null;
								}
								application.Calculation = calculation;
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								range = range.Worksheet.UsedRange;
								if (num > 0)
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
									Forms.InfoMessage(VH.A(52374) + num + VH.A(170854));
								}
								else if (!flag)
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
									Forms.InfoMessage(VH.A(170883));
								}
								Core.LogActivity(VH.A(170958));
								break;
							}
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
				application = null;
				JH.A((object)range);
				return;
			}
		}
	}

	public static void BlankColumns()
	{
		if (!A())
		{
			return;
		}
		checked
		{
			Range range = default(Range);
			IEnumerator enumerator = default(IEnumerator);
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				long num = 0L;
				bool flag = false;
				try
				{
					if (application.Selection is Range)
					{
						range = (Range)application.Selection;
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
							range = range.Worksheet.Cells;
							range.Select();
						}
						if (MessageBox.Show(VH.A(170993), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								XlCalculation calculation = application.Calculation;
								application.Calculation = XlCalculation.xlCalculationManual;
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								Range range2 = null;
								try
								{
									range2 = ((Range)range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).SpecialCells(XlCellType.xlCellTypeBlanks, RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								if (range2 != null)
								{
									try
									{
										try
										{
											enumerator = range2.Areas.GetEnumerator();
											while (enumerator.MoveNext())
											{
												Range range3 = (Range)enumerator.Current;
												for (long num2 = Conversions.ToLong(range3.Columns.CountLarge); num2 >= 1; num2 += -1)
												{
													object instance = range3.Columns[num2, RuntimeHelpers.GetObjectValue(Missing.Value)];
													if (application.WorksheetFunction.CountA(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, VH.A(51255), new object[0], null, null, null)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == 0.0)
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
														NewLateBinding.LateCall(NewLateBinding.LateGet(instance, null, VH.A(51255), new object[0], null, null, null), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
														num++;
													}
													instance = null;
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
												switch (6)
												{
												case 0:
													break;
												default:
													goto end_IL_03a0;
												}
												continue;
												end_IL_03a0:
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
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										Forms.ErrorMessage(ex4.Message);
										flag = true;
										ProjectData.ClearProjectError();
									}
									range2 = null;
								}
								application.Calculation = calculation;
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								range = range.Worksheet.UsedRange;
								if (num > 0)
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
									Forms.InfoMessage(VH.A(52374) + num + VH.A(171149));
								}
								else if (!flag)
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
									Forms.InfoMessage(VH.A(171184));
								}
								Core.LogActivity(VH.A(171265));
								break;
							}
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				application = null;
				JH.A((object)range);
				return;
			}
		}
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
