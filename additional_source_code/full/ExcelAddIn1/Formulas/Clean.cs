using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Clean
{
	public static void Selection()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		Application application = MH.A.Application;
		if (EditMode.IsEditMode(MH.A.Application))
		{
			application = null;
			return;
		}
		Application application2 = application;
		if (application2.Selection is Range)
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
			XlCalculation calculation = default(XlCalculation);
			Range range;
			try
			{
				application2.ScreenUpdating = false;
				application2.EnableEvents = false;
				calculation = application2.Calculation;
				application2.Calculation = XlCalculation.xlCalculationManual;
				range = (Range)application2.Selection;
				bool num = JH.A(range);
				Cells(range);
				if (num)
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
					JH.A(range, VH.A(152262));
				}
				Common.RefreshLiveVisualizations(range);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				application2.Calculation = calculation;
				application2.EnableEvents = true;
				application2.ScreenUpdating = true;
			}
			range = null;
		}
		application2 = null;
		application = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(152285));
	}

	public static void Cells(Range rng)
	{
		string text = VH.A(152312);
		Range range = Helpers.SpecialCellsFormulas(rng);
		Range range2 = Helpers.SpecialCellsValues(rng);
		rng.Application.DisplayAlerts = false;
		try
		{
			if (range != null)
			{
				string name = rng.Worksheet.Name;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = range.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range3 = (Range)enumerator.Current;
						Range range4 = range3;
						try
						{
							if (!Conversions.ToBoolean(range4.HasArray))
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
									string input;
									try
									{
										input = Regex.Replace(NewLateBinding.LateGet(range4, null, VH.A(1998), new object[0], null, null, null).ToString(), VH.A(103782), VH.A(48936));
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										input = Regex.Replace(range4.Formula.ToString(), VH.A(103782), VH.A(48936));
										ProjectData.ClearProjectError();
									}
									input = Regex.Replace(input, VH.A(103791), VH.A(48936));
									input = Helpers.RemoveExtraneousSheetName(input, name);
									if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range4.Value2)) && !input.Contains(VH.A(7827)))
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
										input = Regex.Replace(input, text + VH.A(152349), VH.A(152356));
										input = Regex.Replace(input, VH.A(152349) + text, VH.A(152356));
									}
									input = Helpers.RemoveUnnecessaryEquals(range3, input);
									try
									{
										NewLateBinding.LateSet(range4, null, VH.A(1998), new object[1] { input }, null, null);
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										range4.Formula = input;
										ProjectData.ClearProjectError();
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
						range4 = null;
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
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		rng.Application.DisplayAlerts = true;
		if (range2 != null)
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
			{
				IEnumerator enumerator2 = range2.GetEnumerator();
				try
				{
					while (enumerator2.MoveNext())
					{
						Range range5 = (Range)enumerator2.Current;
						Range range6 = range5;
						try
						{
							if (!Conversions.ToBoolean(range6.HasArray))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									if (!SumBar.B(range5))
									{
										string input = Regex.Replace(range6.Formula.ToString(), VH.A(152349), VH.A(41385));
										input = range5.Application.WorksheetFunction.Trim(input);
										range6.Formula = range5.Application.WorksheetFunction.Clean(input);
									}
									break;
								}
							}
						}
						catch (Exception ex9)
						{
							ProjectData.SetProjectError(ex9);
							Exception ex10 = ex9;
							ProjectData.ClearProjectError();
						}
						range6 = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0346;
						}
						continue;
						end_IL_0346:
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator2 as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
		}
		range = null;
		range2 = null;
	}
}
