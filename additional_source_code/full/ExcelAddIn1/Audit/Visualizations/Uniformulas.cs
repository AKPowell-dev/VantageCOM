using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Formulas;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class Uniformulas
{
	public static void Apply()
	{
		if (!Licensing.AllowVisualizationOperation())
		{
			return;
		}
		XlCalculation calculation = default(XlCalculation);
		IEnumerator enumerator = default(IEnumerator);
		Range range2 = default(Range);
		Worksheet worksheet = default(Worksheet);
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
			Microsoft.Office.Interop.Excel.Window window = null;
			bool flag = false;
			if (!(application.Selection is Range))
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
				Range activeCell = application.ActiveCell;
				if (Conversions.ToBoolean(activeCell.HasArray))
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
					activeCell.CurrentArray.Select();
					activeCell.Activate();
				}
				else
				{
					Range range = ExcelAddIn1.Formulas.Helpers.SpecialCellsFormulas(activeCell.CurrentRegion);
					if (range != null)
					{
						Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
						if (!Workbooks.IsShared(activeWorkbook, true, (System.Windows.Window)null))
						{
							if (activeWorkbook.Windows.Count > 1)
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
								window = application.ActiveWindow;
							}
							try
							{
								Microsoft.Office.Interop.Excel.Application application2 = application;
								application2.ScreenUpdating = false;
								application2.EnableEvents = false;
								application2.DisplayAlerts = false;
								calculation = application2.Calculation;
								application2.Calculation = XlCalculation.xlCalculationManual;
								NewLateBinding.LateCall(application2.ActiveSheet, null, VH.A(224), new object[1] { application2.ActiveSheet }, new string[1] { VH.A(51175) }, null, null, IgnoreReturn: true);
								worksheet = (Worksheet)application2.ActiveSheet;
								application2 = null;
								window?.Activate();
								range2 = activeCell;
								activeCell.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
								((_Worksheet)worksheet).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								try
								{
									enumerator = range.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Range range3 = (Range)enumerator.Current;
										if (!Conversions.ToBoolean(Operators.NotObject(range3.HasArray)) || !Operators.ConditionalCompareObjectEqual(((_Worksheet)worksheet).get_Range((object)range3.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula, range3.Formula, TextCompare: false))
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
										range2 = application.Union(range2, range3, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_0405;
										}
										continue;
										end_IL_0405:
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
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							try
							{
								worksheet.Delete();
								range2.Select();
								activeCell.Activate();
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							Microsoft.Office.Interop.Excel.Application application3 = application;
							application3.ScreenUpdating = true;
							application3.EnableEvents = true;
							application3.DisplayAlerts = true;
							application3.Calculation = calculation;
							_ = null;
							range2 = null;
							worksheet = null;
						}
						else
						{
							flag = true;
						}
						range = null;
					}
				}
				if (!flag)
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
					try
					{
						if (KH.A.UniformRangeEditMode)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								EditMode.EnterEditMode(application);
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
				}
				application = null;
				activeCell = null;
				window = null;
				Common.A(VH.A(51580));
				return;
			}
		}
	}
}
