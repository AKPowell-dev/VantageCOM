using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Visualizations;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Anchor
{
	public static void Cycle()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
			Application application = MH.A.Application;
			try
			{
				Application application2 = application;
				application2.ScreenUpdating = false;
				application2.EnableEvents = false;
				Range range = Helpers.SpecialCellsFormulas();
				if (range == null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						throw new Exception();
					}
				}
				Range activeCell = application2.ActiveCell;
				_ = null;
				Range range2 = activeCell;
				int num;
				try
				{
					string text = Conversions.ToString(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null));
					NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { application.ConvertFormula(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null)), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlRelative, RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null);
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null), text, TextCompare: false))
					{
						num = 4;
					}
					else
					{
						NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { application.ConvertFormula(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null)), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlAbsolute, RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null);
						if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null), text, TextCompare: false))
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
							num = 1;
						}
						else
						{
							NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { application.ConvertFormula(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null)), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlAbsRowRelColumn, RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null);
							if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null), text, TextCompare: false))
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
								num = 2;
							}
							else
							{
								num = 3;
							}
						}
					}
					NewLateBinding.LateSet(range2, null, VH.A(1998), new object[1] { text }, null, null);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					string text = Conversions.ToString(range2.Formula);
					range2.Formula = RuntimeHelpers.GetObjectValue(application.ConvertFormula(RuntimeHelpers.GetObjectValue(range2.Formula), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlRelative, RuntimeHelpers.GetObjectValue(Missing.Value)));
					if (Operators.ConditionalCompareObjectEqual(range2.Formula, text, TextCompare: false))
					{
						num = 4;
					}
					else
					{
						range2.Formula = RuntimeHelpers.GetObjectValue(application.ConvertFormula(RuntimeHelpers.GetObjectValue(range2.Formula), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlAbsolute, RuntimeHelpers.GetObjectValue(Missing.Value)));
						if (Operators.ConditionalCompareObjectEqual(range2.Formula, text, TextCompare: false))
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
							num = 1;
						}
						else
						{
							range2.Formula = RuntimeHelpers.GetObjectValue(application.ConvertFormula(RuntimeHelpers.GetObjectValue(range2.Formula), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, XlReferenceType.xlAbsRowRelColumn, RuntimeHelpers.GetObjectValue(Missing.Value)));
							num = ((!Operators.ConditionalCompareObjectEqual(range2.Formula, text, TextCompare: false)) ? 3 : 2);
						}
					}
					range2.Formula = text;
					ProjectData.ClearProjectError();
				}
				range2 = null;
				A(range, num switch
				{
					1 => 2, 
					2 => 3, 
					3 => 4, 
					_ => 1, 
				});
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			finally
			{
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				application = null;
				Range range = null;
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(151930));
			return;
		}
	}

	public static void Convert(int i)
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		Range range;
		try
		{
			range = Helpers.SpecialCellsFormulas();
			if (range != null)
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
					A(range, i);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.EnableEvents = true;
		application.ScreenUpdating = true;
		application = null;
		range = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(151971));
	}

	private static void A(Range A, int B)
	{
		Application application = A.Application;
		bool flag = JH.A(A);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (Conversions.ToBoolean(range.HasArray))
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
					if (Strings.Len(RuntimeHelpers.GetObjectValue(range.FormulaArray)) < 255)
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
						range.FormulaArray = RuntimeHelpers.GetObjectValue(application.ConvertFormula(RuntimeHelpers.GetObjectValue(range.FormulaArray), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, B, RuntimeHelpers.GetObjectValue(Missing.Value)));
					}
				}
				else
				{
					try
					{
						if (Strings.Len(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null))) < 255)
						{
							NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { application.ConvertFormula(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null)), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, B, RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						if (Strings.Len(RuntimeHelpers.GetObjectValue(range.Formula)) < 255)
						{
							range.Formula = RuntimeHelpers.GetObjectValue(application.ConvertFormula(RuntimeHelpers.GetObjectValue(range.Formula), XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, B, RuntimeHelpers.GetObjectValue(Missing.Value)));
						}
						ProjectData.ClearProjectError();
					}
				}
				range = null;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_01bc;
				}
				continue;
				end_IL_01bc:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
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
			JH.A(A, VH.A(152014));
		}
		Common.RefreshLiveVisualizations(A);
	}
}
