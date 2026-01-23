using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Untranspose
{
	public static void Go()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
			Application application = MH.A.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			Range activeCell;
			Range range;
			try
			{
				activeCell = application.ActiveCell;
				range = (Range)application.Selection;
				if (Conversions.ToBoolean(activeCell.HasArray))
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						if (A(activeCell))
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								if (!Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, 1, TextCompare: false))
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
									if (!Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, 1, TextCompare: false))
									{
										Forms.WarningMessage(VH.A(156155));
										break;
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
								if (Information.IsDBNull(RuntimeHelpers.GetObjectValue(range.FormulaArray)))
								{
									Forms.WarningMessage(VH.A(155933));
								}
								else
								{
									A(range, activeCell);
								}
								break;
							}
						}
						else
						{
							Forms.WarningMessage(VH.A(156242));
						}
						break;
					}
				}
				else if (Conversions.ToBoolean(NewLateBinding.LateGet(activeCell, null, VH.A(46494), new object[0], null, null, null)))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (A((Range)NewLateBinding.LateGet(activeCell, null, VH.A(46511), new object[0], null, null, null)))
						{
							if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									Range obj = (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(activeCell, null, VH.A(46511), new object[0], null, null, null), null, VH.A(103802), new object[0], null, null, null);
									A(obj, (Range)NewLateBinding.LateGet(activeCell, null, VH.A(46511), new object[0], null, null, null));
									obj.Select();
									activeCell.Activate();
									break;
								}
							}
							else
							{
								Forms.WarningMessage(VH.A(156343));
							}
						}
						else
						{
							Forms.WarningMessage(VH.A(156473));
						}
						break;
					}
				}
				else
				{
					Forms.WarningMessage(VH.A(156560));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			activeCell = null;
			range = null;
			application = null;
			return;
		}
	}

	private static bool A(Range A)
	{
		return A.Formula.ToString().StartsWith(VH.A(156681));
	}

	private static void A(Range A, Range B)
	{
		string text = Untranspose.A(B);
		Range range = ((_Application)A.Application).get_Range((object)text, RuntimeHelpers.GetObjectValue(Missing.Value));
		bool flag = text.Contains(VH.A(7827));
		bool flag2 = JH.A(A);
		Range range2 = A;
		range2.ClearContents();
		long num = Conversions.ToLong(range2.Rows.CountLarge);
		long num2 = Conversions.ToLong(range2.Columns.CountLarge);
		checked
		{
			if (num2 > num)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				int num3 = (int)num2;
				for (int i = 1; i <= num3; i++)
				{
					object instance = range2.Cells[1, i];
					string memberName = VH.A(68956);
					object[] array = new object[1];
					string left = VH.A(48936);
					object[] array2;
					bool[] array3;
					object right = NewLateBinding.LateGet(range.Cells[i, 1], null, VH.A(5814), array2 = new object[4]
					{
						0,
						0,
						XlReferenceStyle.xlA1,
						flag
					}, null, null, array3 = new bool[4] { false, false, false, true });
					if (array3[3])
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
						flag = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[3]), typeof(bool));
					}
					array[0] = Operators.ConcatenateObject(left, right);
					NewLateBinding.LateSetComplex(instance, null, memberName, array, null, null, OptimisticSet: false, RValueBase: true);
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
			}
			else
			{
				int num4 = (int)num;
				for (int j = 1; j <= num4; j++)
				{
					object instance2 = range2.Cells[j, 1];
					string memberName2 = VH.A(68956);
					object[] array4 = new object[1];
					string left2 = VH.A(48936);
					object[] array2;
					bool[] array3;
					object right2 = NewLateBinding.LateGet(range.Cells[1, j], null, VH.A(5814), array2 = new object[4]
					{
						0,
						0,
						XlReferenceStyle.xlA1,
						flag
					}, null, null, array3 = new bool[4] { false, false, false, true });
					if (array3[3])
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
						flag = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[3]), typeof(bool));
					}
					array4[0] = Operators.ConcatenateObject(left2, right2);
					NewLateBinding.LateSetComplex(instance2, null, memberName2, array4, null, null, OptimisticSet: false, RValueBase: true);
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
			range2 = null;
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
				JH.A(A, VH.A(156702));
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(156702));
		}
	}

	private static string A(Range A)
	{
		Regex regex = new Regex(VH.A(156725));
		try
		{
			return regex.Match(Conversions.ToString(A.Formula)).Groups[1].ToString();
		}
		finally
		{
			regex = null;
		}
	}
}
