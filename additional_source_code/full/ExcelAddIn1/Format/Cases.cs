using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Cases
{
	public static void Cycle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
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
			try
			{
				Range activeCell = MH.A.Application.ActiveCell;
				object left = activeCell.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value));
				VbStrConv a;
				if (Operators.ConditionalCompareObjectEqual(left, Strings.StrConv(Conversions.ToString(activeCell.Text), VbStrConv.Lowercase), TextCompare: false))
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
					a = VbStrConv.None;
				}
				else if (Operators.ConditionalCompareObjectEqual(left, Strings.StrConv(Conversions.ToString(activeCell.Text), VbStrConv.Uppercase), TextCompare: false))
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
					a = VbStrConv.Lowercase;
				}
				else if (Operators.ConditionalCompareObjectEqual(left, Strings.StrConv(Conversions.ToString(activeCell.Text), VbStrConv.ProperCase), TextCompare: false))
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
					a = VbStrConv.Uppercase;
				}
				else
				{
					a = VbStrConv.ProperCase;
				}
				activeCell = null;
				A(a);
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	public static void DoCycleCase(IRibbonControl control)
	{
		string tag = control.Tag;
		VbStrConv a = default(VbStrConv);
		if (Operators.CompareString(tag, VH.A(148701), TextCompare: false) != 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(tag, VH.A(148718), TextCompare: false) != 0)
			{
				if (Operators.CompareString(tag, VH.A(148729), TextCompare: false) != 0)
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
					if (Operators.CompareString(tag, VH.A(148740), TextCompare: false) != 0)
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
					}
					else
					{
						a = VbStrConv.Lowercase;
					}
				}
				else
				{
					a = VbStrConv.Uppercase;
				}
			}
			else
			{
				a = VbStrConv.ProperCase;
			}
		}
		else
		{
			a = VbStrConv.None;
		}
		A(a);
	}

	private static void A(VbStrConv A)
	{
		Range range = null;
		Application application = MH.A.Application;
		if (application.Selection is Range)
		{
			Range range2 = (Range)application.Selection;
			if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
			{
				if (!Conversions.ToBoolean(range2.HasFormula))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (Base.IsString(range2))
					{
						range = range2;
					}
				}
			}
			else
			{
				try
				{
					range = range2.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlTextValues);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			range2 = null;
			if (range != null)
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
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				try
				{
					bool flag = JH.A(range);
					IEnumerator enumerator = range.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Range range3 = (Range)enumerator.Current;
							if (A == VbStrConv.None)
							{
								string text = Strings.StrConv(Conversions.ToString(range3.Value2), VbStrConv.Lowercase);
								range3.Value2 = Strings.UCase(Strings.Left(text, 1)) + Strings.Right(text, checked(text.Length - 1));
							}
							else
							{
								range3.Value2 = Strings.StrConv(Conversions.ToString(range3.Value2), A);
							}
							range3 = null;
						}
						while (true)
						{
							switch (6)
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
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					if (flag)
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
						JH.A(range, VH.A(148751));
					}
					Base.LogActivity(VH.A(148774));
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				range = null;
			}
			else
			{
				Forms.WarningMessage(VH.A(148795));
			}
		}
		application = null;
	}
}
