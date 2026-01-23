using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class PrependEquals
{
	public static void Go()
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		Range range;
		try
		{
			range = (Range)application.Selection;
			if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
			{
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
					throw new Exception();
				}
			}
			if (Conversions.ToBoolean(range.HasFormula))
			{
				throw new Exception();
			}
			bool flag = JH.A(range);
			try
			{
				if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { Operators.ConcatenateObject(VH.A(48936), range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))) }, null, null);
						break;
					}
				}
				else
				{
					NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(73111), range.Text), VH.A(155660)) }, null, null);
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
					range.Formula = Operators.ConcatenateObject(VH.A(48936), range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
				}
				else
				{
					range.Formula = Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(73111), range.Text), VH.A(155660));
				}
				ProjectData.ClearProjectError();
			}
			if (flag)
			{
				JH.A(range, VH.A(155665) + range.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
			}
			application.SendKeys(VH.A(49606), false);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
		range = null;
	}
}
