using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Reverse
{
	public static void Rows()
	{
		A(B, VH.A(171835), C: true);
	}

	public static void Columns()
	{
		A(A, VH.A(171860), C: false);
	}

	private static void A(Action<Worksheet, Range, string> A, string B, bool C)
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		Application application = MH.A.Application;
		Range range;
		if (application.Selection is Range)
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
			range = (Range)application.Selection;
			if (!Reverse.A(range))
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
				if (!EditMode.IsEditMode(application))
				{
					if (!C)
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
						if (Operators.ConditionalCompareObjectGreater(range.Columns.CountLarge, 1, TextCompare: false))
						{
							goto IL_00dd;
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
					if (C)
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
						if (Operators.ConditionalCompareObjectGreater(range.Rows.CountLarge, 1, TextCompare: false))
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
							goto IL_00dd;
						}
					}
				}
			}
			else
			{
				Forms.WarningMessage(VH.A(171891));
			}
			goto IL_018b;
		}
		goto IL_018d;
		IL_018d:
		application = null;
		return;
		IL_018b:
		range = null;
		goto IL_018d;
		IL_00dd:
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		Worksheet worksheet = default(Worksheet);
		try
		{
			worksheet = range.Worksheet;
			string arg = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			A(worksheet, range, arg);
			Core.LogActivity(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(ex2.Message);
			if (!worksheet.ProtectContents)
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
				clsReporting.LogException(ex2);
			}
			ProjectData.ClearProjectError();
		}
		finally
		{
			worksheet = null;
		}
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		goto IL_018b;
	}

	private static void A(Worksheet A, Range B, string C)
	{
		int num = Conversions.ToInteger(Operators.SubtractObject(B.Columns.CountLarge, 1));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			Range range = ((_Worksheet)A).get_Range((object)C, RuntimeHelpers.GetObjectValue(Missing.Value));
			((Range)range.Columns[RuntimeHelpers.GetObjectValue(range.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
			((Range)range.Columns[i, RuntimeHelpers.GetObjectValue(Missing.Value)]).Insert(XlDirection.xlToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
			range = null;
		}
	}

	private static void B(Worksheet A, Range B, string C)
	{
		int num = Conversions.ToInteger(Operators.SubtractObject(B.Rows.CountLarge, 1));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			Range range = ((_Worksheet)A).get_Range((object)C, RuntimeHelpers.GetObjectValue(Missing.Value));
			((Range)range.Rows[RuntimeHelpers.GetObjectValue(range.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
			((Range)range.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)]).Insert(XlDirection.xlDown, RuntimeHelpers.GetObjectValue(Missing.Value));
			range = null;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}

	private static bool A(Range A)
	{
		bool result;
		try
		{
			result = Conversions.ToBoolean(A.MergeCells);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = true;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
