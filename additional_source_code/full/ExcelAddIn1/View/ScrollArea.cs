using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class ScrollArea
{
	public static void Lock()
	{
		if (!Licensing.AllowAdvancedViewOperation())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = MH.A.Application;
			Worksheet worksheet;
			Range c;
			try
			{
				worksheet = (Worksheet)application.ActiveSheet;
				if (!worksheet.ProtectContents)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						c = (Range)application.Selection;
						A(application, worksheet, c, D: true, VH.A(174893));
						break;
					}
				}
				else
				{
					A();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application = null;
			worksheet = null;
			c = null;
			return;
		}
	}

	public static void Unlock()
	{
		if (!Licensing.AllowAdvancedViewOperation())
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
			Application application = MH.A.Application;
			Worksheet worksheet;
			Range range;
			try
			{
				worksheet = (Worksheet)application.ActiveSheet;
				if (!worksheet.ProtectContents)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						range = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (Operators.ConditionalCompareObjectGreater(range.CountLarge, 1, TextCompare: false))
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
							Range range2 = range.Areas[range.Areas.Count];
							range = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(range.Areas[1].Cells[1, 1]), RuntimeHelpers.GetObjectValue(range2.Cells[RuntimeHelpers.GetObjectValue(range2.Rows.CountLarge), RuntimeHelpers.GetObjectValue(range2.Columns.CountLarge)]));
							range2 = null;
						}
						A(application, worksheet, range, D: false, VH.A(174926));
						break;
					}
				}
				else
				{
					A();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application = null;
			worksheet = null;
			range = null;
			return;
		}
	}

	private static void A(Application A, Worksheet B, Range C, bool D, string E)
	{
		A.ScreenUpdating = false;
		try
		{
			Worksheet worksheet = B;
			if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(C.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null), 1, TextCompare: false))
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
				((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[1, Operators.SubtractObject(NewLateBinding.LateGet(C.Cells[1, 1], null, VH.A(41354), new object[0], null, null, null), 1)])).EntireColumn.Hidden = D;
			}
			if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(C.Cells[1, 1], null, VH.A(41347), new object[0], null, null, null), 1, TextCompare: false))
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
				((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[Operators.SubtractObject(NewLateBinding.LateGet(C.Cells[1, 1], null, VH.A(41347), new object[0], null, null, null), 1), 1])).EntireRow.Hidden = D;
			}
			((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, Operators.AddObject(NewLateBinding.LateGet(C.Columns[RuntimeHelpers.GetObjectValue(C.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null), 1)]), RuntimeHelpers.GetObjectValue(worksheet.Cells[1, RuntimeHelpers.GetObjectValue(worksheet.Columns.CountLarge)])).EntireColumn.Hidden = D;
			((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[Operators.AddObject(NewLateBinding.LateGet(C.Rows[RuntimeHelpers.GetObjectValue(C.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null), 1), 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[RuntimeHelpers.GetObjectValue(worksheet.Rows.CountLarge), 1])).EntireRow.Hidden = D;
			worksheet = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, E);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		A.ScreenUpdating = true;
	}

	private static void A()
	{
		Forms.WarningMessage(VH.A(154094));
	}
}
