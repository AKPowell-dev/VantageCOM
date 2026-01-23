using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Sheets;

public sealed class Move
{
	internal static void A()
	{
		Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		bool flag = false;
		application.ScreenUpdating = false;
		Microsoft.Office.Interop.Excel.Sheets selectedSheets;
		object objectValue;
		try
		{
			objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
			selectedSheets = application.ActiveWindow.SelectedSheets;
			int num = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(selectedSheets.get_Item((object)1), null, VH.A(48135), new object[0], null, null, null), 1));
			while (true)
			{
				if (num >= 1)
				{
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(activeWorkbook.Sheets[num], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
					{
						selectedSheets.Move(RuntimeHelpers.GetObjectValue(activeWorkbook.Sheets[num]), RuntimeHelpers.GetObjectValue(Missing.Value));
						flag = true;
						break;
					}
					num = checked(num + -1);
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
			if (!flag)
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
				selectedSheets.Move(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(activeWorkbook.Sheets[activeWorkbook.Sheets.Count]));
			}
			if (selectedSheets.Count > 1)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					selectedSheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
					NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
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
		application.ScreenUpdating = true;
		application = null;
		activeWorkbook = null;
		selectedSheets = null;
		objectValue = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(102087));
	}

	internal static void B()
	{
		Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		bool flag = false;
		application.ScreenUpdating = false;
		Microsoft.Office.Interop.Excel.Sheets selectedSheets;
		object objectValue;
		try
		{
			objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
			selectedSheets = application.ActiveWindow.SelectedSheets;
			int num = Conversions.ToInteger(Operators.AddObject(NewLateBinding.LateGet(selectedSheets.get_Item((object)1), null, VH.A(48135), new object[0], null, null, null), selectedSheets.Count));
			int count = activeWorkbook.Sheets.Count;
			int num2 = num;
			while (true)
			{
				if (num2 <= count)
				{
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(activeWorkbook.Sheets[num2], null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
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
						selectedSheets.Move(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(activeWorkbook.Sheets[num2]));
						flag = true;
						break;
					}
					num2 = checked(num2 + 1);
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
				break;
			}
			if (!flag)
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
				selectedSheets.Move(RuntimeHelpers.GetObjectValue(activeWorkbook.Sheets[1]), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			if (selectedSheets.Count > 1)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					selectedSheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
					NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
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
		application.ScreenUpdating = true;
		application = null;
		activeWorkbook = null;
		selectedSheets = null;
		objectValue = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(102118));
	}
}
