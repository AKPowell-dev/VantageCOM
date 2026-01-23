using System;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class Arrows
{
	public static void Clear()
	{
		ClearArrowsOnActiveSheet();
		TraceAll.ShowingPrecedents = false;
		TraceAll.ShowingDependents = false;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(52507));
	}

	public static void ClearArrowsOnActiveSheet()
	{
		try
		{
			NewLateBinding.LateCall(MH.A.Application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DisplayObjects(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		if (wb.DisplayDrawingObjects != XlDisplayDrawingObjects.xlHide)
		{
			return;
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
			wb.DisplayDrawingObjects = XlDisplayDrawingObjects.xlDisplayShapes;
			return;
		}
	}
}
