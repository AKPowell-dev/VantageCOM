using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing;

public sealed class Helpers
{
	internal static void A(Range A, Worksheet B, Microsoft.Office.Interop.Excel.Application C, Microsoft.Office.Interop.Excel.Workbook D)
	{
		Range range = null;
		if (Operators.ConditionalCompareObjectGreater(A.Cells.CountLarge, 1, TextCompare: false))
		{
			try
			{
				range = A.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlErrors);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		else if (A.Application.WorksheetFunction.IsError(A))
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
			range = A;
		}
		C.DisplayAlerts = false;
		Worksheet worksheet = (Worksheet)D.Worksheets[B.Name];
		IEnumerator enumerator = A.Areas.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Range range2 = (Range)enumerator.Current;
				try
				{
					range2.Value2 = RuntimeHelpers.GetObjectValue(((_Worksheet)worksheet).get_Range((object)range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0131;
				}
				continue;
				end_IL_0131:
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
		worksheet = null;
		C.DisplayAlerts = true;
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = range.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Range range3 = (Range)enumerator2.Current;
				range3.Value2 = RuntimeHelpers.GetObjectValue(((_Worksheet)(Worksheet)D.Worksheets[range3.Worksheet.Name]).get_Range((object)range3.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Text);
				range3 = null;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_020e;
				}
				continue;
				end_IL_020e:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		range = null;
	}

	internal static bool A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		bool result = true;
		if (RuntimeHelpers.GetObjectValue(A.LinkSources(XlLink.xlExcelLinks)) != null)
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
			if (MessageBox.Show(VH.A(96889), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
			{
				result = false;
			}
		}
		return result;
	}

	internal static Microsoft.Office.Interop.Excel.Workbook A(Microsoft.Office.Interop.Excel.Workbook A, List<int> B)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		Microsoft.Office.Interop.Excel.Workbook workbook = Helpers.A(A, application);
		Microsoft.Office.Interop.Excel.Workbook workbook2 = workbook;
		for (int i = workbook2.Worksheets.Count; i >= 2; i = checked(i + -1))
		{
			NewLateBinding.LateCall(workbook2.Worksheets[i], null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
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
			((Worksheet)workbook2.Worksheets[1]).Name = VH.A(97163);
			NewLateBinding.LateCall(A.Sheets[B.ToArray()], null, VH.A(224), new object[1] { workbook2.Worksheets[1] }, new string[1] { VH.A(80163) }, null, null, IgnoreReturn: true);
			((Worksheet)workbook2.Worksheets[1]).Delete();
			workbook2.Sheets.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			workbook2 = null;
			application.ActiveWindow.Zoom = 100;
			application = null;
			return workbook;
		}
	}

	internal static Microsoft.Office.Interop.Excel.Workbook A(Microsoft.Office.Interop.Excel.Workbook A, Microsoft.Office.Interop.Excel.Application B)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = B.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			Themes.CopyWorkbookColorsAndFonts(A, workbook, true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		Environment.DisplayAlertsOff(B);
		try
		{
			workbook.Styles.Merge(A);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		Environment.DisplayAlertsOn(B);
		return workbook;
	}
}
