using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing.Share;

public sealed class Base
{
	[MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]
	public static string SaveTempWorkbook(Microsoft.Office.Interop.Excel.Workbook wbOrig, Microsoft.Office.Interop.Excel.Workbook wbTemp)
	{
		string text2 = default(string);
		try
		{
			XlFileFormat xlFileFormat;
			string text;
			if (Operators.CompareString(wbOrig.Path, "", TextCompare: false) == 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				xlFileFormat = XlFileFormat.xlOpenXMLWorkbook;
				text = VH.A(98730);
			}
			else
			{
				xlFileFormat = wbOrig.FileFormat;
				text = Path.GetExtension(wbOrig.Name);
			}
			text2 = Path.Combine(global::A.I.A.FileSystem.SpecialDirectories.Temp, VH.A(99907) + text);
			try
			{
				FileSystem.Kill(text2);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Microsoft.Office.Interop.Excel.Application application = wbOrig.Application;
			XlCalculation calculation = application.Calculation;
			bool calculateBeforeSave = application.CalculateBeforeSave;
			application.Calculation = XlCalculation.xlCalculationManual;
			application.CalculateBeforeSave = false;
			wbTemp.SaveAs(text2, xlFileFormat, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			application.Calculation = calculation;
			application.CalculateBeforeSave = calculateBeforeSave;
			_ = null;
			wbTemp.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		return text2;
	}

	public static void FlattenWorkbookPublic()
	{
		if (!A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = default(Microsoft.Office.Interop.Excel.Workbook);
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
			Microsoft.Office.Interop.Excel.Application application2;
			Microsoft.Office.Interop.Excel.Application application = (application2 = MH.A.Application);
			application2.ScreenUpdating = false;
			application2.EnableEvents = false;
			if (activeWorkbook == null)
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
				activeWorkbook = application2.ActiveWorkbook;
			}
			application2 = null;
			Base.A(activeWorkbook);
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			_ = null;
			activeWorkbook = null;
			return;
		}
	}

	internal static void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		try
		{
			IEnumerator enumerator = A.Worksheets.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range cells = ((Worksheet)enumerator.Current).Cells;
					cells.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					cells.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					_ = null;
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
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DeleteCharts(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		if (wb == null)
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
			wb = MH.A.Application.Application.ActiveWorkbook;
		}
		try
		{
			if (wb.Charts.Count > 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					wb.Charts.Delete();
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
		wb = null;
	}

	public static void DeleteAllCharts()
	{
		if (!A())
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
			DeleteChartsPublic();
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.DisplayAlerts = false;
			DeleteCharts();
			application.DisplayAlerts = true;
			_ = null;
			return;
		}
	}

	public static void DeleteWatchesPublic()
	{
		if (!A())
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			_ = null;
			Base.A();
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			_ = null;
			return;
		}
	}

	internal static void A()
	{
		checked
		{
			try
			{
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				for (int i = application.Watches.Count - 1; i >= 0; i += -1)
				{
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(application.Watches[i].Source, null, VH.A(99950), new object[0], null, null, null), null, VH.A(8701), new object[0], null, null, null), null, VH.A(99969), new object[0], null, null, null), application.ActiveWorkbook.FullName, TextCompare: false))
					{
						application.Watches[i].Delete();
					}
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
					application = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void DeleteCommentsPublic()
	{
		if (!A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		try
		{
			IEnumerator enumerator = activeWorkbook.Worksheets.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Base.A((Worksheet)enumerator.Current);
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		activeWorkbook = null;
	}

	internal static void A(Worksheet A)
	{
		A.Cells.ClearComments();
	}

	public static void ColorFontsBlackPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					clsImportExport.RecolorFonts(((Worksheet)enumerator.Current).UsedRange, false);
				}
			}
			finally
			{
				if (enumerator is IDisposable)
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
		application.ScreenUpdating = true;
		application = null;
	}

	public static void DeleteChartsPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					B((Worksheet)enumerator.Current);
				}
			}
			finally
			{
				if (enumerator is IDisposable)
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
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		activeWorkbook = null;
	}

	internal static void B(Worksheet A)
	{
		try
		{
			NewLateBinding.LateCall(A.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DeleteInkPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					C((Worksheet)enumerator.Current);
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
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
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
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
	}

	internal static void C(Worksheet A)
	{
		try
		{
			Shape shape;
			for (int i = A.Shapes.Count; i >= 1; shape = null, i = checked(i + -1))
			{
				shape = A.Shapes.Item(i);
				if (shape.Type != MsoShapeType.msoInk)
				{
					if (shape.Type != MsoShapeType.msoInkComment)
					{
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
				}
				shape.Delete();
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(Shape A)
	{
		Shape shape = A;
		if (shape.Type == MsoShapeType.msoGroup)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape a = (Shape)enumerator.Current;
					this.A(a);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_004b;
					}
					continue;
					end_IL_004b:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (1)
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
		else
		{
			if (shape.Type != MsoShapeType.msoInk)
			{
				if (shape.Type != MsoShapeType.msoInkComment)
				{
					goto IL_009b;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			shape.Delete();
		}
		goto IL_009b;
		IL_009b:
		shape = null;
	}

	internal static void D(Worksheet A)
	{
		try
		{
			A.DisplayPageBreaks = false;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void HideGridlinesPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					E((Worksheet)enumerator.Current);
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
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
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
		NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		objectValue = null;
		activeWorkbook = null;
	}

	internal static void E(Worksheet A)
	{
		try
		{
			A.Activate();
			A.Application.ActiveWindow.DisplayGridlines = false;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ZoomTo100Public()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					F((Worksheet)enumerator.Current);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
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
		NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		objectValue = null;
		activeWorkbook = null;
	}

	internal static void F(Worksheet A)
	{
		try
		{
			A.Activate();
			A.Application.ActiveWindow.Zoom = 100;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ReturnToCellA1Public()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					G((Worksheet)enumerator.Current);
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
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
		NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		objectValue = null;
		activeWorkbook = null;
	}

	internal static void G(Worksheet A)
	{
		try
		{
			A.Activate();
			NewLateBinding.LateCall(A.Cells[1, 1], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			A.Application.Goto(RuntimeHelpers.GetObjectValue(A.Cells[1, 1]), true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void CleanCellsPublic()
	{
		if (!A())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Microsoft.Office.Interop.Excel.Application application2 = application;
			application2.ScreenUpdating = false;
			application2.EnableEvents = false;
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
			_ = null;
			object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
			try
			{
				try
				{
					enumerator = activeWorkbook.Worksheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						A(((Worksheet)enumerator.Current).UsedRange);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0088;
						}
						continue;
						end_IL_0088:
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
			NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			Microsoft.Office.Interop.Excel.Application application3 = application;
			application3.ScreenUpdating = true;
			application3.EnableEvents = true;
			_ = null;
			application = null;
			objectValue = null;
			activeWorkbook = null;
			return;
		}
	}

	internal static void A(Range A)
	{
		try
		{
			Clean.Cells(A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void VeryHidePublic()
	{
		if (A())
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Microsoft.Office.Interop.Excel.Application application2 = application;
			application2.ScreenUpdating = false;
			application2.EnableEvents = false;
			application2.DisplayAlerts = false;
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
			_ = null;
			try
			{
				B(activeWorkbook);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Microsoft.Office.Interop.Excel.Application application3 = application;
			application3.DisplayAlerts = true;
			application3.ScreenUpdating = true;
			application3.EnableEvents = true;
			_ = null;
			application = null;
			activeWorkbook = null;
		}
	}

	internal static void B(Microsoft.Office.Interop.Excel.Workbook A)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Sheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetHidden, TextCompare: false))
					{
						continue;
					}
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
					NewLateBinding.LateSet(objectValue, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetVeryHidden }, null, null);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
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
	}

	public static void DeleteHiddenRowsColsPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				H((Worksheet)enumerator.Current);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
	}

	internal static void H(Worksheet A)
	{
		Range range;
		try
		{
			range = ((_Worksheet)A).get_Range((object)(VH.A(99986) + Conversions.ToString(A.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, RuntimeHelpers.GetObjectValue(Missing.Value)).Row)), RuntimeHelpers.GetObjectValue(Missing.Value));
			Range range2 = range.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			range.EntireRow.Hidden = false;
			range2.EntireRow.Hidden = true;
			try
			{
				range.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireRow.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range.EntireRow.Hidden = false;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[1, 1]), RuntimeHelpers.GetObjectValue(A.Cells[1, A.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, RuntimeHelpers.GetObjectValue(Missing.Value)).Column]));
			Range range3 = range.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			range.EntireColumn.Hidden = false;
			range3.EntireColumn.Hidden = true;
			try
			{
				range.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireColumn.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			range.EntireColumn.Hidden = false;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		range = null;
	}

	public static void CollapseGroupedRowsColsPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				I((Worksheet)enumerator.Current);
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
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
	}

	internal static void I(Worksheet A)
	{
		try
		{
			A.Outline.ShowLevels(1, RuntimeHelpers.GetObjectValue(Missing.Value));
			A.Outline.ShowLevels(RuntimeHelpers.GetObjectValue(Missing.Value), 1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void J(Worksheet A)
	{
		try
		{
			A.Outline.ShowLevels(8, RuntimeHelpers.GetObjectValue(Missing.Value));
			A.Outline.ShowLevels(RuntimeHelpers.GetObjectValue(Missing.Value), 8);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void ResetPrintAreasPublic()
	{
		if (!A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		application2.PrintCommunication = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		object objectValue = RuntimeHelpers.GetObjectValue(activeWorkbook.ActiveSheet);
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWorkbook.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					K((Worksheet)enumerator.Current);
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
		NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		application3.PrintCommunication = true;
		_ = null;
		application = null;
		objectValue = null;
		activeWorkbook = null;
	}

	internal static void K(Worksheet A)
	{
		try
		{
			A.PageSetup.PrintArea = "";
			A.Activate();
			A.Application.ActiveWindow.View = XlWindowView.xlNormalView;
			A.DisplayPageBreaks = false;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DeleteHiddenSheetsPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.ScreenUpdating = false;
		application2.EnableEvents = false;
		application2.DisplayAlerts = false;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
		_ = null;
		try
		{
			C(activeWorkbook);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Microsoft.Office.Interop.Excel.Application application3 = application;
		application3.DisplayAlerts = true;
		application3.ScreenUpdating = true;
		application3.EnableEvents = true;
		_ = null;
		application = null;
		activeWorkbook = null;
	}

	internal static void C(Microsoft.Office.Interop.Excel.Workbook A)
	{
		try
		{
			IEnumerator enumerator = A.Sheets.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					if (!Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(objectValue, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
					{
						continue;
					}
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
					NewLateBinding.LateCall(objectValue, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void BreakHyperlinksPublic()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		try
		{
			IEnumerator enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					L((Worksheet)enumerator.Current);
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
	}

	internal static void L(Worksheet A)
	{
		try
		{
			for (int i = A.Hyperlinks.Count; i >= 1; i = checked(i + -1))
			{
				A.Hyperlinks.get_Item((object)i).Delete();
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
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void A(Microsoft.Office.Interop.Excel.Workbook A, bool B)
	{
		if (!B)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)A.CustomDocumentProperties).GetEnumerator();
				while (enumerator.MoveNext())
				{
					DocumentProperty documentProperty = (DocumentProperty)enumerator.Current;
					try
					{
						documentProperty.Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_005e;
					}
					continue;
					end_IL_005e:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
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
		foreach (DocumentProperty item in (IEnumerable)A.BuiltinDocumentProperties)
		{
			try
			{
				switch (item.Type)
				{
				case MsoDocProperties.msoPropertyTypeString:
					item.Value = string.Empty;
					break;
				case MsoDocProperties.msoPropertyTypeDate:
					try
					{
						if (!(item.Value is DateTime))
						{
							break;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							item.Value = DateTime.MinValue;
							break;
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					break;
				case MsoDocProperties.msoPropertyTypeNumber:
				case MsoDocProperties.msoPropertyTypeBoolean:
				case MsoDocProperties.msoPropertyTypeFloat:
					break;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
	}

	internal static void D(Microsoft.Office.Interop.Excel.Workbook A)
	{
		for (int i = A.CustomXMLParts.Count; i >= 1; i = checked(i + -1))
		{
			try
			{
				A.CustomXMLParts[i].Delete();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
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
			return;
		}
	}

	public static void CheckFormulaErrors()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.Workbooks.Count > 0)
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
			application.ScreenUpdating = false;
			IEnumerator enumerator = default(IEnumerator);
			bool flag = default(bool);
			Range range;
			try
			{
				enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Worksheet worksheet = (Worksheet)enumerator.Current;
						range = null;
						try
						{
							range = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlErrors);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (range == null)
						{
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							flag = true;
							break;
						}
						break;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00ad;
						}
						continue;
						end_IL_00ad:
						break;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			application.ScreenUpdating = true;
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
				Forms.WarningMessage(VH.A(99995));
			}
			else
			{
				Forms.InfoMessage(VH.A(100072));
			}
			range = null;
		}
		application = null;
	}

	public static void RemoveNames()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		Microsoft.Office.Interop.Excel.Names names = activeWorkbook.Names;
		_ = names.Count;
		if (names.Count > 1000)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(VH.A(100155));
					return;
				}
			}
		}
		if (names.Count > 100 && MessageBox.Show(VH.A(100295), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
		{
			application = null;
			activeWorkbook = null;
			return;
		}
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		int num = A(activeWorkbook);
		try
		{
			activeWorkbook = application.ActiveWorkbook;
			for (int i = names.Count; i >= 1; i = checked(i + -1))
			{
				Name name = names.Item(i, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				ExcelAddIn1.Formulas.Names.Unapply(name);
				name.Delete();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (A(activeWorkbook) > num)
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
			Forms.WarningMessage(VH.A(100653));
		}
		application.ScreenUpdating = true;
		application.EnableEvents = true;
		application = null;
		activeWorkbook = null;
		names = null;
	}

	private static int A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		int num = default(int);
		Range range;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Worksheets.GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					range = null;
					try
					{
						range = Base.A((Worksheet)objectValue);
						if (range != null)
						{
							num = Conversions.ToInteger(Operators.AddObject(num, range.Cells.CountLarge));
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
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
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		range = null;
		return num;
	}

	internal static Range A(Worksheet A)
	{
		Range result = default(Range);
		try
		{
			result = A.Cells.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlErrors);
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)5, (Restriction)1, false);
	}
}
