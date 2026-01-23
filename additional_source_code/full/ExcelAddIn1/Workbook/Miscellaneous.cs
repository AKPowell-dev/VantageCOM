using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

public sealed class Miscellaneous
{
	public static void Reopen(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		string fullName = default(string);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		XlCalculation calculation = default(XlCalculation);
		bool flag = default(bool);
		bool iteration = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Licensing.AllowRestrictedMode())
					{
						goto end_IL_0000;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0021;
				case 664:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_0037;
						case 6:
							goto IL_003c;
						case 7:
							goto IL_0041;
						case 8:
							goto IL_004f;
						case 9:
							goto IL_0061;
						case 10:
							goto IL_007f;
						case 11:
							goto IL_008c;
						case 12:
							goto IL_00d7;
						case 14:
							goto IL_00e3;
						case 13:
						case 15:
							goto IL_00e9;
						case 16:
							goto IL_00fd;
						case 17:
							goto IL_010a;
						case 18:
							goto IL_0115;
						case 19:
							goto IL_0120;
						case 20:
							goto IL_012a;
						case 21:
							goto IL_01e1;
						case 22:
							goto IL_01ec;
						case 23:
							goto IL_01f7;
						case 24:
							goto IL_020d;
						case 25:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 26:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0115:
					num2 = 18;
					fullName = wb.FullName;
					goto IL_0120;
					IL_0120:
					num2 = 19;
					Workbooks.ForceClose(wb, false);
					goto IL_012a;
					IL_012a:
					num2 = 20;
					application.Workbooks.Open(fullName, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_01e1;
					IL_01e1:
					num2 = 21;
					application.Calculation = calculation;
					goto IL_01ec;
					IL_0021:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0028;
					IL_0028:
					num2 = 4;
					application = MH.A.Application;
					goto IL_0037;
					IL_0037:
					num2 = 5;
					flag = false;
					goto IL_003c;
					IL_003c:
					num2 = 6;
					if (wb == null)
					{
						goto IL_0041;
					}
					goto IL_004f;
					IL_0041:
					num2 = 7;
					wb = application.ActiveWorkbook;
					goto IL_004f;
					IL_004f:
					num2 = 8;
					if (wb == null)
					{
						break;
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
					goto IL_0061;
					IL_01ec:
					num2 = 22;
					application.Iteration = iteration;
					goto IL_01f7;
					IL_0061:
					num2 = 9;
					if (Operators.CompareString(wb.Name, wb.FullName, TextCompare: false) != 0)
					{
						goto IL_007f;
					}
					goto IL_020d;
					IL_007f:
					num2 = 10;
					if (!wb.Saved)
					{
						goto IL_008c;
					}
					goto IL_00e9;
					IL_008c:
					num2 = 11;
					if (MessageBox.Show(VH.A(181437) + wb.Name + VH.A(117068), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
						goto IL_00d7;
					}
					goto IL_00e3;
					IL_01f7:
					num2 = 23;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(181526));
					goto IL_020d;
					IL_00d7:
					num2 = 12;
					wb.Saved = true;
					goto IL_00e9;
					IL_00e3:
					num2 = 14;
					flag = true;
					goto IL_00e9;
					IL_00e9:
					num2 = 15;
					if (!flag)
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
						goto IL_00fd;
					}
					goto IL_01f7;
					IL_020d:
					num2 = 24;
					JH.A((object)wb);
					break;
					IL_00fd:
					num2 = 16;
					calculation = application.Calculation;
					goto IL_010a;
					IL_010a:
					num2 = 17;
					iteration = application.Iteration;
					goto IL_0115;
					end_IL_0000_3:
					break;
				}
				num2 = 25;
				application = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 664;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	public static void ClearConstants()
	{
		if (MessageBox.Show(VH.A(181557), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			try
			{
				try
				{
					enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
						if (!(objectValue is Worksheet))
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
						try
						{
							((Worksheet)objectValue).UsedRange.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers).Clear();
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
							break;
						default:
							goto end_IL_00c6;
						}
						continue;
						end_IL_00c6:
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
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			return;
		}
	}

	private static void A()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.DisplayAlerts = false;
		try
		{
			NewLateBinding.LateCall(application.Worksheets[VH.A(181731)], null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.DisplayAlerts = true;
		application = null;
	}

	public static void Duplicate()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
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
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
			try
			{
				workbook = MH.A.Application.ActiveWorkbook;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (workbook == null)
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
				DuplicateWorkbook(workbook);
				workbook = null;
				return;
			}
		}
	}

	public static Microsoft.Office.Interop.Excel.Workbook DuplicateWorkbook(Microsoft.Office.Interop.Excel.Workbook wbOld)
	{
		Microsoft.Office.Interop.Excel.Workbook result = null;
		try
		{
			string text = ((wbOld.Path.Length <= 0) ? VH.A(98730) : Path.GetExtension(wbOld.FullName));
			string text2 = Path.Combine(Interaction.Environ(VH.A(86342)), VH.A(86439) + text);
			File.Delete(text2);
			wbOld.SaveCopyAs(text2);
			result = wbOld.Application.Workbooks.Add(text2);
			File.Delete(text2);
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(181776));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if (!EditMode.IsEditMode(wbOld.Application))
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
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
			}
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void OpenFolder(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
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
			if (wb == null)
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
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				if (application.Workbooks.Count > 0)
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
					wb = application.ActiveWorkbook;
				}
				application = null;
			}
			if (wb != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						try
						{
							clsFile.OpenExplorerToFile(wb.FullName);
							clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(181813));
							return;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							clsReporting.LogException(ex2);
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
			}
			Forms.WarningMessage(VH.A(181842));
			return;
		}
	}

	public static void GoToDialog(List<Range> RangeList)
	{
		new wpfGoTo(RangeList).ShowDialog();
		_ = null;
	}

	public static void CloseOthers(Microsoft.Office.Interop.Excel.Workbook wbThis = null)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook workbook;
		try
		{
			if (wbThis == null)
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
				wbThis = application.ActiveWorkbook;
			}
			Workbooks workbooks = application.Workbooks;
			for (int i = workbooks.Count; i >= 1; i = checked(i + -1))
			{
				workbook = workbooks.get_Item((object)i);
				if (workbook == wbThis)
				{
					continue;
				}
				if (!workbook.Saved)
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
					workbook.Activate();
				}
				workbook.Close(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				workbooks = null;
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (wbThis != null)
		{
			wbThis.Activate();
			wbThis = null;
		}
		workbook = null;
		application = null;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(181883));
	}

	public static bool DisplayObjects(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		bool result = true;
		if (wb.DisplayDrawingObjects == XlDisplayDrawingObjects.xlHide)
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
			if (MessageBox.Show(VH.A(181908), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
				wb.DisplayDrawingObjects = XlDisplayDrawingObjects.xlDisplayShapes;
			}
			else
			{
				result = false;
			}
		}
		return result;
	}

	public static bool IsProtectedView(bool SuppressMessages)
	{
		bool result = false;
		try
		{
			if (MH.A.Application.ActiveProtectedViewWindow != null)
			{
				if (!SuppressMessages)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(VH.A(182176));
				}
				result = true;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void CopyPath(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		try
		{
			if (wb == null)
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
				wb = MH.A.Application.ActiveWorkbook;
			}
			if (wb.Path.Length > 0)
			{
				clsClipboard.SetText(wb.FullName);
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(182279));
			}
			else
			{
				Forms.WarningMessage(VH.A(182298));
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

	public static void Calculate()
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = null;
		try
		{
			workbook = MH.A.Application.ActiveWorkbook;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (workbook == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			XlCalculation calculation = application.Calculation;
			application.ScreenUpdating = false;
			application.Calculation = XlCalculation.xlCalculationManual;
			try
			{
				try
				{
					enumerator = workbook.Worksheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Worksheet obj = (Worksheet)enumerator.Current;
						obj.EnableCalculation = false;
						obj.EnableCalculation = true;
						obj.Calculate();
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00a8;
						}
						continue;
						end_IL_00a8:
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
			application.Calculation = calculation;
			application.ScreenUpdating = true;
			application = null;
			return;
		}
	}

	public static void RemoveBuiltInProperties(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		try
		{
			if (wb == null)
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
				wb = MH.A.Application.ActiveWorkbook;
			}
			wb.RemovePersonalInformation = true;
			wb.RemoveDocumentInformation(XlRemoveDocInfoType.xlRDIAll);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((DocumentProperties)wb.BuiltinDocumentProperties).GetEnumerator();
				while (enumerator.MoveNext())
				{
					DocumentProperty documentProperty = (DocumentProperty)enumerator.Current;
					try
					{
						switch (documentProperty.Type)
						{
						case MsoDocProperties.msoPropertyTypeString:
							documentProperty.Value = "";
							break;
						case MsoDocProperties.msoPropertyTypeNumber:
						case MsoDocProperties.msoPropertyTypeDate:
						case MsoDocProperties.msoPropertyTypeFloat:
							documentProperty.Value = 0;
							break;
						case MsoDocProperties.msoPropertyTypeBoolean:
							documentProperty.Value = false;
							break;
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
					switch (6)
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
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}
}
