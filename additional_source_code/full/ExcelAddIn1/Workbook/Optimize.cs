using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

public sealed class Optimize
{
	public static void NameScrubber()
	{
		if (!Licensing.AllowOptimizationOperation())
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Microsoft.Office.Interop.Excel.Application application2 = application;
			if (application2.Workbooks.Count == 0)
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
				int count = application2.ActiveWorkbook.Names.Count;
				if (count == 0)
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
					Forms.InfoMessage(VH.A(182343));
				}
				else if (Environment.IsEditMode(application))
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
					Forms.WarningMessage(VH.A(182416));
				}
				else
				{
					if (count > 25000)
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
						if (!application2.ActiveWorkbook.Saved)
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
							Forms.WarningMessage(VH.A(182475));
							goto IL_0107;
						}
					}
					new wpfNameScrubber().ShowDialog();
					_ = null;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(182586));
				}
				goto IL_0107;
				IL_0107:
				application2 = null;
				application = null;
				return;
			}
		}
	}

	public static void StyleScrubber()
	{
		if (!Licensing.AllowOptimizationOperation())
		{
			return;
		}
		bool flag = false;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activeWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				if (!((Worksheet)enumerator.Current).ProtectContents)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(VH.A(182613));
					flag = true;
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
		if (!flag)
		{
			if (activeWorkbook.MultiUserEditing)
			{
				if (MessageBox.Show(VH.A(182734), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
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
					application.DisplayAlerts = false;
					try
					{
						activeWorkbook.ExclusiveAccess();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						MessageBox.Show(VH.A(182942) + ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
						clsReporting.LogException(ex2);
						flag = true;
						ProjectData.ClearProjectError();
					}
					application.DisplayAlerts = true;
				}
				else
				{
					flag = true;
				}
			}
			if (!flag)
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
				new wpfStyleScrubber().ShowDialog();
				_ = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(180625));
			}
		}
		activeWorkbook = null;
		application = null;
	}

	public static void CleanUpUsedRanges()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		int intOrigCellCount = default(int);
		int intNewCellCount = default(int);
		Worksheet worksheet = default(Worksheet);
		long maxRows = default(long);
		int maxColumns = default(int);
		Worksheet worksheet2 = default(Worksheet);
		DialogResult dialogResult = default(DialogResult);
		object instance = default(object);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Licensing.AllowOptimizationOperation())
					{
						goto end_IL_0000;
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
					goto IL_0021;
				case 920:
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
							goto IL_0026;
						case 5:
							goto IL_002b;
						case 6:
							goto IL_0032;
						case 7:
							goto IL_0044;
						case 8:
							goto IL_0064;
						case 9:
							goto IL_0078;
						case 10:
							goto IL_009f;
						case 11:
							goto IL_00b1;
						case 13:
							goto IL_00c4;
						case 12:
						case 14:
							goto IL_00e9;
						case 15:
							goto IL_00fe;
						case 16:
							goto IL_0109;
						case 17:
							goto IL_0114;
						case 18:
							goto IL_011f;
						case 19:
							goto IL_012a;
						case 20:
							goto IL_0135;
						case 21:
							goto IL_014b;
						case 22:
							goto IL_016d;
						case 23:
							goto IL_01b0;
						case 24:
							goto IL_01f8;
						case 25:
							goto IL_01fb;
						case 26:
							goto IL_0227;
						case 27:
							goto IL_022e;
						case 28:
							goto IL_0238;
						case 30:
							goto IL_0250;
						case 31:
							goto IL_0261;
						case 29:
						case 32:
							goto IL_0264;
						case 33:
							goto IL_027c;
						case 34:
							goto IL_0294;
						case 35:
							goto IL_029f;
						case 36:
							goto IL_02aa;
						case 37:
							goto IL_02b5;
						case 38:
							goto IL_02c5;
						case 39:
							goto IL_02d0;
						case 40:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 41:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_02b5:
					num2 = 37;
					application.StatusBar = false;
					goto IL_02c5;
					IL_02c5:
					num2 = 38;
					application.EnableCancelKey = XlEnableCancelKey.xlInterrupt;
					goto IL_02d0;
					IL_02d0:
					num2 = 39;
					FinishCleanUsedRanges(intOrigCellCount, intNewCellCount);
					break;
					IL_0250:
					num2 = 30;
					CleanUsedRange(worksheet, maxRows, maxColumns, ref intOrigCellCount, ref intNewCellCount);
					goto IL_0261;
					IL_0021:
					num2 = 3;
					intOrigCellCount = 0;
					goto IL_0026;
					IL_0026:
					num2 = 4;
					intNewCellCount = 0;
					goto IL_002b;
					IL_002b:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0032;
					IL_0032:
					num2 = 6;
					application = MH.A.Application;
					goto IL_0044;
					IL_0044:
					num2 = 7;
					if (application.Workbooks.Count <= 0)
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
						break;
					}
					goto IL_0064;
					IL_0227:
					num2 = 26;
					worksheet2 = worksheet;
					goto IL_022e;
					IL_0064:
					num2 = 8;
					if (!application.ActiveWorkbook.Saved)
					{
						goto IL_0078;
					}
					goto IL_00c4;
					IL_0078:
					num2 = 9;
					dialogResult = MessageBox.Show(VH.A(183067), VH.A(40448), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
					goto IL_009f;
					IL_009f:
					num2 = 10;
					if (dialogResult == DialogResult.Yes)
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
						goto IL_00b1;
					}
					goto IL_00e9;
					IL_022e:
					num2 = 27;
					ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
					goto IL_0238;
					IL_00b1:
					num2 = 11;
					application.ActiveWorkbook.Save();
					goto IL_00e9;
					IL_00c4:
					num2 = 13;
					dialogResult = MessageBox.Show(VH.A(183593), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
					goto IL_00e9;
					IL_00e9:
					num2 = 14;
					if (dialogResult == DialogResult.Cancel)
					{
						break;
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
					goto IL_00fe;
					IL_0238:
					num2 = 28;
					if (!worksheet2.ProtectContents)
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
						goto IL_0250;
					}
					goto IL_0264;
					IL_00fe:
					num2 = 15;
					application.EnableCancelKey = XlEnableCancelKey.xlDisabled;
					goto IL_0109;
					IL_0109:
					num2 = 16;
					application.ScreenUpdating = false;
					goto IL_0114;
					IL_0114:
					num2 = 17;
					application.EnableEvents = false;
					goto IL_011f;
					IL_011f:
					num2 = 18;
					application.DisplayAlerts = false;
					goto IL_012a;
					IL_012a:
					num2 = 19;
					application.DisplayStatusBar = true;
					goto IL_0135;
					IL_0135:
					num2 = 20;
					application.StatusBar = VH.A(184035);
					goto IL_014b;
					IL_014b:
					num2 = 21;
					instance = application.ActiveWorkbook.Worksheets[1];
					goto IL_016d;
					IL_016d:
					num2 = 22;
					maxRows = Conversions.ToLong(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(152043), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null));
					goto IL_01b0;
					IL_01b0:
					num2 = 23;
					maxColumns = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(152073), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null));
					goto IL_01f8;
					IL_01f8:
					instance = null;
					goto IL_01fb;
					IL_01fb:
					num2 = 25;
					enumerator = application.ActiveWorkbook.Worksheets.GetEnumerator();
					goto IL_0267;
					IL_0267:
					if (enumerator.MoveNext())
					{
						worksheet = (Worksheet)enumerator.Current;
						goto IL_0227;
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
					goto IL_027c;
					IL_0264:
					num2 = 32;
					goto IL_0267;
					IL_027c:
					num2 = 33;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0294;
					IL_0261:
					worksheet2 = null;
					goto IL_0264;
					IL_0294:
					num2 = 34;
					application.ScreenUpdating = true;
					goto IL_029f;
					IL_029f:
					num2 = 35;
					application.DisplayAlerts = true;
					goto IL_02aa;
					IL_02aa:
					num2 = 36;
					application.EnableEvents = true;
					goto IL_02b5;
					end_IL_0000_3:
					break;
				}
				application = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 920;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void CleanUsedRange(Worksheet ws, long maxRows, int maxColumns, ref int intOrigCellCount, ref int intNewCellCount)
	{
		Worksheet worksheet = ws;
		checked
		{
			if (Operators.ConditionalCompareObjectNotEqual(worksheet.UsedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), NewLateBinding.LateGet(worksheet.Cells[1, 1], null, VH.A(5814), new object[0], null, null, null), TextCompare: false))
			{
				intOrigCellCount = Conversions.ToInteger(Operators.AddObject(intOrigCellCount, worksheet.UsedRange.Cells.CountLarge));
				(long, int) lastRowColumn = GetLastRowColumn(ws);
				((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, lastRowColumn.Item2 + 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[1, maxColumns])).EntireColumn.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
				((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[lastRowColumn.Item1 + 1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[maxRows, 1])).EntireRow.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range usedRange = ws.UsedRange;
				intNewCellCount = Conversions.ToInteger(Operators.AddObject(intNewCellCount, usedRange.Cells.CountLarge));
				usedRange = null;
			}
			worksheet = null;
		}
	}

	public static (long LastRow, int LastColumn) GetLastRowColumn(Worksheet ws)
	{
		Worksheet worksheet = ws;
		long num = worksheet.Cells.Find(VH.A(75231), ((_Worksheet)worksheet).get_Range((object)VH.A(78844), RuntimeHelpers.GetObjectValue(Missing.Value)), XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Row;
		int num2 = worksheet.Cells.Find(VH.A(75231), ((_Worksheet)worksheet).get_Range((object)VH.A(78844), RuntimeHelpers.GetObjectValue(Missing.Value)), XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Column;
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = worksheet.Shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shape shape = (Shape)enumerator.Current;
					long num3 = 0L;
					long num4 = 0L;
					try
					{
						num3 = shape.TopLeftCell.Row;
						num4 = shape.TopLeftCell.Column;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					if (num3 <= 0)
					{
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (num4 <= 0)
					{
						continue;
					}
					for (; !Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(worksheet.Cells[num3, num4], null, VH.A(57409), new object[0], null, null, null), shape.Top + shape.Height, TextCompare: false); num3++)
					{
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
					if (num3 > num)
					{
						num = num3;
					}
					for (; !Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(worksheet.Cells[num3, num4], null, VH.A(56582), new object[0], null, null, null), shape.Left + shape.Width, TextCompare: false); num4++)
					{
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					if (num4 <= num2)
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
					num2 = (int)num4;
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
			worksheet = null;
			return (LastRow: num, LastColumn: num2);
		}
	}

	public static void FinishCleanUsedRanges(int intOrigCellCount, int intNewCellCount)
	{
		if (intNewCellCount > 0)
		{
			Forms.InfoMessage(VH.A(184082) + (1.0 - (double)intNewCellCount / (double)intOrigCellCount).ToString(VH.A(75211)) + VH.A(64021));
		}
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(184137));
	}

	public static void DeleteUnusedCustomNumberFormats()
	{
		if (!Licensing.AllowOptimizationOperation())
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			int num4 = default(int);
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
				Range range = null;
				Range range2 = null;
				string[] array;
				bool[] arySrc;
				Range range3;
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook;
				try
				{
					activeWorkbook = application.ActiveWorkbook;
					if (activeWorkbook.Worksheets.Count == 0)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							throw new Exception();
						}
					}
					if (application.ActiveSheet is Worksheet)
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
						if (application.Selection is Range)
						{
							range = (Range)application.Selection;
							range2 = application.ActiveCell;
						}
					}
					if (activeWorkbook.Saved)
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
						if (activeWorkbook.Path.Length != 0)
						{
							goto IL_02ab;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					DialogResult dialogResult = MessageBox.Show(VH.A(184172), VH.A(40448), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
					if (dialogResult == DialogResult.Cancel)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							throw new Exception();
						}
					}
					if (dialogResult == DialogResult.Yes)
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
						if (Strings.Len(activeWorkbook.Path) == 0)
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
							application.Dialogs[XlBuiltInDialog.xlDialogSaveAs].Show(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						else
						{
							application.ActiveWorkbook.Save();
						}
					}
					goto IL_02ab;
					IL_02ab:
					if (MessageBox.Show(VH.A(184328), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
					{
						throw new Exception();
					}
					application.Cursor = XlMousePointer.xlWait;
					application.ScreenUpdating = false;
					array = new string[1001];
					arySrc = new bool[1001];
					object obj;
					if (range2 != null)
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
						obj = ((_Worksheet)range2.Worksheet).get_Range((object)VH.A(78844), RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					else
					{
						obj = NewLateBinding.LateGet(activeWorkbook.Worksheets[1], null, VH.A(41315), new object[1] { VH.A(78844) }, null, null, null);
					}
					range3 = (Range)obj;
					range3.Select();
					string numberFormatLocal = Conversions.ToString(range3.NumberFormatLocal);
					range3.NumberFormat = VH.A(20593);
					array[0] = VH.A(20593);
					string arg = Conversions.ToString(range3.NumberFormatLocal);
					int num = 1;
					do
					{
						application.SendKeys(VH.A(184858), RuntimeHelpers.GetObjectValue(Missing.Value));
						application.Dialogs[XlBuiltInDialog.xlDialogFormatNumber].Show(arg, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						array[num] = Conversions.ToString(range3.NumberFormat);
						arg = Conversions.ToString(range3.NumberFormatLocal);
						num++;
					}
					while (Operators.CompareString(array[num - 1], array[num - 2], TextCompare: false) != 0);
					range3.NumberFormatLocal = numberFormatLocal;
					array = (string[])Utils.CopyArray(array, new string[num - 2 + 1]);
					arySrc = (bool[])Utils.CopyArray(arySrc, new bool[num - 2 + 1]);
					try
					{
						enumerator = activeWorkbook.Worksheets.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Worksheet worksheet = (Worksheet)enumerator.Current;
							try
							{
								enumerator2 = worksheet.UsedRange.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									object objectValue = RuntimeHelpers.GetObjectValue(enumerator2.Current);
									int num2 = Information.UBound(array);
									num = 0;
									while (true)
									{
										if (num <= num2)
										{
											if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(57240), new object[0], null, null, null), array[num], TextCompare: false))
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
												arySrc[num] = true;
												break;
											}
											num++;
											continue;
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
										break;
									}
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
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_06ac;
							}
							continue;
							end_IL_06ac:
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
					int num3 = Information.UBound(array);
					for (num = 0; num <= num3; num++)
					{
						if (arySrc[num])
						{
							continue;
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
						try
						{
							activeWorkbook.DeleteNumberFormat(array[num]);
							num4++;
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
							continue;
						}
						if (range2 != null)
						{
							range.Select();
							range2.Activate();
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(184899));
						Forms.InfoMessage(Conversions.ToString(num4) + VH.A(184940));
						break;
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				array = null;
				arySrc = null;
				application.Cursor = XlMousePointer.xlDefault;
				application.ScreenUpdating = true;
				application = null;
				range3 = null;
				activeWorkbook = null;
				return;
			}
		}
	}
}
