using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

public sealed class QuickSave
{
	public enum SaveDownConvention
	{
		VersionNumber,
		DateStamp
	}

	public static void Save(Microsoft.Office.Interop.Excel.Workbook wb = null)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		XlCalculation calculation = default(XlCalculation);
		bool calculateBeforeSave = default(bool);
		bool flag = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 818:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0017;
						case 4:
							goto IL_0023;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_003e;
						case 7:
							goto IL_0048;
						case 8:
							goto IL_0052;
						case 9:
							goto IL_005c;
						case 10:
							goto IL_0072;
						case 11:
							goto IL_008b;
						case 12:
							goto IL_009b;
						case 13:
							goto IL_00be;
						case 14:
							goto IL_00c7;
						case 16:
							goto IL_00d2;
						case 15:
						case 17:
							goto IL_0240;
						case 18:
							goto IL_024f;
						case 19:
							goto IL_025f;
						case 20:
							goto IL_026a;
						case 21:
							goto IL_0275;
						case 22:
							goto IL_0281;
						case 23:
							goto IL_0284;
						case 24:
							goto IL_0295;
						case 25:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 26:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_026a:
					num2 = 20;
					application.Calculation = calculation;
					goto IL_0275;
					IL_0007:
					num2 = 2;
					application = MH.A.Application;
					goto IL_0017;
					IL_0017:
					num2 = 3;
					calculation = application.Calculation;
					goto IL_0023;
					IL_0023:
					num2 = 4;
					calculateBeforeSave = application.CalculateBeforeSave;
					goto IL_0030;
					IL_0030:
					num2 = 5;
					application.Calculation = XlCalculation.xlCalculationManual;
					goto IL_003e;
					IL_003e:
					num2 = 6;
					application.CalculateBeforeSave = false;
					goto IL_0048;
					IL_0048:
					num2 = 7;
					application.EnableEvents = false;
					goto IL_0052;
					IL_0052:
					num2 = 8;
					application.DisplayStatusBar = true;
					goto IL_005c;
					IL_005c:
					num2 = 9;
					application.StatusBar = VH.A(185031);
					goto IL_0072;
					IL_0072:
					num2 = 10;
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
						goto IL_008b;
					}
					goto IL_009b;
					IL_0275:
					num2 = 21;
					application.CalculateBeforeSave = calculateBeforeSave;
					goto IL_0281;
					IL_0281:
					application = null;
					goto IL_0284;
					IL_0284:
					num2 = 23;
					if (!flag)
					{
						break;
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
					goto IL_0295;
					IL_008b:
					num2 = 11;
					wb = application.ActiveWorkbook;
					goto IL_009b;
					IL_009b:
					num2 = 12;
					if (Operators.CompareString(wb.Name, wb.FullName, TextCompare: false) != 0)
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
						goto IL_00be;
					}
					goto IL_00d2;
					IL_0295:
					num2 = 24;
					wb.Saved = true;
					break;
					IL_00be:
					num2 = 13;
					wb.Save();
					goto IL_00c7;
					IL_00c7:
					num2 = 14;
					flag = true;
					goto IL_0240;
					IL_00d2:
					num2 = 16;
					application.Dialogs[XlBuiltInDialog.xlDialogSaveAs].Show(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0240;
					IL_0240:
					num2 = 17;
					application.StatusBar = "";
					goto IL_024f;
					IL_024f:
					num2 = 18;
					application.StatusBar = false;
					goto IL_025f;
					IL_025f:
					num2 = 19;
					application.EnableEvents = true;
					goto IL_026a;
					end_IL_0000_2:
					break;
				}
				num2 = 25;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(185062));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 818;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	public static void SaveAll()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		XlCalculation calculation = default(XlCalculation);
		bool calculateBeforeSave = default(bool);
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
					if (!Licensing.AllowRestrictedMode())
					{
						goto end_IL_0000;
					}
					goto IL_000c;
				case 521:
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
							goto IL_000c;
						case 4:
							goto IL_0013;
						case 5:
							goto IL_0023;
						case 6:
							goto IL_002f;
						case 7:
							goto IL_003a;
						case 8:
							goto IL_0048;
						case 9:
							goto IL_0052;
						case 10:
							goto IL_005d;
						case 11:
							goto IL_0068;
						case 12:
							goto IL_0073;
						case 13:
							goto IL_0087;
						case 14:
							goto IL_00ac;
						case 15:
							goto IL_00de;
						case 16:
							goto IL_00e8;
						case 17:
							goto IL_00f3;
						case 18:
							goto IL_010b;
						case 19:
							goto IL_0123;
						case 20:
							goto IL_0132;
						case 21:
							goto IL_0142;
						case 22:
							goto IL_014d;
						case 23:
							goto IL_0158;
						case 24:
							goto IL_0163;
						case 25:
							goto IL_016f;
						case 26:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 27:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_00ac:
					num2 = 14;
					if (Operators.CompareString(workbook.Name, workbook.FullName, TextCompare: false) != 0)
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
						goto IL_00de;
					}
					goto IL_00f3;
					IL_000c:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0013;
					IL_0013:
					num2 = 4;
					application = MH.A.Application;
					goto IL_0023;
					IL_0023:
					num2 = 5;
					calculation = application.Calculation;
					goto IL_002f;
					IL_002f:
					num2 = 6;
					calculateBeforeSave = application.CalculateBeforeSave;
					goto IL_003a;
					IL_003a:
					num2 = 7;
					application.Calculation = XlCalculation.xlCalculationManual;
					goto IL_0048;
					IL_0048:
					num2 = 8;
					application.CalculateBeforeSave = false;
					goto IL_0052;
					IL_0052:
					num2 = 9;
					application.EnableEvents = false;
					goto IL_005d;
					IL_005d:
					num2 = 10;
					application.ScreenUpdating = false;
					goto IL_0068;
					IL_0068:
					num2 = 11;
					application.DisplayStatusBar = true;
					goto IL_0073;
					IL_0073:
					num2 = 12;
					application.StatusBar = VH.A(185031);
					goto IL_0087;
					IL_0087:
					num2 = 13;
					enumerator = application.Workbooks.GetEnumerator();
					goto IL_00f6;
					IL_00f6:
					if (enumerator.MoveNext())
					{
						workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
						goto IL_00ac;
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
					goto IL_010b;
					IL_00de:
					num2 = 15;
					workbook.Save();
					goto IL_00e8;
					IL_010b:
					num2 = 18;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0123;
					IL_00f3:
					num2 = 17;
					goto IL_00f6;
					IL_0123:
					num2 = 19;
					application.StatusBar = "";
					goto IL_0132;
					IL_0132:
					num2 = 20;
					application.StatusBar = false;
					goto IL_0142;
					IL_0142:
					num2 = 21;
					application.EnableEvents = true;
					goto IL_014d;
					IL_014d:
					num2 = 22;
					application.ScreenUpdating = true;
					goto IL_0158;
					IL_0158:
					num2 = 23;
					application.Calculation = calculation;
					goto IL_0163;
					IL_0163:
					num2 = 24;
					application.CalculateBeforeSave = calculateBeforeSave;
					goto IL_016f;
					IL_016f:
					application = null;
					break;
					IL_00e8:
					num2 = 16;
					workbook.Saved = true;
					goto IL_00f3;
					end_IL_0000_3:
					break;
				}
				num2 = 26;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(185083));
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 521;
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
			switch (5)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void SaveAs()
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
			SaveAs(MH.A.Application.ActiveWorkbook);
			return;
		}
	}

	public static void SaveAs(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		XlCalculation calculation = default(XlCalculation);
		bool calculateBeforeSave = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 753:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0019;
						case 5:
							goto IL_002e;
						case 6:
							goto IL_0038;
						case 7:
							goto IL_0045;
						case 8:
							goto IL_0053;
						case 9:
							goto IL_005d;
						case 10:
							goto IL_0068;
						case 11:
							goto IL_0073;
						case 12:
							goto IL_0089;
						case 13:
							goto IL_0097;
						case 14:
							goto IL_00a0;
						case 15:
							goto IL_0212;
						case 16:
							goto IL_021c;
						case 17:
							goto IL_022b;
						case 18:
							goto IL_023b;
						case 19:
							goto IL_0246;
						case 20:
							goto IL_0251;
						case 21:
							goto IL_025d;
						case 22:
							goto IL_0260;
						case 23:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
						case 24:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0260:
					num2 = 22;
					workbook = null;
					break;
					IL_0007:
					num2 = 2;
					application = MH.A.Application;
					goto IL_0019;
					IL_0019:
					num2 = 3;
					if (application.Workbooks.Count == 0)
					{
						goto end_IL_0000_3;
					}
					goto IL_002e;
					IL_002e:
					num2 = 5;
					calculation = application.Calculation;
					goto IL_0038;
					IL_0038:
					num2 = 6;
					calculateBeforeSave = application.CalculateBeforeSave;
					goto IL_0045;
					IL_0045:
					num2 = 7;
					application.Calculation = XlCalculation.xlCalculationManual;
					goto IL_0053;
					IL_0053:
					num2 = 8;
					application.CalculateBeforeSave = false;
					goto IL_005d;
					IL_005d:
					num2 = 9;
					application.EnableEvents = false;
					goto IL_0068;
					IL_0068:
					num2 = 10;
					application.DisplayStatusBar = true;
					goto IL_0073;
					IL_0073:
					num2 = 11;
					application.StatusBar = VH.A(185031);
					goto IL_0089;
					IL_0089:
					num2 = 12;
					workbook = application.ActiveWorkbook;
					goto IL_0097;
					IL_0097:
					num2 = 13;
					wb.Activate();
					goto IL_00a0;
					IL_00a0:
					num2 = 14;
					application.Dialogs[XlBuiltInDialog.xlDialogSaveAs].Show(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0212;
					IL_0212:
					num2 = 15;
					workbook.Activate();
					goto IL_021c;
					IL_021c:
					num2 = 16;
					application.StatusBar = "";
					goto IL_022b;
					IL_022b:
					num2 = 17;
					application.StatusBar = false;
					goto IL_023b;
					IL_023b:
					num2 = 18;
					application.EnableEvents = true;
					goto IL_0246;
					IL_0246:
					num2 = 19;
					application.Calculation = calculation;
					goto IL_0251;
					IL_0251:
					num2 = 20;
					application.CalculateBeforeSave = calculateBeforeSave;
					goto IL_025d;
					IL_025d:
					application = null;
					goto IL_0260;
					end_IL_0000_2:
					break;
				}
				num2 = 23;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(185112));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 753;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void SaveUp()
	{
		if (Licensing.AllowRestrictedMode())
		{
			SaveUp(MH.A.Application.ActiveWorkbook);
		}
	}

	public static void SaveUp(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Workbooks.Count > 0)
			{
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
					if (wb.Path.Length == 0)
					{
						SaveAs(wb);
						break;
					}
					string text = clsFile.BaseName(wb.Name);
					string extension = Path.GetExtension(wb.Name);
					int num = clsFile.VersionNumber(wb.Name);
					if (num == 0)
					{
						Forms.WarningMessage(VH.A(185139));
						break;
					}
					num = checked(num + 1);
					string path = text + Conversions.ToString(num) + extension;
					string text2 = Path.Combine(wb.Path, path);
					if (!clsFile.IsPathUrl(text2))
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
						if (File.Exists(text2))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								Forms.WarningMessage(VH.A(185291));
								break;
							}
							break;
						}
					}
					if (clsFile.NewerVersions(text2).Count > 0)
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
						if (MessageBox.Show(VH.A(185392), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0139;
								}
								continue;
								end_IL_0139:
								break;
							}
							break;
						}
					}
					Microsoft.Office.Interop.Excel.Application application2 = application;
					XlCalculation calculation = application2.Calculation;
					bool calculateBeforeSave = application2.CalculateBeforeSave;
					application2.Calculation = XlCalculation.xlCalculationManual;
					application2.CalculateBeforeSave = false;
					application2.EnableEvents = false;
					application2.DisplayStatusBar = true;
					application2.StatusBar = VH.A(185031);
					try
					{
						wb.SaveAs(text2, wb.FileFormat, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						wb.Saved = true;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Forms.ErrorMessage(VH.A(185582) + ex2.Message);
						ProjectData.ClearProjectError();
					}
					application2.StatusBar = "";
					application2.StatusBar = false;
					application2.EnableEvents = true;
					application2.Calculation = calculation;
					application2.CalculateBeforeSave = calculateBeforeSave;
					application2 = null;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(185697));
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	public static void SaveDown()
	{
		if (!Licensing.AllowRestrictedMode())
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
			SaveDown(MH.A.Application.ActiveWorkbook);
			return;
		}
	}

	public static void SaveDown(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		Microsoft.Office.Interop.Excel.Application application = wb.Application;
		try
		{
			if (wb.Path.Length == 0)
			{
				SaveAs(wb);
			}
			else
			{
				string fullName = wb.FullName;
				if (!clsFile.IsPathUrl(fullName))
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
						string text = A(fullName);
						if (text.Length > 0)
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
							Microsoft.Office.Interop.Excel.Application application2 = application;
							XlCalculation calculation = application2.Calculation;
							bool calculateBeforeSave = application2.CalculateBeforeSave;
							application2.Calculation = XlCalculation.xlCalculationManual;
							application2.CalculateBeforeSave = false;
							application2.EnableEvents = false;
							application2.DisplayStatusBar = true;
							application2.StatusBar = VH.A(185031);
							try
							{
								wb.SaveCopyAs(text);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Forms.ErrorMessage(VH.A(185582) + ex2.Message);
								ProjectData.ClearProjectError();
							}
							application2.StatusBar = "";
							application2.StatusBar = false;
							application2.EnableEvents = true;
							application2.Calculation = calculation;
							application2.CalculateBeforeSave = calculateBeforeSave;
							application2 = null;
						}
						else
						{
							Forms.WarningMessage(VH.A(185724));
						}
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(185809));
						break;
					}
				}
				else
				{
					Forms.WarningMessage(VH.A(185840));
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static string A(string A)
	{
		XmlNode xmlNode = KH.A.SettingsXml.SelectSingleNode(VH.A(185967));
		int num = Conversions.ToInteger(xmlNode.SelectSingleNode(VH.A(185988)).InnerText);
		string innerText = xmlNode.SelectSingleNode(VH.A(186009)).InnerText;
		string innerText2;
		bool flag;
		if (num == 1)
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
			innerText2 = xmlNode.SelectSingleNode(VH.A(186028)).InnerText;
			flag = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(186049)).InnerText);
		}
		xmlNode = null;
		innerText = VH.A(41385);
		innerText2 = VH.A(186068);
		flag = true;
		string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(A);
		string extension = Path.GetExtension(fileNameWithoutExtension);
		string text = Path.GetDirectoryName(A) + Conversions.ToString(Path.DirectorySeparatorChar) + fileNameWithoutExtension + innerText;
		checked
		{
			if (0 == 0)
			{
				int num2 = 0;
				string fileName = Path.GetFileName(A);
				try
				{
					FileInfo[] filesOfType = clsFile.GetFilesOfType(A, extension);
					foreach (FileInfo fileInfo in filesOfType)
					{
						try
						{
							if (!fileInfo.Name.StartsWith(fileNameWithoutExtension))
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
								if (Operators.CompareString(fileInfo.Name, fileName, TextCompare: false) != 0)
								{
									int num3 = clsFile.VersionNumber(fileInfo.FullName);
									if (num3 > num2)
									{
										num2 = num3;
									}
								}
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
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_018f;
						}
						continue;
						end_IL_018f:
						break;
					}
				}
				catch (DirectoryNotFoundException ex3)
				{
					ProjectData.SetProjectError(ex3);
					DirectoryNotFoundException ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				catch (UnauthorizedAccessException ex5)
				{
					ProjectData.SetProjectError(ex5);
					UnauthorizedAccessException ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				catch (PathTooLongException ex7)
				{
					ProjectData.SetProjectError(ex7);
					PathTooLongException ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					clsReporting.LogException(ex10);
					ProjectData.ClearProjectError();
				}
				return text + (num2 + 1) + extension;
			}
			DateTime dateTime;
			if (flag)
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
				dateTime = DateTime.Now;
			}
			else
			{
				dateTime = DateTime.UtcNow;
			}
			text += dateTime.ToString(innerText2);
			if (!File.Exists(text + extension))
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return text + extension;
					}
				}
			}
			int num4 = 100;
			int num5 = num4;
			for (int j = 1; j <= num5; j++)
			{
				string text2 = text + VH.A(63506) + j;
				if (!File.Exists(text2 + extension))
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return text2 + extension;
						}
					}
				}
				if (j != num4)
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
					return "";
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				return "";
			}
		}
	}

	public static bool IsVersionInFileName(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		if (Operators.CompareString(wb.FullName, wb.Name, TextCompare: false) != 0)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return clsFile.FileNameVersionRegex().IsMatch(wb.Name);
				}
			}
		}
		return false;
	}
}
