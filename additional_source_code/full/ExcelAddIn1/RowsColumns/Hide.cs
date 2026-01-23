using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Hide
{
	[CompilerGenerated]
	internal sealed class HG
	{
		public Range A;

		public HG(HG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Range A(int A, int B)
		{
			return (Range)this.A.Rows[string.Format(VH.A(212114), A, B), RuntimeHelpers.GetObjectValue(Missing.Value)];
		}
	}

	[CompilerGenerated]
	internal sealed class IG
	{
		public Range A;

		public IG(IG A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Range A(int A, int B)
		{
			return (Range)this.A.Columns[string.Format(VH.A(212114), Hide.A(A), Hide.A(B)), RuntimeHelpers.GetObjectValue(Missing.Value)];
		}
	}

	public static void Rows()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Application application = default(Application);
		Application application2 = default(Application);
		XlCalculation calculation = default(XlCalculation);
		Range range = default(Range);
		IEnumerator enumerator = default(IEnumerator);
		Range range2 = default(Range);
		Range range3 = default(Range);
		IEnumerator enumerator2 = default(IEnumerator);
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
				case 534:
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
							goto IL_0018;
						case 4:
							goto IL_001d;
						case 5:
							goto IL_0027;
						case 6:
							goto IL_0034;
						case 7:
							goto IL_0042;
						case 8:
							goto IL_006f;
						case 9:
							goto IL_0072;
						case 10:
							goto IL_0092;
						case 11:
							goto IL_00ac;
						case 13:
							goto IL_00b6;
						case 12:
						case 14:
							goto IL_00d7;
						case 15:
							goto IL_00ef;
						case 16:
							goto IL_010d;
						case 17:
							goto IL_0133;
						case 18:
							goto IL_0143;
						case 19:
							goto IL_0151;
						case 20:
							goto IL_0169;
						case 21:
							goto IL_0174;
						case 22:
							goto IL_017e;
						case 23:
							goto IL_0192;
						case 24:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 25:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0174:
					num2 = 21;
					application.ScreenUpdating = true;
					goto IL_017e;
					IL_0007:
					num2 = 2;
					application = MH.A.Application;
					goto IL_0018;
					IL_0018:
					num2 = 3;
					application2 = application;
					goto IL_001d;
					IL_001d:
					num2 = 4;
					application2.ScreenUpdating = false;
					goto IL_0027;
					IL_0027:
					num2 = 5;
					calculation = application2.Calculation;
					goto IL_0034;
					IL_0034:
					num2 = 6;
					application2.Calculation = XlCalculation.xlCalculationManual;
					goto IL_0042;
					IL_0042:
					num2 = 7;
					range = (Range)NewLateBinding.LateGet(application2.Selection, null, VH.A(152043), new object[0], null, null, null);
					goto IL_006f;
					IL_006f:
					application2 = null;
					goto IL_0072;
					IL_0072:
					num2 = 9;
					enumerator = range.GetEnumerator();
					goto IL_00b9;
					IL_00b9:
					if (enumerator.MoveNext())
					{
						range2 = (Range)enumerator.Current;
						goto IL_0092;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_00d7;
					IL_017e:
					num2 = 22;
					Core.LogActivity(VH.A(171514));
					goto IL_0192;
					IL_0133:
					num2 = 17;
					range3.ShowDetail = false;
					goto IL_0143;
					IL_010d:
					num2 = 16;
					if (Operators.ConditionalCompareObjectGreater(range3.OutlineLevel, 1, TextCompare: false))
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
						goto IL_0133;
					}
					goto IL_0143;
					IL_0192:
					num2 = 23;
					range = null;
					break;
					IL_0092:
					num2 = 10;
					if (Operators.ConditionalCompareObjectEqual(range2.OutlineLevel, 1, TextCompare: false))
					{
						goto IL_00ac;
					}
					goto IL_00b6;
					IL_00ac:
					num2 = 11;
					A();
					goto IL_00d7;
					IL_00d7:
					num2 = 14;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00ef;
					IL_00b6:
					num2 = 13;
					goto IL_00b9;
					IL_00ef:
					num2 = 15;
					enumerator2 = range.GetEnumerator();
					goto IL_0146;
					IL_0146:
					if (enumerator2.MoveNext())
					{
						range3 = (Range)enumerator2.Current;
						goto IL_010d;
					}
					goto IL_0151;
					IL_0151:
					num2 = 19;
					if (enumerator2 is IDisposable)
					{
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_0169;
					IL_0143:
					num2 = 18;
					goto IL_0146;
					IL_0169:
					num2 = 20;
					application.Calculation = calculation;
					goto IL_0174;
					end_IL_0000_2:
					break;
				}
				num2 = 24;
				application = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 534;
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
			switch (2)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void Columns()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Application application = default(Application);
		Application application2 = default(Application);
		XlCalculation calculation = default(XlCalculation);
		Range range = default(Range);
		IEnumerator enumerator = default(IEnumerator);
		Range range2 = default(Range);
		Range range3 = default(Range);
		IEnumerator enumerator2 = default(IEnumerator);
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
				case 546:
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
							goto IL_0018;
						case 4:
							goto IL_001d;
						case 5:
							goto IL_0027;
						case 6:
							goto IL_0034;
						case 7:
							goto IL_0042;
						case 8:
							goto IL_006f;
						case 9:
							goto IL_0072;
						case 10:
							goto IL_0092;
						case 11:
							goto IL_00c1;
						case 13:
							goto IL_00cb;
						case 12:
						case 14:
							goto IL_00d9;
						case 15:
							goto IL_00fb;
						case 16:
							goto IL_011b;
						case 17:
							goto IL_0137;
						case 18:
							goto IL_0147;
						case 19:
							goto IL_0153;
						case 20:
							goto IL_0175;
						case 21:
							goto IL_0180;
						case 22:
							goto IL_018a;
						case 23:
							goto IL_019e;
						case 24:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 25:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00cb:
					num2 = 13;
					goto IL_00ce;
					IL_0007:
					num2 = 2;
					application = MH.A.Application;
					goto IL_0018;
					IL_0018:
					num2 = 3;
					application2 = application;
					goto IL_001d;
					IL_001d:
					num2 = 4;
					application2.ScreenUpdating = false;
					goto IL_0027;
					IL_0027:
					num2 = 5;
					calculation = application2.Calculation;
					goto IL_0034;
					IL_0034:
					num2 = 6;
					application2.Calculation = XlCalculation.xlCalculationManual;
					goto IL_0042;
					IL_0042:
					num2 = 7;
					range = (Range)NewLateBinding.LateGet(application2.Selection, null, VH.A(152073), new object[0], null, null, null);
					goto IL_006f;
					IL_006f:
					application2 = null;
					goto IL_0072;
					IL_0072:
					num2 = 9;
					enumerator = range.GetEnumerator();
					goto IL_00ce;
					IL_00ce:
					if (enumerator.MoveNext())
					{
						range2 = (Range)enumerator.Current;
						goto IL_0092;
					}
					goto IL_00d9;
					IL_0175:
					num2 = 20;
					application.Calculation = calculation;
					goto IL_0180;
					IL_0092:
					num2 = 10;
					if (Operators.ConditionalCompareObjectEqual(range2.OutlineLevel, 1, TextCompare: false))
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
						goto IL_00c1;
					}
					goto IL_00cb;
					IL_0180:
					num2 = 21;
					application.ScreenUpdating = true;
					goto IL_018a;
					IL_018a:
					num2 = 22;
					Core.LogActivity(VH.A(171533));
					goto IL_019e;
					IL_019e:
					num2 = 23;
					range = null;
					break;
					IL_00c1:
					num2 = 11;
					GroupColumns();
					goto IL_00d9;
					IL_00d9:
					num2 = 14;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00fb;
					IL_0137:
					num2 = 17;
					range3.ShowDetail = false;
					goto IL_0147;
					IL_011b:
					num2 = 16;
					if (Operators.ConditionalCompareObjectGreater(range3.OutlineLevel, 1, TextCompare: false))
					{
						goto IL_0137;
					}
					goto IL_0147;
					IL_00fb:
					num2 = 15;
					enumerator2 = range.GetEnumerator();
					goto IL_014a;
					IL_014a:
					if (enumerator2.MoveNext())
					{
						range3 = (Range)enumerator2.Current;
						goto IL_011b;
					}
					goto IL_0153;
					IL_0153:
					num2 = 19;
					if (enumerator2 is IDisposable)
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
						(enumerator2 as IDisposable).Dispose();
					}
					goto IL_0175;
					IL_0147:
					num2 = 18;
					goto IL_014a;
					end_IL_0000_2:
					break;
				}
				num2 = 24;
				application = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 546;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static void A()
	{
		Range range = default(Range);
		try
		{
			range = (Range)MH.A.Application.Selection;
			if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, range.Worksheet.Rows.CountLarge, TextCompare: false))
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
					B();
					break;
				}
			}
			else
			{
				range.Rows.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		JH.A((object)range);
	}

	public static void GroupColumns()
	{
		Range range = default(Range);
		try
		{
			range = (Range)MH.A.Application.Selection;
			if (Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, range.Worksheet.Columns.CountLarge, TextCompare: false))
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
					C();
					break;
				}
			}
			else
			{
				range.Columns.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		JH.A((object)range);
	}

	private static void B()
	{
		Forms.WarningMessage(VH.A(171379));
	}

	private static void C()
	{
		Forms.WarningMessage(VH.A(171457));
	}

	public static void ProperHide()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					application = null;
					return;
				}
			}
		}
		application.ScreenUpdating = false;
		object objectValue = default(object);
		try
		{
			if (application.Windows.Count > 0)
			{
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (application.ActiveWindow.SelectedSheets.Count > 1)
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
						if (!Core.ConfirmMultipleSheets())
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0089;
								}
								continue;
								end_IL_0089:
								break;
							}
							break;
						}
					}
					objectValue = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
					enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
							if (!(objectValue2 is Worksheet))
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
							Worksheet obj = (Worksheet)objectValue2;
							obj.Activate();
							A(obj);
							B(obj);
							D(obj);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0108;
							}
							continue;
							end_IL_0108:
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
					Core.LogActivity(VH.A(171558));
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		if (objectValue != null)
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
			NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
		}
		application.ScreenUpdating = true;
		application = null;
		objectValue = null;
	}

	private static void A(Worksheet A)
	{
		Range range = null;
		Range range3 = default(Range);
		Range range4 = default(Range);
		try
		{
			Range range2 = Hide.A(A);
			if (range2 == null)
			{
				return;
			}
			range3 = (Range)A.Columns[range2.Column, RuntimeHelpers.GetObjectValue(Missing.Value)];
			range4 = range3.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			range3.EntireRow.Hidden = false;
			range4.EntireRow.Hidden = true;
			try
			{
				range = range3.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range3.EntireRow.Hidden = false;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range range2 = null;
		}
		if (range != null)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				HG hG = default(HG);
				while (enumerator.MoveNext())
				{
					Range range5 = (Range)enumerator.Current;
					hG = new HG(hG);
					try
					{
						hG.A = range5.EntireRow;
						Hide.A(hG.A, 1, hG.A.Count);
					}
					finally
					{
						hG.A = null;
						range5 = null;
					}
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
		JH.A((object)range3);
		JH.A((object)range4);
		JH.A((object)range);
	}

	private static void B(Worksheet A)
	{
		Range range = null;
		Range range3 = default(Range);
		Range range4 = default(Range);
		try
		{
			Range range2 = Hide.A(A);
			if (range2 == null)
			{
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
					return;
				}
			}
			range3 = (Range)A.Rows[range2.Row, RuntimeHelpers.GetObjectValue(Missing.Value)];
			range4 = range3.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			range3.EntireColumn.Hidden = false;
			range4.EntireColumn.Hidden = true;
			try
			{
				range = range3.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range3.EntireColumn.Hidden = false;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range range2 = null;
		}
		if (range != null)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				IG iG = default(IG);
				while (enumerator.MoveNext())
				{
					Range range5 = (Range)enumerator.Current;
					iG = new IG(iG);
					try
					{
						iG.A = range5.EntireColumn;
						Hide.A(iG.A, 1, iG.A.Count);
					}
					finally
					{
						iG.A = null;
						range5 = null;
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0157;
					}
					continue;
					end_IL_0157:
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
		JH.A((object)range3);
		JH.A((object)range4);
		JH.A((object)range);
	}

	private static Range A(Worksheet A)
	{
		try
		{
			Range range = Ranges.FirstVisibleCell(A);
			if (range != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return range;
					}
				}
			}
			C(A);
			return Ranges.FirstVisibleCell(A);
		}
		finally
		{
			Range range = null;
		}
	}

	private static void C(Worksheet A)
	{
		try
		{
			Range usedRange = A.UsedRange;
			Range cells = A.Cells;
			cells.EntireRow.Hidden = false;
			cells.EntireColumn.Hidden = false;
			usedRange.EntireRow.Hidden = true;
			usedRange.EntireColumn.Hidden = true;
		}
		finally
		{
		}
	}

	private static string A(int A)
	{
		return Ranges.ColNumToLetter(A);
	}

	private static void A(Func<int, int, Range> A, int B, int C)
	{
		if (B > C)
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
			if (Hide.A(A, B, C, D: true))
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
				int num = B;
				int num2 = C;
				int num3;
				while (true)
				{
					num3 = checked(num + num2) / 2;
					if (num == num3)
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
					if (Hide.A(A, B, num3, D: false))
					{
						num = num3;
					}
					else
					{
						num2 = num3;
					}
				}
				Hide.A(A, B, num3);
				Hide.A(A, checked(num3 + 1), C);
				return;
			}
		}
	}

	private static bool A(Func<int, int, Range> A, int B, int C, bool D)
	{
		try
		{
			Range range = A(B, C);
			object objectValue = RuntimeHelpers.GetObjectValue(range.OutlineLevel);
			int? num = null;
			if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(objectValue)))
			{
				num = Conversions.ToInteger(objectValue);
			}
			if (D)
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
				if (object.Equals(num, 1))
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
					range.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
			int result;
			if (B != C)
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
				result = (num.HasValue ? 1 : 0);
			}
			else
			{
				result = 1;
			}
			return (byte)result != 0;
		}
		finally
		{
			Range range = null;
		}
	}

	private static void D(Worksheet A)
	{
		A.Outline.ShowLevels(1, 1);
	}
}
