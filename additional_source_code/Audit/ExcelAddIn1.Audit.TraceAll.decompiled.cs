using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class TraceAll
{
	private static readonly int m_A = 1000;

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static bool m_B;

	public static bool ShowingPrecedents
	{
		[CompilerGenerated]
		get
		{
			return TraceAll.m_A;
		}
		[CompilerGenerated]
		set
		{
			TraceAll.m_A = value;
		}
	} = false;

	public static bool ShowingDependents
	{
		[CompilerGenerated]
		get
		{
			return TraceAll.m_B;
		}
		[CompilerGenerated]
		set
		{
			TraceAll.m_B = value;
		}
	} = false;

	public static void Precedents()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
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
					goto IL_001c;
				case 183:
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
							goto IL_001c;
						case 4:
							goto IL_0023;
						case 5:
							goto IL_0038;
						case 6:
							goto IL_003f;
						case 8:
							goto IL_0049;
						case 7:
						case 9:
							goto IL_0057;
						case 10:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 11:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0057:
					num2 = 9;
					ShowingDependents = false;
					break;
					IL_0038:
					num2 = 5;
					Arrows.ClearArrowsOnActiveSheet();
					goto IL_003f;
					IL_003f:
					num2 = 6;
					ShowingPrecedents = false;
					goto IL_0057;
					IL_0049:
					num2 = 8;
					ShowingPrecedents = A();
					goto IL_0057;
					IL_001c:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0023;
					IL_0023:
					num2 = 4;
					if (ShowingPrecedents)
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
						goto IL_0038;
					}
					goto IL_0049;
					end_IL_0000_3:
					break;
				}
				num2 = 10;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(39907));
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 183;
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

	public static void Dependents()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
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
					goto IL_001e;
				case 173:
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
							goto IL_001e;
						case 4:
							goto IL_0025;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_0037;
						case 8:
							goto IL_0041;
						case 7:
						case 9:
							goto IL_004d;
						case 10:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 11:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0030:
					num2 = 5;
					Arrows.ClearArrowsOnActiveSheet();
					goto IL_0037;
					IL_0037:
					num2 = 6;
					ShowingDependents = false;
					goto IL_004d;
					IL_0041:
					num2 = 8;
					ShowingDependents = B();
					goto IL_004d;
					IL_004d:
					num2 = 9;
					ShowingPrecedents = false;
					break;
					IL_001e:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0025;
					IL_0025:
					num2 = 4;
					if (ShowingDependents)
					{
						goto IL_0030;
					}
					goto IL_0041;
					end_IL_0000_3:
					break;
				}
				num2 = 10;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(39946));
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 173;
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
			switch (3)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static bool A()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Range range = null;
		bool flag = false;
		Range range2 = default(Range);
		Range activeCell = default(Range);
		try
		{
			if (!(application.Selection is Range))
			{
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
					application = null;
					break;
				}
			}
			else
			{
				range2 = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectGreater(range2.Cells.CountLarge, 1, TextCompare: false))
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
					try
					{
						range = range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				else if (Conversions.ToBoolean(range2.HasFormula))
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
					range = range2;
				}
				if (range == null)
				{
					goto IL_0247;
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
				if (!Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, TraceAll.m_A, TextCompare: false))
				{
					goto IL_0103;
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
				if (C())
				{
					goto IL_0103;
				}
			}
			goto end_IL_0012;
			IL_0103:
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			try
			{
				Arrows.DisplayObjects(application.ActiveWorkbook);
				activeCell = application.ActiveCell;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = range.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range3 = (Range)enumerator.Current;
						try
						{
							range3.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
							if (flag)
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
								flag = A(range3, B: true, 1);
								if (flag)
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
									flag = A(range3, B: true, 2);
									break;
								}
								break;
							}
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
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_01a7;
						}
						continue;
						end_IL_01a7:
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
				Range range4 = range2;
				NewLateBinding.LateCall(range4.Worksheet.Parent, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
				range4.Worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
				range4.Select();
				_ = null;
				activeCell.Activate();
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			goto IL_0247;
			IL_0247:
			if (!flag)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					Forms.InfoMessage(VH.A(40002));
					break;
				}
			}
			end_IL_0012:;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		application = null;
		JH.A((object)range2);
		JH.A((object)activeCell);
		JH.A((object)range);
		return flag;
	}

	private static bool B()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool flag = false;
		Range range = default(Range);
		Range range2 = default(Range);
		Range activeCell = default(Range);
		try
		{
			if (!(application.Selection is Range))
			{
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
					application = null;
					break;
				}
			}
			else
			{
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				Arrows.DisplayObjects(application.ActiveWorkbook);
				range = (Range)application.Selection;
				range2 = JH.A((Range)null);
				if (range2 == null)
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
					range2 = range;
				}
				if (!Operators.ConditionalCompareObjectGreater(range2.Cells.CountLarge, TraceAll.m_A, TextCompare: false))
				{
					goto IL_00bf;
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
				if (C())
				{
					goto IL_00bf;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00b0;
					}
					continue;
					end_IL_00b0:
					break;
				}
			}
			goto end_IL_0012;
			IL_00bf:
			activeCell = application.ActiveCell;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range2.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range3 = (Range)enumerator.Current;
					try
					{
						range3.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
						if (flag)
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
							flag = A(range3, B: false, 1);
							if (flag)
							{
								break;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								flag = A(range3, B: false, 2);
								break;
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
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0148;
					}
					continue;
					end_IL_0148:
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
			Range range4 = range;
			NewLateBinding.LateCall(range4.Worksheet.Parent, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			range4.Worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			range4.Select();
			_ = null;
			activeCell.Activate();
			if (!flag)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					Forms.InfoMessage(VH.A(40093));
					break;
				}
			}
			end_IL_0012:;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		finally
		{
			application.ScreenUpdating = true;
			application.EnableEvents = true;
		}
		application = null;
		JH.A((object)range2);
		JH.A((object)range);
		JH.A((object)activeCell);
		return flag;
	}

	private static bool C()
	{
		return MessageBox.Show(VH.A(40184), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.OK;
	}

	private static bool A(Range A, bool B, int C)
	{
		bool result = false;
		try
		{
			A.NavigateArrow(B, C, 1);
			if (Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), A.Application.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
			{
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
}
