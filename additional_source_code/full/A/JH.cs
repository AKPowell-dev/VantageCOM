using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using System.Windows.Threading;
using ExcelAddIn1.UndoRedo;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class JH
{
	public static void A(object A)
	{
		try
		{
			int num = 0;
			do
			{
				num = Marshal.ReleaseComObject(RuntimeHelpers.GetObjectValue(A));
			}
			while (num > 0);
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
				A = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A = null;
			ProjectData.ClearProjectError();
		}
		finally
		{
			System.GC.Collect();
			System.GC.WaitForPendingFinalizers();
			System.GC.Collect();
			System.GC.WaitForPendingFinalizers();
		}
	}

	public static bool A(Range A)
	{
		if (KH.A.UndoEnabled)
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
					return Core.IndexSelection(A);
				}
			}
		}
		bool result = default(bool);
		return result;
	}

	public static void A(Range A, string B = "")
	{
		if (!KH.A.UndoEnabled)
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
			Core.SaveToUndoStack(A, B);
			return;
		}
	}

	public static Range A(Range A = null)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Application application = default(Application);
		bool flag = default(bool);
		bool flag2 = default(bool);
		Range result = default(Range);
		Range range = default(Range);
		Range range2 = default(Range);
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
				case 1104:
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
						case 4:
							goto IL_0031;
						case 5:
							goto IL_004b;
						case 7:
							goto IL_005f;
						case 6:
						case 8:
							goto IL_0068;
						case 9:
							goto IL_006d;
						case 10:
							goto IL_0110;
						case 11:
							goto IL_0116;
						case 12:
							goto IL_01bb;
						case 13:
							goto IL_01c1;
						case 14:
							goto IL_01c4;
						case 15:
							goto IL_01e3;
						case 17:
							goto IL_01ed;
						case 18:
							goto IL_0205;
						case 19:
							goto IL_0383;
						case 20:
							goto IL_0394;
						case 22:
							goto IL_039b;
						case 24:
							goto end_IL_0000_2;
						case 16:
						case 21:
						case 23:
						case 25:
							application = null;
							goto end_IL_0000_3;
						default:
							goto end_IL_0000;
						case 26:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_01c4:
					num2 = 14;
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
						if (!flag2)
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
							goto IL_01e3;
						}
					}
					goto IL_01ed;
					IL_0007:
					num2 = 2;
					application = MH.A.Application;
					goto IL_0019;
					IL_0019:
					num2 = 3;
					if (A == null)
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
						goto IL_0031;
					}
					goto IL_0068;
					IL_0394:
					num2 = 20;
					result = A;
					goto end_IL_0000_3;
					IL_0205:
					num2 = 18;
					range = application.Intersect(A, (Range)NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(82416), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0383;
					IL_0383:
					num2 = 19;
					if (range == null)
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
						goto IL_0394;
					}
					goto IL_039b;
					IL_0031:
					num2 = 4;
					if (application.Selection is Range)
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
						goto IL_004b;
					}
					goto IL_005f;
					IL_01e3:
					num2 = 15;
					result = A;
					goto end_IL_0000_3;
					IL_004b:
					num2 = 5;
					A = (Range)application.Selection;
					goto IL_0068;
					IL_005f:
					num2 = 7;
					result = null;
					goto end_IL_0000_3;
					IL_0068:
					num2 = 8;
					range2 = A;
					goto IL_006d;
					IL_006d:
					num2 = 9;
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
						goto IL_0110;
					}
					goto IL_0116;
					IL_01ed:
					num2 = 17;
					if (!flag)
					{
						if (!flag2)
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
							break;
						}
					}
					goto IL_0205;
					IL_0110:
					num2 = 10;
					flag = true;
					goto IL_0116;
					IL_0116:
					num2 = 11;
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51255), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
						goto IL_01bb;
					}
					goto IL_01c1;
					IL_039b:
					num2 = 22;
					result = range;
					goto end_IL_0000_3;
					IL_01bb:
					num2 = 12;
					flag2 = true;
					goto IL_01c1;
					IL_01c1:
					range2 = null;
					goto IL_01c4;
					end_IL_0000_2:
					break;
				}
				num2 = 24;
				result = (Range)NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(82416), new object[0], null, null, null);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1104;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static Range A(Range A, Application B = null)
	{
		if (!JH.B(A))
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
					return A;
				}
			}
		}
		return JH.B(A, B);
	}

	private static bool B(Range A)
	{
		if (A == null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return false;
				}
			}
		}
		try
		{
			string right = A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (Operators.CompareString(((Range)A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).EntireRow.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
			{
				return true;
			}
			if (Operators.CompareString(((Range)A.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).EntireColumn.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
		return false;
	}

	private static Range B(Range A, Application B = null)
	{
		Range result;
		try
		{
			Range usedRange = A.Worksheet.UsedRange;
			if (usedRange == null)
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
					result = null;
					break;
				}
			}
			else
			{
				if (B == null)
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
					B = A.Application;
				}
				result = B.Intersect(A, usedRange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = A;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range usedRange = null;
			B = null;
		}
		return result;
	}

	public static int A(this Range A)
	{
		try
		{
			Range rows = A.Rows;
			return checked(A.Row - 1 + rows.Count);
		}
		finally
		{
			Range rows = null;
		}
	}

	public static int B(this Range A)
	{
		try
		{
			Range columns = A.Columns;
			return checked(A.Column - 1 + columns.Count);
		}
		finally
		{
			Range columns = null;
		}
	}

	public static Range A(this Worksheet A, int B, int C, int D, int E)
	{
		try
		{
			Range range = (Range)A.Cells[B, C];
			if (B == D && C == E)
			{
				return range;
			}
			Range cell = (Range)A.Cells[D, E];
			return ((_Worksheet)A).get_Range((object)range, (object)cell);
		}
		finally
		{
			Range cell = null;
			Range range = null;
		}
	}

	public static void A(this Range A, ref int B, ref int C, ref int D, ref int E)
	{
		B = A.Row;
		C = A.Column;
		D = JH.A(A);
		E = JH.B(A);
	}

	internal static Range A(Range A, bool B, int C)
	{
		checked
		{
			try
			{
				int num = (B ? A.Row : A.Column) + C;
				Range range;
				if (!B)
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
					range = A.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)C);
				}
				else
				{
					range = A.get_Offset((object)C, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Range range2 = range;
				if (C == 0)
				{
					return range2;
				}
				if ((B ? range2.Row : range2.Column) == num)
				{
					return range2;
				}
				try
				{
					Worksheet worksheet = A.Worksheet;
					_ = worksheet.Cells;
					Range cells = worksheet.Cells;
					int row = A.Row;
					int num2;
					if (!B)
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
						num2 = 0;
					}
					else
					{
						num2 = C;
					}
					object rowIndex = row + num2;
					int column = A.Column;
					int num3;
					if (!B)
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
						num3 = C;
					}
					else
					{
						num3 = 0;
					}
					return (Range)cells[rowIndex, column + num3];
				}
				finally
				{
				}
			}
			finally
			{
				Range range2 = null;
			}
		}
	}

	internal static Range A(this Range A, Range B, [Optional][DefaultParameterValue(0)] out int C, [Optional][DefaultParameterValue(0)] out int D)
	{
		try
		{
			Range rows = B.Rows;
			Range columns = B.Columns;
			C = rows.Count;
			D = columns.Count;
			return A.get_Resize((object)C, (object)D);
		}
		finally
		{
			Range columns = null;
			Range rows = null;
		}
	}

	public static string B(string A)
	{
		if (!A.Contains(VH.A(39830)))
		{
			return string.Format(VH.A(205055), A);
		}
		if (!A.Contains(VH.A(39851)))
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return string.Format(VH.A(205066), A);
				}
			}
		}
		StringBuilder stringBuilder = new StringBuilder(VH.A(205077));
		int num = 0;
		checked
		{
			for (int num2 = A.IndexOf('"'); num2 != -1; num2 = A.IndexOf('"', num))
			{
				if (num != 0)
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
					stringBuilder.AppendFormat(VH.A(205092), A.Substring(num));
				}
				stringBuilder.AppendFormat(VH.A(205107), A.Substring(num, num2 - num));
				num += num2;
			}
			stringBuilder.Append(VH.A(39904));
			return stringBuilder.ToString();
		}
	}

	[SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
	internal static void A()
	{
		DispatcherFrame dispatcherFrame = new DispatcherFrame();
		Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(C), dispatcherFrame);
		Dispatcher.PushFrame(dispatcherFrame);
	}

	private static object C(object A)
	{
		((DispatcherFrame)A).Continue = false;
		return null;
	}

	internal static void A(this BackgroundWorker A, int? B, int? C = 200)
	{
		DateTime utcNow = DateTime.UtcNow;
		while (true)
		{
			if (A.IsBusy)
			{
				if (B.HasValue)
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
					if ((DateTime.UtcNow - utcNow).TotalMilliseconds > (double)B.Value)
					{
						break;
					}
				}
				JH.A();
				if (!A.IsBusy)
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
				if (!C.HasValue)
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
				Thread.Sleep(C.Value);
				continue;
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
	}

	public static bool A(object A)
	{
		if (A is int)
		{
			return true;
		}
		if (A is long)
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
					return true;
				}
			}
		}
		if (A is double)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is float)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is decimal)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is byte)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is short)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is ushort)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is uint)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if (A is ulong)
		{
			return true;
		}
		if (A is sbyte)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}
}
