using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class RangeHelpers
{
	public interface IActionStatusUpdater
	{
		void ActionStarted(string actionDesc, long numItems = 1L);

		bool ItemCancelled(string itemDesc = "");

		void ActionEnded();
	}

	internal static Range A(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static Range A(Worksheet A)
	{
		try
		{
			return RangeHelpers.A(A.UsedRange);
		}
		finally
		{
		}
	}

	internal static Range B(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlNumbers);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static Range CellsWithNumbers(Range rng)
	{
		Range range = null;
		Range result = null;
		Range range2;
		if (Operators.ConditionalCompareObjectGreater(rng.Cells.CountLarge, 1, TextCompare: false))
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
			try
			{
				range = rng.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlNumbers);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range2 = RangeHelpers.B(rng);
			if (range2 != null)
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
					result = rng.Application.Union(range2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_01da;
				}
			}
			if (range != null)
			{
				result = range;
			}
			else if (range2 != null)
			{
				result = range2;
			}
			goto IL_01da;
		}
		try
		{
			if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(rng.Value2)))
			{
				result = rng;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		goto IL_020a;
		IL_01da:
		range = null;
		range2 = null;
		goto IL_020a;
		IL_020a:
		return result;
	}

	internal static Range C(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeFormulas, XlSpecialCellsValue.xlErrors);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static Range D(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeAllFormatConditions, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static List<Range> A(Worksheet A, IActionStatusUpdater B = null, string C = "")
	{
		List<Range> list = new List<Range>();
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		Range range = null;
		Range range2 = null;
		try
		{
			range = A.Cells.SpecialCells(XlCellType.xlCellTypeAllFormatConditions, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range != null)
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
			if (B != null)
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
				B.ActionStarted(C, Conversions.ToLong(range.CountLarge));
			}
			{
				IEnumerator enumerator = range.GetEnumerator();
				try
				{
					while (true)
					{
						if (enumerator.MoveNext())
						{
							Range range3 = (Range)enumerator.Current;
							if (B == null)
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
							}
							else if (B.ItemCancelled())
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_00bc;
									}
									continue;
									end_IL_00bc:
									break;
								}
								break;
							}
							if (range2 != null)
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
								if (application.Intersect(range2, range3, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
							}
							Range range4 = range3.SpecialCells(XlCellType.xlCellTypeSameFormatConditions, RuntimeHelpers.GetObjectValue(Missing.Value));
							if (range2 != null)
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
								if (application.Intersect(range2, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
								{
									list.Add(range4);
									range2 = application.Union(range2, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
							}
							else
							{
								range2 = range4;
								list.Add(range4);
							}
							range4 = null;
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_052b;
							}
							continue;
							end_IL_052b:
							break;
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
			if (B != null)
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
				B.ActionEnded();
			}
			range = null;
		}
		application = null;
		range2 = null;
		return list;
	}

	internal static Range E(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeAllValidation, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static Range F(Range A)
	{
		Range result;
		try
		{
			result = A.SpecialCells(XlCellType.xlCellTypeComments, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static Range B(Worksheet A)
	{
		Range result;
		try
		{
			result = A.Cells.SpecialCells(XlCellType.xlCellTypeComments, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static List<Range> A(Range A)
	{
		List<Range> list = new List<Range>();
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		List<Range> result;
		try
		{
			object obj = NewLateBinding.LateGet(A.Worksheet, null, VH.A(8668), new object[0], null, null, null);
			int num = Conversions.ToInteger(NewLateBinding.LateGet(obj, null, VH.A(52690), new object[0], null, null, null));
			for (int i = 1; i <= num; i = checked(i + 1))
			{
				Microsoft.Office.Interop.Excel.Application application2 = application;
				object[] array;
				bool[] array2;
				object instance = NewLateBinding.LateGet(obj, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
				if (array2[0])
				{
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				if (application2.Intersect(A, (Range)NewLateBinding.LateGet(instance, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
				{
					continue;
				}
				object instance2 = obj;
				string memberName = VH.A(140662);
				object[] obj2 = new object[1] { i };
				array = obj2;
				bool[] obj3 = new bool[1] { true };
				array2 = obj3;
				object instance3 = NewLateBinding.LateGet(instance2, null, memberName, obj2, null, null, obj3);
				if (array2[0])
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
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				list.Add((Range)NewLateBinding.LateGet(instance3, null, VH.A(8701), new object[0], null, null, null));
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				obj = null;
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
			goto IL_030b;
		}
		application = null;
		result = ((!list.Any()) ? null : list);
		goto IL_030b;
		IL_030b:
		return result;
	}

	internal static Range G(Range A)
	{
		Range A2 = null;
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		try
		{
			object instance = NewLateBinding.LateGet(A.Worksheet, null, VH.A(8668), new object[0], null, null, null);
			int num = Conversions.ToInteger(NewLateBinding.LateGet(instance, null, VH.A(52690), new object[0], null, null, null));
			for (int i = 1; i <= num; i = checked(i + 1))
			{
				object[] array;
				bool[] array2;
				object instance2 = NewLateBinding.LateGet(instance, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
				if (array2[0])
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
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				Range range = (Range)NewLateBinding.LateGet(instance2, null, VH.A(8701), new object[0], null, null, null);
				if (application.Intersect(A, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
					RangeHelpers.A(ref A2, range);
				}
				range = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				instance = null;
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application = null;
		return A2;
	}

	internal static Range C(Worksheet A)
	{
		Range A2 = null;
		try
		{
			object instance = NewLateBinding.LateGet(A, null, VH.A(8668), new object[0], null, null, null);
			int num = Conversions.ToInteger(NewLateBinding.LateGet(instance, null, VH.A(52690), new object[0], null, null, null));
			for (int i = 1; i <= num; i = checked(i + 1))
			{
				object[] array;
				bool[] array2;
				object instance2 = NewLateBinding.LateGet(instance, null, VH.A(140662), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
				if (array2[0])
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
					i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
				}
				Range b = (Range)NewLateBinding.LateGet(instance2, null, VH.A(8701), new object[0], null, null, null);
				RangeHelpers.A(ref A2, b);
				b = null;
			}
			instance = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return A2;
	}

	internal static Range H(Range A)
	{
		return Ranges.NonBlankCells(A);
	}

	internal static Range A(Range A, IActionStatusUpdater B = null, string C = "", long? D = null)
	{
		Range A2 = null;
		Range range = RangeHelpers.B(A);
		IEnumerator enumerator = default(IEnumerator);
		if (range != null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					long num = D ?? Conversions.ToLong(range.CountLarge);
					if (B != null)
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
						B.ActionStarted(C, num);
					}
					long num2 = 0L;
					try
					{
						try
						{
							enumerator = range.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Range range2 = (Range)enumerator.Current;
								num2 = checked(num2 + 1);
								if (num2 > num)
								{
									break;
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
								if (B == null)
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
								}
								else if (B.ItemCancelled())
								{
									break;
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
								if (!Core.HasDependents(range2))
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
									RangeHelpers.A(ref A2, range2);
								}
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										(enumerator as IDisposable).Dispose();
										goto end_IL_00ea;
									}
									continue;
									end_IL_00ea:
									break;
								}
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
					finally
					{
						if (B != null)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									B.ActionEnded();
									goto end_IL_0123;
								}
								continue;
								end_IL_0123:
								break;
							}
						}
					}
					return A2;
				}
				}
			}
		}
		return null;
	}

	internal static List<Range> B(Range A)
	{
		List<Range> list = new List<Range>();
		List<string> list2 = new List<string>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				try
				{
					if (!Conversions.ToBoolean(range.HasArray))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						Range currentArray = range.CurrentArray;
						if (!list2.Contains(currentArray.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))))
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
							list2.Add(currentArray.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
							list.Add(currentArray);
						}
						currentArray = null;
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
		list2 = null;
		return list;
	}

	internal static List<Range> A(Range A, IActionStatusUpdater B = null, string C = "")
	{
		List<Range> list = new List<Range>();
		if (RangeHelpers.A(A))
		{
			List<string> list2 = new List<string>();
			B?.ActionStarted(C, Conversions.ToLong(A.CountLarge));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						if (B != null && B.ItemCancelled())
						{
							break;
						}
						try
						{
							if (!Conversions.ToBoolean(range.MergeCells))
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
								Range mergeArea = range.MergeArea;
								if (!list2.Contains(mergeArea.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))))
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
									list2.Add(mergeArea.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
									list.Add(mergeArea);
								}
								mergeArea = null;
								break;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						continue;
					}
					while (true)
					{
						switch (1)
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
			if (B != null)
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
				B.ActionEnded();
			}
			list2 = null;
		}
		return list;
	}

	internal static bool A(Range A)
	{
		if (!Information.IsDBNull(RuntimeHelpers.GetObjectValue(A.MergeCells)))
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
					return Conversions.ToBoolean(A.MergeCells);
				}
			}
		}
		return true;
	}

	internal static Range I(Range A)
	{
		new List<Range>();
		Range A2 = null;
		if (!Information.IsDBNull(RuntimeHelpers.GetObjectValue(A.MergeCells)))
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
			if (!Conversions.ToBoolean(A.MergeCells))
			{
				goto IL_00c4;
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
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				try
				{
					if (!Conversions.ToBoolean(range.MergeCells))
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
						RangeHelpers.A(ref A2, range);
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
		A2 = null;
		goto IL_00c4;
		IL_00c4:
		return A2;
	}

	internal static Range D(Worksheet A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		Range result = null;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		application.Calculation = XlCalculation.xlCalculationManual;
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		Range range2;
		Range range;
		try
		{
			range = A.Cells.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
			workbook = application.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value));
			Worksheet obj = (Worksheet)workbook.Worksheets[1];
			((_Worksheet)obj).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(140671);
			range2 = obj.Cells.SpecialCells(XlCellType.xlCellTypeBlanks, RuntimeHelpers.GetObjectValue(Missing.Value));
			result = ((_Worksheet)A).get_Range((object)range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Workbooks.ForceClose(workbook, false);
		}
		XlCalculation calculation = default(XlCalculation);
		application.Calculation = calculation;
		application.EnableEvents = true;
		application.ScreenUpdating = true;
		workbook = null;
		range2 = null;
		range = null;
		application = null;
		return result;
	}

	internal static List<Range> C(Range A)
	{
		List<Range> list = new List<Range>();
		if (A != null)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					if (range.Errors.get_Item((object)XlErrorChecks.xlEmptyCellReferences).Value)
					{
						list.Add(range);
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0065;
					}
					continue;
					end_IL_0065:
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
		return list;
	}

	internal static List<Range> D(Range A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		List<Range> list = new List<Range>();
		application.EnableEvents = false;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)A.Worksheet.Parent;
		XlDisplayDrawingObjects displayDrawingObjects = workbook.DisplayDrawingObjects;
		workbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
		bool.TryParse(Conversions.ToString(A.MergeCells), out var result);
		string right;
		if (result)
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
			right = A.MergeArea.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		else
		{
			right = "";
		}
		string right2 = A.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		long num = 0L;
		checked
		{
			Range range;
			try
			{
				A.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
				bool flag;
				do
				{
					flag = true;
					num++;
					long num2 = 0L;
					while (true)
					{
						num2++;
						range = null;
						try
						{
							range = (Range)A.NavigateArrow(true, num, num2);
							if (Operators.CompareString(range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									list.Add(range);
									break;
								}
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
							break;
						}
						string left = range.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (num > 1)
						{
							if (Operators.CompareString(left, right2, TextCompare: false) == 0)
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
						}
						if (num > 1)
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
							if (result)
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
								if (Operators.CompareString(left, right, TextCompare: false) == 0)
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
							}
						}
						flag = false;
						if (num != 1)
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
						if (num2 <= 100)
						{
							continue;
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
						break;
					}
				}
				while (!flag);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				if (A.Worksheet.ProtectContents)
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
					MessageBox.Show(ex4.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
				else
				{
					Forms.ErrorMessage(ex4.Message);
				}
				ProjectData.ClearProjectError();
			}
			workbook.DisplayDrawingObjects = displayDrawingObjects;
			try
			{
				A.Worksheet.ClearArrows();
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			application.EnableEvents = true;
			application = null;
			workbook = null;
			range = null;
			return list;
		}
	}

	internal static void A(ref Range A, Range B)
	{
		Range range;
		if (A == null)
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
			range = B;
		}
		else
		{
			range = B.Application.Union(A, B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		A = range;
	}

	internal static bool B(Range A)
	{
		return string.IsNullOrEmpty(Conversions.ToString(A.Text));
	}
}
