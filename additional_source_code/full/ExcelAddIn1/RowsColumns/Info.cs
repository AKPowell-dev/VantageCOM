using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class Info
{
	[CompilerGenerated]
	private static Range m_A;

	public static Range CopiedRange
	{
		[CompilerGenerated]
		get
		{
			return Info.m_A;
		}
		[CompilerGenerated]
		set
		{
			Info.m_A = value;
		}
	} = null;

	public static void Copy()
	{
		Application application = MH.A.Application;
		if (application.Selection is Range)
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
			CopiedRange = (Range)application.Selection;
			StatusBar.SetText(VH.A(171581));
		}
		application = null;
	}

	public static void Paste()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		if (CopiedRange == null)
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
					Forms.WarningMessage(VH.A(171626));
					return;
				}
			}
		}
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		checked
		{
			Range activeCell;
			Range range;
			try
			{
				activeCell = application.ActiveCell;
				range = (Range)application.Selection;
				Range copiedRange = CopiedRange;
				long num = Conversions.ToLong(copiedRange.Rows.CountLarge);
				long num2 = Conversions.ToLong(copiedRange.Columns.CountLarge);
				Range usedRange = copiedRange.Worksheet.UsedRange;
				Type typeFromHandle = typeof(Math);
				string memberName = VH.A(53859);
				object[] obj = new object[2]
				{
					num,
					Operators.SubtractObject(Operators.AddObject(usedRange.Rows.CountLarge, usedRange.Row), 1)
				};
				object[] array = obj;
				bool[] obj2 = new bool[2] { true, false };
				bool[] array2 = obj2;
				object value = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
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
					num = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(long));
				}
				long num3 = Conversions.ToLong(value);
				long num4;
				if (Operators.ConditionalCompareObjectEqual(usedRange.Columns.CountLarge, 1, TextCompare: false))
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
					num4 = num2;
				}
				else
				{
					Type typeFromHandle2 = typeof(Math);
					string memberName2 = VH.A(53859);
					object[] obj3 = new object[2]
					{
						num2,
						Operators.SubtractObject(Operators.AddObject(usedRange.Columns.CountLarge, usedRange.Column), 1)
					};
					array = obj3;
					bool[] obj4 = new bool[2] { true, false };
					array2 = obj4;
					object value2 = NewLateBinding.LateGet(null, typeFromHandle2, memberName2, obj3, null, null, obj4);
					if (array2[0])
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
						num2 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(long));
					}
					num4 = Conversions.ToLong(value2);
				}
				usedRange = null;
				if (Operators.CompareString(copiedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), copiedRange.EntireRow.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
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
					if (Operators.CompareString(copiedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), copiedRange.EntireColumn.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
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
						A((int)num3);
						B((int)num4);
						activeCell.get_Resize((object)num3, (object)num4).Select();
						goto IL_07d8;
					}
				}
				if (Operators.CompareString(copiedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), copiedRange.EntireRow.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
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
					if (num == 1)
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
						if (Operators.ConditionalCompareObjectGreater(range.Rows.CountLarge, 1, TextCompare: false))
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
							Range range2 = JH.A(range, application);
							if (range2 != null)
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
								IEnumerator enumerator = default(IEnumerator);
								try
								{
									enumerator = range2.Rows.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Range obj5 = (Range)enumerator.Current;
										obj5.RowHeight = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151632), new object[0], null, null, null));
										obj5.OutlineLevel = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(171725), new object[0], null, null, null));
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
								range2 = null;
							}
							range.EntireRow.Select();
							goto IL_058a;
						}
					}
					A((int)num3);
					activeCell.get_Resize((object)num3, RuntimeHelpers.GetObjectValue(Missing.Value)).EntireRow.Select();
					goto IL_058a;
				}
				if (Operators.CompareString(copiedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), copiedRange.EntireColumn.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
				{
					if (num2 == 1)
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
						if (Operators.ConditionalCompareObjectGreater(range.Columns.CountLarge, 1, TextCompare: false))
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
							Range range3 = JH.A(range, application);
							if (range3 != null)
							{
								IEnumerator enumerator2 = default(IEnumerator);
								try
								{
									enumerator2 = range3.Columns.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										Range obj6 = (Range)enumerator2.Current;
										obj6.ColumnWidth = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151729), new object[0], null, null, null));
										obj6.OutlineLevel = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(171725), new object[0], null, null, null));
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_072d;
										}
										continue;
										end_IL_072d:
										break;
									}
								}
								finally
								{
									if (enumerator2 is IDisposable)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											(enumerator2 as IDisposable).Dispose();
											break;
										}
									}
								}
								range3 = null;
							}
							range.EntireColumn.Select();
							goto IL_079c;
						}
					}
					B((int)num4);
					activeCell.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num4).EntireColumn.Select();
					goto IL_079c;
				}
				num3 = num;
				num4 = num2;
				A((int)num3);
				B((int)num4);
				activeCell.get_Resize((object)num3, (object)num4).Select();
				goto IL_07d8;
				IL_079c:
				activeCell.Activate();
				goto IL_07d8;
				IL_058a:
				activeCell.Activate();
				goto IL_07d8;
				IL_07d8:
				copiedRange = null;
				Core.LogActivity(VH.A(171750));
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
			activeCell = null;
			range = null;
		}
	}

	private static void A(int A)
	{
		Range copiedRange = CopiedRange;
		checked
		{
			for (int i = 1; i <= A; i++)
			{
				Range entireRow = copiedRange.Application.ActiveCell.get_Offset((object)(i - 1), (object)0).EntireRow;
				entireRow.RowHeight = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151632), new object[0], null, null, null));
				entireRow.OutlineLevel = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(171725), new object[0], null, null, null));
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
				copiedRange = null;
				return;
			}
		}
	}

	private static void B(int A)
	{
		Range copiedRange = CopiedRange;
		checked
		{
			for (int i = 1; i <= A; i++)
			{
				Range entireColumn = copiedRange.Application.ActiveCell.get_Offset((object)0, (object)(i - 1)).EntireColumn;
				entireColumn.ColumnWidth = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Columns[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151729), new object[0], null, null, null));
				entireColumn.OutlineLevel = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(copiedRange.Columns[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(171725), new object[0], null, null, null));
			}
			copiedRange = null;
		}
	}
}
