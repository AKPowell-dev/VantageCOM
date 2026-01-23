using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Paintbrush
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Paintbrush.m_A;
		}
		[CompilerGenerated]
		set
		{
			Paintbrush.m_A = value;
		}
	}

	public static void Capture()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		int num5 = default(int);
		Application application = default(Application);
		Range range = default(Range);
		Range range2 = default(Range);
		int num6 = default(int);
		string[] array = default(string[]);
		Worksheet worksheet = default(Worksheet);
		Worksheet worksheet2 = default(Worksheet);
		Range range3 = default(Range);
		int num7 = default(int);
		object instance = default(object);
		int count = default(int);
		int num8 = default(int);
		Application application2 = default(Application);
		int count2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
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
						goto IL_0021;
					case 1249:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000_2;
							}
							int num4 = unchecked(num + 1);
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
								goto IL_003a;
							case 6:
								goto IL_0047;
							case 8:
								goto IL_0051;
							case 9:
								goto IL_006e;
							case 10:
								goto IL_0079;
							case 11:
								goto IL_0084;
							case 12:
								goto IL_008e;
							case 13:
								goto IL_0095;
							case 14:
								goto IL_00a8;
							case 15:
								goto IL_00c4;
							case 16:
								goto IL_013a;
							case 17:
								goto IL_0153;
							case 18:
								goto IL_0172;
							case 19:
								goto IL_01b3;
							case 20:
								goto IL_0206;
							case 21:
								goto IL_0209;
							case 22:
								goto IL_020c;
							case 23:
								goto IL_026a;
							case 24:
								goto IL_0270;
							case 25:
								goto IL_028c;
							case 26:
								goto IL_02a0;
							case 27:
								goto IL_02bb;
							case 28:
								goto IL_02c4;
							case 29:
								goto IL_02e4;
							case 30:
								goto IL_0311;
							case 31:
								goto IL_0341;
							case 32:
								goto IL_0383;
							case 33:
								goto IL_039f;
							case 34:
								goto IL_03a2;
							case 35:
								goto IL_03bb;
							case 36:
								goto IL_03c5;
							case 37:
								goto IL_03cb;
							case 38:
								goto IL_03d1;
							case 39:
								goto IL_03d8;
							case 40:
								goto IL_03e3;
							case 41:
								goto IL_03ee;
							case 42:
								goto IL_03f9;
							case 43:
								goto IL_03fc;
							case 44:
								goto end_IL_0000_3;
							default:
								goto end_IL_0000_2;
							case 2:
							case 7:
							case 45:
								goto end_IL_0000;
							}
							goto default;
						}
						IL_02bb:
						num2 = 27;
						num5++;
						goto IL_02c4;
						IL_02c4:
						num2 = 28;
						if (num5 > KH.A.PaintbrushesLimit)
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
							goto IL_02e4;
						}
						goto IL_0383;
						IL_03fc:
						num2 = 43;
						application = null;
						break;
						IL_02e4:
						num2 = 29;
						range = (Range)range2.Cells[num6, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0311;
						IL_0021:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0028;
						IL_0028:
						num2 = 4;
						application = MH.A.Application;
						goto IL_003a;
						IL_003a:
						num2 = 5;
						if (EditMode.IsEditMode(application))
						{
							goto IL_0047;
						}
						goto IL_0051;
						IL_0047:
						num2 = 6;
						application = null;
						goto end_IL_0000;
						IL_0051:
						num2 = 8;
						if (application.Selection is Range)
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
							goto IL_006e;
						}
						goto IL_03fc;
						IL_0311:
						num2 = 30;
						array = Strings.Split(Conversions.ToString(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))), VH.A(2378));
						goto IL_0341;
						IL_006e:
						num2 = 9;
						application.ScreenUpdating = false;
						goto IL_0079;
						IL_0079:
						num2 = 10;
						application.EnableEvents = false;
						goto IL_0084;
						IL_0084:
						num2 = 11;
						worksheet = A();
						goto IL_008e;
						IL_008e:
						num2 = 12;
						worksheet2 = worksheet;
						goto IL_0095;
						IL_0095:
						num2 = 13;
						range3 = (Range)application.Selection;
						goto IL_00a8;
						IL_00a8:
						num2 = 14;
						num7 = Conversions.ToInteger(range3.Rows.CountLarge);
						goto IL_00c4;
						IL_00c4:
						num2 = 15;
						((_Worksheet)worksheet2).get_Range(RuntimeHelpers.GetObjectValue(worksheet2.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet2.Cells[num7, 1])).EntireRow.Insert(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_013a;
						IL_013a:
						num2 = 16;
						range3.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0153;
						IL_0153:
						num2 = 17;
						instance = worksheet2.Cells[1, 1];
						goto IL_0172;
						IL_0172:
						num2 = 18;
						NewLateBinding.LateCall(instance, null, VH.A(147355), new object[1] { XlPasteType.xlPasteFormats }, new string[1] { VH.A(1102) }, null, null, IgnoreReturn: true);
						goto IL_01b3;
						IL_01b3:
						num2 = 19;
						NewLateBinding.LateSetComplex(instance, null, VH.A(41636), new object[1] { Operators.ConcatenateObject(Conversions.ToString(num7) + VH.A(2378), range3.Columns.CountLarge) }, null, null, OptimisticSet: false, RValueBase: true);
						goto IL_0206;
						IL_0206:
						instance = null;
						goto IL_0209;
						IL_0209:
						worksheet2 = null;
						goto IL_020c;
						IL_020c:
						num2 = 22;
						range3 = (Range)NewLateBinding.LateGet(worksheet.UsedRange.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(86222), new object[2]
						{
							XlCellType.xlCellTypeConstants,
							XlSpecialCellsValue.xlTextValues
						}, null, null, null);
						goto IL_026a;
						IL_026a:
						num2 = 23;
						num5 = 0;
						goto IL_0270;
						IL_0270:
						num2 = 24;
						count = range3.Areas.Count;
						num8 = 1;
						goto IL_03a9;
						IL_03a9:
						if (num8 <= count)
						{
							goto IL_028c;
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
						goto IL_03bb;
						IL_0341:
						num2 = 31;
						((_Worksheet)worksheet).get_Range((object)range, (object)range.get_Offset((object)(Conversions.ToDouble(array[0]) - 1.0), (object)0)).EntireRow.Clear();
						goto IL_0383;
						IL_03bb:
						num2 = 35;
						CycleIndex = num5;
						goto IL_03c5;
						IL_03c5:
						num2 = 36;
						range3 = null;
						goto IL_03cb;
						IL_03cb:
						num2 = 37;
						range = null;
						goto IL_03d1;
						IL_03d1:
						num2 = 38;
						application2 = application;
						goto IL_03d8;
						IL_03d8:
						num2 = 39;
						application2.EnableEvents = true;
						goto IL_03e3;
						IL_03e3:
						num2 = 40;
						application2.ScreenUpdating = true;
						goto IL_03ee;
						IL_03ee:
						num2 = 41;
						application2.CutCopyMode = (XlCutCopyMode)0;
						goto IL_03f9;
						IL_03f9:
						application2 = null;
						goto IL_03fc;
						IL_028c:
						num2 = 25;
						range2 = range3.Areas[num8];
						goto IL_02a0;
						IL_02a0:
						num2 = 26;
						count2 = range2.Cells.Count;
						num6 = 1;
						goto IL_038c;
						IL_038c:
						if (num6 <= count2)
						{
							goto IL_02bb;
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
						goto IL_039f;
						IL_0383:
						num2 = 32;
						num6++;
						goto IL_038c;
						IL_039f:
						range2 = null;
						goto IL_03a2;
						IL_03a2:
						num2 = 34;
						num8++;
						goto IL_03a9;
						end_IL_0000_3:
						break;
					}
					num2 = 44;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(147380));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1249;
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
			switch (4)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void Apply()
	{
		Range range = null;
		Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
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
					application = null;
					return;
				}
			}
		}
		Range range2;
		Range range4;
		Range activeCell;
		Worksheet worksheet;
		try
		{
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			worksheet = A();
			object obj = NewLateBinding.LateGet(worksheet.UsedRange.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(86222), new object[1] { XlCellType.xlCellTypeConstants }, null, null, null);
			int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(obj, null, VH.A(147417), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
			int num2 = 1;
			Range range3;
			int num6;
			int num7;
			int num4;
			checked
			{
				int num5 = default(int);
				while (true)
				{
					object[] array;
					bool[] array2;
					if (num2 <= num)
					{
						object instance = NewLateBinding.LateGet(obj, null, VH.A(147417), array = new object[1] { num2 }, null, null, array2 = new bool[1] { true });
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
							num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
						}
						int num3 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(62391), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
						num4 = 1;
						while (num4 <= num3)
						{
							num5++;
							if (num5 != CycleIndex + 1)
							{
								num4++;
								continue;
							}
							goto IL_01a8;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_029c;
							}
							continue;
							end_IL_029c:
							break;
						}
						num2++;
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
					if (range == null)
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
						range = (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(obj, null, VH.A(147417), new object[1] { 1 }, null, null, null), null, VH.A(62391), new object[1] { 1 }, null, null, null);
						CycleIndex = 1;
						obj = null;
						break;
					}
					throw new Exception();
					IL_01a8:
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					object instance2 = obj;
					string memberName = VH.A(147417);
					object[] obj2 = new object[1] { num2 };
					object[] array3 = obj2;
					bool[] obj3 = new bool[1] { true };
					bool[] array4 = obj3;
					object instance3 = NewLateBinding.LateGet(instance2, null, memberName, obj2, null, null, obj3);
					if (array4[0])
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
						num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(int));
					}
					object obj4 = NewLateBinding.LateGet(instance3, null, VH.A(62391), array = new object[1] { num4 }, null, null, array2 = new bool[1] { true });
					if (array2[0])
					{
						num4 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
					}
					range = (Range)obj4;
					CycleIndex++;
					break;
				}
				activeCell = application.ActiveCell;
				range2 = (Range)NewLateBinding.LateGet(application.Selection, null, VH.A(62391), new object[2] { 1, 1 }, null, null, null);
				string[] array5 = Strings.Split(Conversions.ToString(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))), VH.A(2378));
				range3 = ((_Worksheet)worksheet).get_Range((object)range, (object)range.get_Offset((object)(Conversions.ToDouble(array5[0]) - 1.0), (object)(Conversions.ToDouble(array5[1]) - 1.0)));
				range4 = (Range)application.Selection;
				num6 = Conversions.ToInteger(range3.Rows.CountLarge);
				num7 = Conversions.ToInteger(range3.Columns.CountLarge);
				num2 = Conversions.ToInteger(range4.Rows.CountLarge);
				num4 = Conversions.ToInteger(range4.Columns.CountLarge);
			}
			if (num2 % num6 != 0 || num4 % num7 != 0)
			{
				range4 = range2.get_Resize((object)num6, (object)num7);
			}
			bool flag = JH.A(range4);
			range3.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
			range4.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			range4.Select();
			activeCell.Activate();
			if (flag)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					JH.A(range4, VH.A(147428));
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
		Application application2 = application;
		application2.EnableEvents = true;
		application2.ScreenUpdating = true;
		application2.CutCopyMode = (XlCutCopyMode)0;
		_ = null;
		range = null;
		range2 = null;
		range4 = null;
		activeCell = null;
		worksheet = null;
		application = null;
		if (CycleIndex != 1)
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
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(147449));
			return;
		}
	}

	private static Worksheet A()
	{
		Application application = MH.A.Application;
		return (Worksheet)application.Workbooks[application.AddIns[clsUtilities.ADDIN_NAME].Name].Worksheets[1];
	}

	public static void Clear()
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		try
		{
			A().Cells.Clear();
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application.EnableEvents = true;
		application.ScreenUpdating = true;
		application = null;
		CycleIndex = 0;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(147482));
	}
}
