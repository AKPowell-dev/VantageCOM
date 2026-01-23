using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Links
{
	public void ListLinks()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Application application = default(Application);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		Worksheet worksheet = default(Worksheet);
		int num5 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		object objectValue = default(object);
		IEnumerator enumerator2 = default(IEnumerator);
		object objectValue2 = default(object);
		bool flag = default(bool);
		long num6 = default(long);
		long num7 = default(long);
		string text = default(string);
		IEnumerator enumerator3 = default(IEnumerator);
		ChartObject chartObject = default(ChartObject);
		bool flag2 = default(bool);
		bool flag3 = default(bool);
		bool flag4 = default(bool);
		Application application2 = default(Application);
		string text2 = default(string);
		IEnumerator enumerator4 = default(IEnumerator);
		object objectValue3 = default(object);
		object left = default(object);
		IEnumerator enumerator5 = default(IEnumerator);
		Range range = default(Range);
		Range range2 = default(Range);
		Range range3 = default(Range);
		Range range4 = default(Range);
		string right = default(string);
		IEnumerator enumerator6 = default(IEnumerator);
		IEnumerator enumerator7 = default(IEnumerator);
		IEnumerator enumerator8 = default(IEnumerator);
		IEnumerator enumerator9 = default(IEnumerator);
		Range range5 = default(Range);
		Chart chart = default(Chart);
		IEnumerator enumerator10 = default(IEnumerator);
		IEnumerator enumerator11 = default(IEnumerator);
		object objectValue4 = default(object);
		IEnumerator enumerator12 = default(IEnumerator);
		IEnumerator enumerator13 = default(IEnumerator);
		IEnumerator enumerator14 = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						application = MH.A.Application;
						goto IL_0011;
					case 8727:
						{
							num = num2;
							switch (num3)
							{
							case 2:
							case 3:
								break;
							case 1:
								goto IL_1f97;
							default:
								goto end_IL_0000;
							}
							goto IL_1f08;
						}
						IL_1f97:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0011;
						case 3:
							goto IL_001a;
						case 4:
							goto IL_0021;
						case 5:
							goto IL_002d;
						case 6:
							goto IL_0075;
						case 7:
							goto IL_007a;
						case 8:
							goto IL_00a1;
						case 10:
							goto IL_00eb;
						case 12:
							goto IL_010d;
						case 13:
							goto IL_0131;
						case 14:
							goto IL_0137;
						case 15:
							goto IL_013d;
						case 16:
							goto IL_0143;
						case 17:
							goto IL_014a;
						case 18:
							goto IL_019b;
						case 19:
							goto IL_01d2;
						case 20:
							goto IL_01e6;
						case 21:
							goto IL_01ed;
						case 22:
							goto IL_0242;
						case 23:
							goto IL_0256;
						case 24:
							goto IL_0293;
						case 25:
							goto IL_02cc;
						case 26:
							goto IL_036a;
						case 27:
							goto IL_0370;
						case 29:
							goto IL_0379;
						case 28:
						case 30:
							goto IL_0394;
						case 31:
						case 32:
							goto IL_03b6;
						case 33:
							goto IL_03c6;
						case 34:
							goto IL_040e;
						case 35:
							goto IL_0447;
						case 36:
							goto IL_04e1;
						case 38:
							goto IL_0636;
						case 37:
						case 39:
							goto IL_0651;
						case 40:
							goto IL_0673;
						case 41:
							goto IL_06c3;
						case 42:
							goto IL_06c6;
						case 43:
							goto IL_06da;
						case 44:
							goto IL_0702;
						case 45:
							goto IL_0897;
						case 46:
							goto IL_08a0;
						case 47:
							goto IL_08bb;
						case 48:
							goto IL_08dd;
						case 49:
							goto IL_0920;
						case 50:
							goto IL_095b;
						case 51:
							goto IL_099d;
						case 52:
							goto IL_09e7;
						case 53:
							goto IL_0a1e;
						case 54:
							goto IL_0ae4;
						case 55:
							goto IL_0aea;
						case 56:
							goto IL_0c3e;
						case 58:
							goto IL_0c49;
						case 57:
						case 59:
							goto IL_0c64;
						case 60:
							goto IL_0c86;
						case 61:
							goto IL_0ca1;
						case 62:
							goto IL_0cc3;
						case 63:
							goto IL_0ce5;
						case 64:
							goto IL_0d21;
						case 66:
							goto IL_0d8a;
						case 67:
							goto IL_0d97;
						case 68:
							goto IL_0db7;
						case 69:
							goto IL_0dd5;
						case 70:
							goto IL_0e12;
						case 71:
							goto IL_0e94;
						case 72:
							goto IL_0e9a;
						case 73:
							goto IL_0fac;
						case 75:
							goto IL_0fb7;
						case 74:
						case 76:
							goto IL_0fd2;
						case 65:
						case 77:
							goto IL_0fea;
						case 78:
							goto IL_1003;
						case 79:
							goto IL_101b;
						case 80:
							goto IL_1060;
						case 81:
							goto IL_10ab;
						case 82:
							goto IL_10e8;
						case 83:
							goto IL_11b0;
						case 84:
							goto IL_12f8;
						case 85:
							goto IL_1301;
						case 87:
							goto IL_1309;
						case 86:
						case 88:
							goto IL_131a;
						case 89:
							goto IL_133c;
						case 90:
							goto IL_1357;
						case 92:
							goto IL_1381;
						case 93:
							goto IL_13a3;
						case 94:
							goto IL_13a9;
						case 95:
							goto IL_13b0;
						case 96:
							goto IL_13bc;
						case 97:
							goto IL_13c3;
						case 98:
							goto IL_13d7;
						case 99:
							goto IL_1410;
						case 100:
							goto IL_144d;
						case 101:
							goto IL_1493;
						case 102:
							goto IL_14d0;
						case 103:
							goto IL_159e;
						case 104:
							goto IL_15a4;
						case 105:
							goto IL_16e2;
						case 107:
							goto IL_16ed;
						case 106:
						case 108:
							goto IL_1706;
						case 109:
							goto IL_1728;
						case 110:
							goto IL_1739;
						case 111:
							goto IL_1751;
						case 112:
							goto IL_176c;
						case 113:
							goto IL_1776;
						case 114:
							goto IL_1794;
						case 115:
							goto IL_17bc;
						case 116:
							goto IL_17f7;
						case 117:
							goto IL_187d;
						case 118:
							goto IL_1883;
						case 119:
							goto IL_1991;
						case 121:
							goto IL_199c;
						case 120:
						case 122:
							goto IL_19b7;
						case 9:
						case 11:
						case 91:
						case 123:
						case 124:
							goto IL_19d9;
						case 125:
							goto IL_19ea;
						case 126:
							goto IL_1a0c;
						case 128:
							goto IL_1a2c;
						case 129:
							goto IL_1a44;
						case 130:
							goto IL_1a66;
						case 131:
							goto IL_1ab0;
						case 132:
							goto IL_1af0;
						case 133:
							goto IL_1bb3;
						case 134:
							goto IL_1bbc;
						case 135:
							goto IL_1d08;
						case 137:
							goto IL_1d16;
						case 136:
						case 138:
							goto IL_1d2a;
						case 139:
							goto IL_1d4f;
						case 140:
							goto IL_1d66;
						case 141:
							goto IL_1d6f;
						case 142:
							goto IL_1e9b;
						case 143:
							goto IL_1ea7;
						case 127:
						case 144:
							goto IL_1ec7;
						case 145:
							goto IL_1eed;
						case 146:
						case 147:
							goto IL_1f08;
						case 148:
							goto IL_1f11;
						case 149:
							goto IL_1f1f;
						case 150:
							goto IL_1f2d;
						case 151:
							goto IL_1f3b;
						case 152:
							goto IL_1f3e;
						case 153:
							goto IL_1f46;
						case 154:
							goto IL_1f4f;
						case 155:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 156:
							goto end_IL_0000_3;
						}
						goto default;
						IL_0011:
						num2 = 2;
						application.ScreenUpdating = false;
						goto IL_001a;
						IL_001a:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_0021;
						IL_0021:
						num2 = 4;
						workbook = application.ActiveWorkbook;
						goto IL_002d;
						IL_002d:
						num2 = 5;
						worksheet = (Worksheet)application.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0075;
						IL_0075:
						num2 = 6;
						num5 = 1;
						goto IL_007a;
						IL_007a:
						num2 = 7;
						enumerator = workbook.Sheets.GetEnumerator();
						goto IL_19dc;
						IL_19dc:
						if (enumerator.MoveNext())
						{
							objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							goto IL_00a1;
						}
						goto IL_19ea;
						IL_19ea:
						num2 = 125;
						if (enumerator is IDisposable)
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
							(enumerator as IDisposable).Dispose();
						}
						goto IL_1a0c;
						IL_0fba:
						if (enumerator2.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
							goto IL_0e12;
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
						goto IL_0fd2;
						IL_159e:
						num2 = 103;
						flag = true;
						goto IL_15a4;
						IL_1a0c:
						num2 = 126;
						if (workbook.Names.Count != 0)
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
							goto IL_1a2c;
						}
						goto IL_1ec7;
						IL_187d:
						num2 = 117;
						flag = true;
						goto IL_1883;
						IL_1a2c:
						num2 = 128;
						application.StatusBar = VH.A(153600);
						goto IL_1a44;
						IL_1a44:
						num2 = 129;
						num6 = workbook.Names.Count;
						num7 = 1L;
						goto IL_1eb4;
						IL_1eb4:
						if (num7 <= num6)
						{
							goto IL_1a66;
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
						goto IL_1ec7;
						IL_0e12:
						num2 = 70;
						if ((Strings.InStr(Strings.Replace(text, VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(text, Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_0e94;
						}
						goto IL_0fb7;
						IL_1a66:
						num2 = 130;
						if (Strings.InStr(Conversions.ToString(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7)), VH.A(6144)) != 0)
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
							goto IL_1ab0;
						}
						goto IL_1ea7;
						IL_1883:
						num2 = 118;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153551), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), Strings.Replace(text, VH.A(2378), VH.A(77635))));
						goto IL_1991;
						IL_1ab0:
						num2 = 131;
						enumerator3 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_1d1c;
						IL_1d1c:
						if (enumerator3.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator3.Current);
							goto IL_1af0;
						}
						goto IL_1d2a;
						IL_0e94:
						num2 = 71;
						flag = true;
						goto IL_0e9a;
						IL_1af0:
						num2 = 132;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7)), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_1bb3;
						}
						goto IL_1d16;
						IL_0e9a:
						num2 = 72;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153406) + chartObject.Name + VH.A(2378), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), text));
						goto IL_0fac;
						IL_1bb3:
						num2 = 133;
						flag2 = true;
						goto IL_1bbc;
						IL_1bbc:
						num2 = 134;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153643), NewLateBinding.LateGet(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7), null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), NewLateBinding.LateGet(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7), null, VH.A(153696), new object[0], null, null, null)));
						goto IL_1d08;
						IL_1d08:
						num2 = 135;
						num5++;
						goto IL_1d2a;
						IL_1d2a:
						num2 = 138;
						if (enumerator3 is IDisposable)
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
							(enumerator3 as IDisposable).Dispose();
						}
						goto IL_1d4f;
						IL_0fac:
						num2 = 73;
						num5++;
						goto IL_0fd2;
						IL_0fd2:
						num2 = 76;
						if (enumerator2 is IDisposable)
						{
							(enumerator2 as IDisposable).Dispose();
						}
						goto IL_0fea;
						IL_1d4f:
						num2 = 139;
						if (!flag2)
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
							goto IL_1d66;
						}
						goto IL_1ea7;
						IL_1991:
						num2 = 119;
						num5++;
						goto IL_19b7;
						IL_1d66:
						num2 = 140;
						flag2 = true;
						goto IL_1d6f;
						IL_1d6f:
						num2 = 141;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153643), NewLateBinding.LateGet(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7), null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), workbook.Names.Cast<object>().ElementAtOrDefault((int)num7)), VH.A(153224)), Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(workbook.Names.Cast<object>().ElementAtOrDefault((int)num7), null, VH.A(153696), new object[0], null, null, null)), VH.A(2378), VH.A(77635))));
						goto IL_1e9b;
						IL_1e9b:
						num2 = 142;
						num5++;
						goto IL_1ea7;
						IL_1d16:
						num2 = 137;
						goto IL_1d1c;
						IL_1ea7:
						num2 = 143;
						num7++;
						goto IL_1eb4;
						IL_1ec7:
						num2 = 144;
						if (unchecked(flag == flag2 == flag3 == flag4 && !flag3))
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
							goto IL_1eed;
						}
						goto IL_1f08;
						IL_0fb7:
						num2 = 75;
						goto IL_0fba;
						IL_1eed:
						num2 = 145;
						Interaction.MsgBox(VH.A(153713), MsgBoxStyle.Critical);
						goto IL_1f08;
						IL_1f08:
						num2 = 147;
						application2 = application;
						goto IL_1f11;
						IL_1f11:
						num2 = 148;
						application2.StatusBar = null;
						goto IL_1f1f;
						IL_1f1f:
						num2 = 149;
						application2.DisplayAlerts = true;
						goto IL_1f2d;
						IL_1f2d:
						num2 = 150;
						application2.ScreenUpdating = true;
						goto IL_1f3b;
						IL_1f3b:
						application2 = null;
						goto IL_1f3e;
						IL_1f3e:
						num2 = 152;
						application = null;
						goto IL_1f46;
						IL_1f46:
						num2 = 153;
						workbook = null;
						goto IL_1f4f;
						IL_1f4f:
						num2 = 154;
						if (Operators.CompareString(text2, null, TextCompare: false) == 0)
						{
							goto end_IL_0000_3;
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
						break;
						IL_0fea:
						num2 = 77;
						goto IL_0fed;
						IL_0dd5:
						num2 = 69;
						enumerator2 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_0fba;
						IL_00a1:
						num2 = 8;
						if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null), worksheet.Name, TextCompare: false))
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
							goto IL_00eb;
						}
						goto IL_19d9;
						IL_17bc:
						num2 = 115;
						enumerator4 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_199f;
						IL_099d:
						num2 = 51;
						if (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(6144)) != 0)
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
							goto IL_09e7;
						}
						goto IL_0c86;
						IL_19b7:
						num2 = 122;
						if (enumerator4 is IDisposable)
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
							(enumerator4 as IDisposable).Dispose();
						}
						goto IL_19d9;
						IL_00eb:
						num2 = 10;
						left = NewLateBinding.LateGet(objectValue, null, VH.A(101861), new object[0], null, null, null);
						goto IL_010d;
						IL_010d:
						num2 = 12;
						if (Operators.ConditionalCompareObjectEqual(left, XlSheetType.xlWorksheet, TextCompare: false))
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
							goto IL_0131;
						}
						goto IL_1381;
						IL_09e7:
						num2 = 52;
						enumerator5 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_0c4c;
						IL_0131:
						num2 = 13;
						range = null;
						goto IL_0137;
						IL_0137:
						num2 = 14;
						range2 = null;
						goto IL_013d;
						IL_013d:
						num2 = 15;
						range3 = null;
						goto IL_0143;
						IL_0143:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_014a;
						IL_014a:
						num2 = 17;
						range = (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, VH.A(62391), new object[0], null, null, null), null, VH.A(86222), new object[1] { XlCellType.xlCellTypeFormulas }, null, null, null);
						goto IL_019b;
						IL_019b:
						num2 = 18;
						application.StatusBar = Operators.ConcatenateObject(VH.A(153127), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null));
						goto IL_01d2;
						IL_01d2:
						num2 = 19;
						if (range != null)
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
							goto IL_01e6;
						}
						goto IL_06c6;
						IL_0c4c:
						if (enumerator5.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator5.Current);
							goto IL_0a1e;
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
						goto IL_0c64;
						IL_01e6:
						num2 = 20;
						range4 = range;
						goto IL_01ed;
						IL_01ed:
						num2 = 21;
						range2 = range4.Find(VH.A(153184), RuntimeHelpers.GetObjectValue(Missing.Value), XlFindLookIn.xlFormulas, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0242;
						IL_0242:
						num2 = 22;
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
							goto IL_0256;
						}
						goto IL_06c3;
						IL_15a4:
						num2 = 104;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153500), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(2378), VH.A(77635))));
						goto IL_16e2;
						IL_0256:
						num2 = 23;
						right = range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0293;
						IL_0293:
						num2 = 24;
						enumerator6 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_037c;
						IL_037c:
						if (enumerator6.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator6.Current);
							goto IL_02cc;
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
						goto IL_0394;
						IL_16e2:
						num2 = 105;
						num5++;
						goto IL_1706;
						IL_0a1e:
						num2 = 53;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
						{
							goto IL_0ae4;
						}
						goto IL_0c49;
						IL_02cc:
						num2 = 25;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(range2.Formula), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(range2.Formula), Strings.Right(Conversions.ToString(range2.Formula), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_036a;
						}
						goto IL_0379;
						IL_0ae4:
						num2 = 54;
						flag = true;
						goto IL_0aea;
						IL_036a:
						num2 = 26;
						flag3 = true;
						goto IL_0370;
						IL_0370:
						num2 = 27;
						range3 = range2;
						goto IL_0394;
						IL_0394:
						num2 = 30;
						if (enumerator6 is IDisposable)
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
							(enumerator6 as IDisposable).Dispose();
						}
						goto IL_03b6;
						IL_0aea:
						num2 = 55;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153307) + chartObject.Name + VH.A(2378), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(2378), VH.A(77635))));
						goto IL_0c3e;
						IL_0c3e:
						num2 = 56;
						num5++;
						goto IL_0c64;
						IL_03b6:
						num2 = 32;
						range2 = range4.FindNext(range2);
						goto IL_03c6;
						IL_03c6:
						num2 = 33;
						if (Operators.CompareString(range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) != 0)
						{
							goto IL_040e;
						}
						goto IL_0673;
						IL_040e:
						num2 = 34;
						enumerator7 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_0639;
						IL_0639:
						if (enumerator7.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator7.Current);
							goto IL_0447;
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
						goto IL_0651;
						IL_0c64:
						num2 = 59;
						if (enumerator5 is IDisposable)
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
							(enumerator5 as IDisposable).Dispose();
						}
						goto IL_0c86;
						IL_199c:
						num2 = 121;
						goto IL_199f;
						IL_0447:
						num2 = 35;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(range2.Formula), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(range2.Formula), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_04e1;
						}
						goto IL_0636;
						IL_1493:
						num2 = 101;
						enumerator8 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_16f0;
						IL_04e1:
						num2 = 36;
						range3 = application.Union(range3, range2, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0651;
						IL_0651:
						num2 = 39;
						if (enumerator7 is IDisposable)
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
							(enumerator7 as IDisposable).Dispose();
						}
						goto IL_0673;
						IL_0c49:
						num2 = 58;
						goto IL_0c4c;
						IL_0c86:
						num2 = 60;
						goto IL_0c89;
						IL_0636:
						num2 = 38;
						goto IL_0639;
						IL_0673:
						num2 = 40;
						if (Operators.CompareString(range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) != 0)
						{
							goto IL_03b6;
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
						goto IL_06c3;
						IL_1381:
						num2 = 92;
						if (Operators.ConditionalCompareObjectEqual(left, 3, TextCompare: false))
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
							goto IL_13a3;
						}
						goto IL_19d9;
						IL_0379:
						num2 = 29;
						goto IL_037c;
						IL_06c3:
						range4 = null;
						goto IL_06c6;
						IL_06c6:
						num2 = 42;
						if (range3 != null)
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
							goto IL_06da;
						}
						goto IL_08dd;
						IL_144d:
						num2 = 100;
						if (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(6144)) != 0)
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
							goto IL_1493;
						}
						goto IL_1728;
						IL_06da:
						num2 = 43;
						enumerator9 = range3.Areas.GetEnumerator();
						goto IL_08a3;
						IL_08a3:
						if (enumerator9.MoveNext())
						{
							range5 = (Range)enumerator9.Current;
							goto IL_0702;
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
						goto IL_08bb;
						IL_13a3:
						num2 = 93;
						chart = null;
						goto IL_13a9;
						IL_08bb:
						num2 = 47;
						if (enumerator9 is IDisposable)
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
							(enumerator9 as IDisposable).Dispose();
						}
						goto IL_08dd;
						IL_13b0:
						num2 = 95;
						chart = (Chart)objectValue;
						goto IL_13bc;
						IL_13bc:
						ProjectData.ClearProjectError();
						num3 = 3;
						goto IL_13c3;
						IL_13a9:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_13b0;
						IL_0702:
						num2 = 44;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153195), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(7827)), Strings.Replace(range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(2378), VH.A(77635))), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), NewLateBinding.LateGet(range3.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[0], null, null, null)));
						goto IL_0897;
						IL_0897:
						num2 = 45;
						num5++;
						goto IL_08a0;
						IL_08a0:
						num2 = 46;
						goto IL_08a3;
						IL_08dd:
						num2 = 48;
						enumerator10 = ((IEnumerable)NewLateBinding.LateGet(objectValue, null, VH.A(153229), new object[0], null, null, null)).GetEnumerator();
						goto IL_0fed;
						IL_0fed:
						if (enumerator10.MoveNext())
						{
							chartObject = (ChartObject)enumerator10.Current;
							goto IL_0920;
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
						goto IL_1003;
						IL_13c3:
						num2 = 97;
						if (chart != null)
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
							goto IL_13d7;
						}
						goto IL_19d9;
						IL_1003:
						num2 = 78;
						if (enumerator10 is IDisposable)
						{
							(enumerator10 as IDisposable).Dispose();
						}
						goto IL_101b;
						IL_13d7:
						num2 = 98;
						application.StatusBar = Operators.ConcatenateObject(VH.A(153254), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null));
						goto IL_1410;
						IL_101b:
						num2 = 79;
						enumerator11 = ((IEnumerable)NewLateBinding.LateGet(objectValue, null, VH.A(153431), new object[0], null, null, null)).GetEnumerator();
						goto IL_133f;
						IL_133f:
						if (enumerator11.MoveNext())
						{
							objectValue4 = RuntimeHelpers.GetObjectValue(enumerator11.Current);
							goto IL_1060;
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
						goto IL_1357;
						IL_1410:
						num2 = 99;
						enumerator12 = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_172b;
						IL_1357:
						num2 = 90;
						if (enumerator11 is IDisposable)
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
							(enumerator11 as IDisposable).Dispose();
						}
						goto IL_19d9;
						IL_1739:
						num2 = 110;
						if (enumerator12 is IDisposable)
						{
							(enumerator12 as IDisposable).Dispose();
						}
						goto IL_1751;
						IL_16f0:
						if (enumerator8.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator8.Current);
							goto IL_14d0;
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
						goto IL_1706;
						IL_172b:
						if (enumerator12.MoveNext())
						{
							objectValue3 = RuntimeHelpers.GetObjectValue(enumerator12.Current);
							goto IL_144d;
						}
						goto IL_1739;
						IL_1060:
						num2 = 80;
						if (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue4, null, VH.A(153454), new object[0], null, null, null)), VH.A(6144)) > 0)
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
							goto IL_10ab;
						}
						goto IL_133c;
						IL_1751:
						num2 = 111;
						if (chart.HasTitle)
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
							goto IL_176c;
						}
						goto IL_19d9;
						IL_10ab:
						num2 = 81;
						enumerator13 = ((IEnumerable)workbook.LinkSources(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_130c;
						IL_130c:
						if (enumerator13.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator13.Current);
							goto IL_10e8;
						}
						goto IL_131a;
						IL_1728:
						num2 = 109;
						goto IL_172b;
						IL_10e8:
						num2 = 82;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(objectValue4, null, VH.A(153454), new object[0], null, null, null)), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue4, null, VH.A(153454), new object[0], null, null, null)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_11b0;
						}
						goto IL_1309;
						IL_176c:
						num2 = 112;
						chart.Activate();
						goto IL_1776;
						IL_11b0:
						num2 = 83;
						((_Worksheet)worksheet).get_Range((object)(VH.A(57237) + Conversions.ToString(num5)), RuntimeHelpers.GetObjectValue(Missing.Value)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(153475), NewLateBinding.LateGet(objectValue4, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(2378)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))), VH.A(2378)), objectValue2), VH.A(153224)), NewLateBinding.LateGet(objectValue4, null, VH.A(153454), new object[0], null, null, null)));
						goto IL_12f8;
						IL_12f8:
						num2 = 84;
						num5++;
						goto IL_1301;
						IL_1301:
						num2 = 85;
						flag4 = true;
						goto IL_131a;
						IL_131a:
						num2 = 88;
						if (enumerator13 is IDisposable)
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
							(enumerator13 as IDisposable).Dispose();
						}
						goto IL_133c;
						IL_1776:
						num2 = 113;
						text = Conversions.ToString(application.ExecuteExcel4Macro(VH.A(153365)));
						goto IL_1794;
						IL_1794:
						num2 = 114;
						if (Strings.InStr(text, VH.A(6144)) != 0)
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
							goto IL_17bc;
						}
						goto IL_19d9;
						IL_1309:
						num2 = 87;
						goto IL_130c;
						IL_133c:
						num2 = 89;
						goto IL_133f;
						IL_1706:
						num2 = 108;
						if (enumerator8 is IDisposable)
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
							(enumerator8 as IDisposable).Dispose();
						}
						goto IL_1728;
						IL_0920:
						num2 = 49;
						application.StatusBar = Operators.ConcatenateObject(VH.A(153254), NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null));
						goto IL_095b;
						IL_095b:
						num2 = 50;
						enumerator14 = ((IEnumerable)chartObject.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						goto IL_0c89;
						IL_0c89:
						if (enumerator14.MoveNext())
						{
							objectValue3 = RuntimeHelpers.GetObjectValue(enumerator14.Current);
							goto IL_099d;
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
						goto IL_0ca1;
						IL_19d9:
						num2 = 124;
						goto IL_19dc;
						IL_0ca1:
						num2 = 61;
						if (enumerator14 is IDisposable)
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
							(enumerator14 as IDisposable).Dispose();
						}
						goto IL_0cc3;
						IL_199f:
						if (enumerator4.MoveNext())
						{
							objectValue2 = RuntimeHelpers.GetObjectValue(enumerator4.Current);
							goto IL_17f7;
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
						goto IL_19b7;
						IL_14d0:
						num2 = 102;
						if ((Strings.InStr(Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(Conversions.ToString(NewLateBinding.LateGet(objectValue3, null, VH.A(68956), new object[0], null, null, null)), Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_159e;
						}
						goto IL_16ed;
						IL_0cc3:
						num2 = 62;
						if (chartObject.Chart.HasTitle)
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
							goto IL_0ce5;
						}
						goto IL_0fea;
						IL_16ed:
						num2 = 107;
						goto IL_16f0;
						IL_0ce5:
						num2 = 63;
						if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue, null, VH.A(153334), new object[0], null, null, null), true, TextCompare: false))
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
							goto IL_0d21;
						}
						goto IL_0d8a;
						IL_17f7:
						num2 = 116;
						if ((Strings.InStr(Strings.Replace(text, VH.A(7120), null), Conversions.ToString(objectValue2)) > 0) | (Strings.InStr(text, Strings.Right(Conversions.ToString(objectValue2), Strings.Len(RuntimeHelpers.GetObjectValue(objectValue2)) - Strings.InStrRev(Conversions.ToString(objectValue2), VH.A(48247)))) > 0))
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
							goto IL_187d;
						}
						goto IL_199c;
						IL_0d21:
						num2 = 64;
						text2 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(text2, NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null)), VH.A(56830)), chartObject.Name), VH.A(7803)));
						goto IL_0fea;
						IL_0d8a:
						num2 = 66;
						chartObject.Activate();
						goto IL_0d97;
						IL_0d97:
						num2 = 67;
						text = Conversions.ToString(application.ExecuteExcel4Macro(VH.A(153365)));
						goto IL_0db7;
						IL_0db7:
						num2 = 68;
						if (Strings.InStr(text, VH.A(6144)) != 0)
						{
							goto IL_0dd5;
						}
						goto IL_0fea;
						end_IL_0000_2:
						break;
					}
					num2 = 155;
					Interaction.MsgBox(VH.A(153758) + text2, MsgBoxStyle.Critical);
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 8727;
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
			switch (5)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}
