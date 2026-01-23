using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Summarize
{
	public static void Go()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		int[] array = default(int[]);
		int num5 = default(int);
		int num6 = default(int);
		string text = default(string);
		string[] array2 = default(string[]);
		Application application = default(Application);
		string pattern = default(string);
		Application application2 = default(Application);
		Range range = default(Range);
		Worksheet worksheet = default(Worksheet);
		string text2 = default(string);
		MatchCollection matchCollection = default(MatchCollection);
		IEnumerator enumerator = default(IEnumerator);
		Match match = default(Match);
		int num7 = default(int);
		IEnumerator enumerator2 = default(IEnumerator);
		Match match2 = default(Match);
		Match match3 = default(Match);
		int num8 = default(int);
		string[] array3 = default(string[]);
		string[] array4 = default(string[]);
		int num9 = default(int);
		string text3 = default(string);
		Range range2 = default(Range);
		bool flag = default(bool);
		Range range3 = default(Range);
		Range range4 = default(Range);
		Range range5 = default(Range);
		Application application3 = default(Application);
		string text4 = default(string);
		Range range6 = default(Range);
		Match match4 = default(Match);
		string[] array5 = default(string[]);
		int num11 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					object instance;
					string memberName;
					object[] array6;
					ref string reference;
					object[] array7;
					bool[] obj;
					bool[] array8;
					object obj2;
					int num10;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
						goto IL_0024;
					case 4449:
						{
							num = num2;
							switch (num3)
							{
							case 2:
								break;
							case 1:
								goto IL_0f11;
							default:
								goto end_IL_0000_2;
							}
							goto IL_0e90;
						}
						IL_0f11:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0024;
						case 4:
							goto IL_0034;
						case 5:
							goto IL_0039;
						case 6:
							goto IL_0049;
						case 7:
							goto IL_0050;
						case 8:
							goto IL_0056;
						case 9:
							goto IL_0061;
						case 10:
							goto IL_00c3;
						case 12:
							goto IL_00ed;
						case 13:
							goto IL_00fb;
						case 14:
							goto IL_0106;
						case 15:
							goto IL_0111;
						case 16:
							goto IL_011c;
						case 17:
							goto IL_011f;
						case 18:
							goto IL_013c;
						case 19:
							goto IL_0156;
						case 20:
							goto IL_016b;
						case 21:
							goto IL_0182;
						case 22:
							goto IL_01a5;
						case 23:
							goto IL_01ac;
						case 24:
							goto IL_01bd;
						case 25:
							goto IL_01ce;
						case 26:
							goto IL_01d7;
						case 27:
							goto IL_0213;
						case 28:
							goto IL_0216;
						case 29:
							goto IL_0227;
						case 30:
							goto IL_0249;
						case 31:
							goto IL_0261;
						case 32:
							goto IL_0274;
						case 33:
							goto IL_0290;
						case 34:
							goto IL_02ac;
						case 35:
							goto IL_02cf;
						case 36:
							goto IL_02d6;
						case 37:
							goto IL_02f4;
						case 38:
							goto IL_0305;
						case 39:
							goto IL_030e;
						case 40:
							goto IL_034a;
						case 41:
							goto IL_034d;
						case 42:
							goto IL_0368;
						case 43:
							goto IL_038a;
						case 44:
							goto IL_0399;
						case 45:
							goto IL_03a8;
						case 46:
							goto IL_03b7;
						case 47:
							goto IL_03c6;
						case 48:
							goto IL_03d0;
						case 49:
							goto IL_03dd;
						case 50:
							goto IL_03e7;
						case 51:
							goto IL_03f8;
						case 52:
							goto IL_0405;
						case 53:
							goto IL_0416;
						case 54:
							goto IL_042f;
						case 55:
							goto IL_0441;
						case 56:
							goto IL_0451;
						case 57:
							goto IL_0461;
						case 58:
							goto IL_046b;
						case 59:
							goto IL_0472;
						case 60:
							goto IL_0478;
						case 61:
							goto IL_047e;
						case 62:
							goto IL_04b7;
						case 63:
							goto IL_04cd;
						case 65:
							goto IL_04e4;
						case 66:
							goto IL_0533;
						case 67:
							goto IL_053d;
						case 68:
							goto IL_0556;
						case 69:
							goto IL_0570;
						case 70:
							goto IL_05c0;
						case 71:
							goto IL_05db;
						case 72:
							goto IL_0607;
						case 73:
							goto IL_062a;
						case 75:
							goto IL_0665;
						case 76:
							goto IL_068c;
						case 77:
							goto IL_06be;
						case 78:
							goto IL_06c7;
						case 79:
							goto IL_06ea;
						case 81:
							goto IL_0723;
						case 82:
							goto IL_0756;
						case 83:
							goto IL_0778;
						case 84:
							goto IL_079b;
						case 85:
							goto IL_07e3;
						case 86:
							goto IL_07ff;
						case 87:
							goto IL_0890;
						case 88:
							goto IL_08a0;
						case 90:
							goto IL_08d9;
						case 91:
							goto IL_08e6;
						case 92:
							goto IL_08f7;
						case 94:
							goto IL_092f;
						case 95:
							goto IL_0998;
						case 89:
						case 93:
						case 96:
							goto IL_099e;
						case 97:
							goto IL_09a4;
						case 98:
							goto IL_09d1;
						case 74:
						case 80:
						case 99:
							goto IL_09d6;
						case 100:
							goto IL_09d9;
						case 64:
						case 101:
							goto IL_09e2;
						case 102:
							goto IL_0a00;
						case 103:
							goto IL_0a07;
						case 104:
							goto IL_0a21;
						case 105:
							goto IL_0a5b;
						case 106:
							goto IL_0a88;
						case 107:
							goto IL_0ab7;
						case 108:
							goto IL_0abd;
						case 109:
							goto IL_0aea;
						case 110:
							goto IL_0af0;
						case 111:
							goto IL_0b99;
						case 112:
							goto IL_0bce;
						case 113:
							goto IL_0bd4;
						case 114:
							goto IL_0be8;
						case 115:
							goto IL_0c0b;
						case 117:
							goto IL_0c5e;
						case 118:
							goto IL_0c6a;
						case 119:
							goto IL_0c7b;
						case 121:
							goto IL_0cb1;
						case 120:
						case 122:
							goto IL_0d1e;
						case 123:
							goto IL_0d5c;
						case 124:
							goto IL_0d5f;
						case 125:
							goto IL_0d9d;
						case 126:
							goto IL_0dd5;
						case 127:
							goto IL_0de6;
						case 128:
							goto IL_0e25;
						case 129:
							goto IL_0e43;
						case 130:
							goto IL_0e53;
						case 131:
							goto IL_0e6e;
						case 132:
							goto IL_0e8d;
						case 11:
						case 116:
						case 133:
							goto IL_0e90;
						case 134:
							goto IL_0e9a;
						case 135:
							goto IL_0ea8;
						case 136:
							goto IL_0eb6;
						case 137:
							goto IL_0ec4;
						case 138:
							goto IL_0ec7;
						case 139:
							goto IL_0ed0;
						case 140:
							goto IL_0ed8;
						case 141:
							goto IL_0ee1;
						case 142:
							goto IL_0eea;
						case 143:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 144:
							goto end_IL_0000;
						}
						goto default;
						IL_03b7:
						num2 = 46;
						if (array[num5] > array[num6])
						{
							goto IL_03c6;
						}
						goto IL_0416;
						IL_03c6:
						num2 = 47;
						text = array2[num6];
						goto IL_03d0;
						IL_03d0:
						num2 = 48;
						array2[num6] = array2[num5];
						goto IL_03dd;
						IL_0024:
						num2 = 3;
						application = MH.A.Application;
						goto IL_0034;
						IL_0034:
						num2 = 4;
						num5 = 1;
						goto IL_0039;
						IL_0039:
						num2 = 5;
						pattern = VH.A(155686);
						goto IL_0049;
						IL_0049:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_0050;
						IL_0050:
						num2 = 7;
						application2 = application;
						goto IL_0056;
						IL_0056:
						num2 = 8;
						range = application2.ActiveCell;
						goto IL_0061;
						IL_0061:
						num2 = 9;
						if (Conversions.ToBoolean(Operators.OrObject(Operators.OrObject(!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), Operators.NotObject(range.HasFormula)), range.HasArray)))
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
							goto IL_00c3;
						}
						goto IL_00ed;
						IL_03dd:
						num2 = 49;
						array2[num5] = text;
						goto IL_03e7;
						IL_00c3:
						num2 = 10;
						Interaction.MsgBox(VH.A(155721), MsgBoxStyle.Exclamation, VH.A(40448));
						goto IL_0e90;
						IL_00ed:
						num2 = 12;
						worksheet = range.Worksheet;
						goto IL_00fb;
						IL_00fb:
						num2 = 13;
						application2.ScreenUpdating = false;
						goto IL_0106;
						IL_0106:
						num2 = 14;
						application2.EnableEvents = false;
						goto IL_0111;
						IL_0111:
						num2 = 15;
						application2.DisplayAlerts = false;
						goto IL_011c;
						IL_011c:
						application2 = null;
						goto IL_011f;
						IL_011f:
						num2 = 17;
						text2 = Strings.Mid(Conversions.ToString(range.Formula), 2);
						goto IL_013c;
						IL_013c:
						num2 = 18;
						matchCollection = Regex.Matches(text2, VH.A(155782));
						goto IL_0156;
						IL_0156:
						num2 = 19;
						array2 = new string[matchCollection.Count - 1 + 1];
						goto IL_016b;
						IL_016b:
						num2 = 20;
						array = new int[matchCollection.Count - 1 + 1];
						goto IL_0182;
						IL_0182:
						num2 = 21;
						enumerator = matchCollection.GetEnumerator();
						goto IL_0219;
						IL_0219:
						if (enumerator.MoveNext())
						{
							match = (Match)enumerator.Current;
							goto IL_01a5;
						}
						goto IL_0227;
						IL_0227:
						num2 = 29;
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
						goto IL_0249;
						IL_03f8:
						num2 = 51;
						array[num6] = array[num5];
						goto IL_0405;
						IL_0405:
						num2 = 52;
						array[num5] = Conversions.ToInteger(text);
						goto IL_0416;
						IL_0249:
						num2 = 30;
						matchCollection = Regex.Matches(text2, VH.A(4549));
						goto IL_0261;
						IL_0261:
						num2 = 31;
						num7 = num6 + matchCollection.Count - 1;
						goto IL_0274;
						IL_0274:
						num2 = 32;
						array2 = (string[])Utils.CopyArray(array2, new string[num7 + 1]);
						goto IL_0290;
						IL_0290:
						num2 = 33;
						array = (int[])Utils.CopyArray(array, new int[num7 + 1]);
						goto IL_02ac;
						IL_02ac:
						num2 = 34;
						enumerator2 = matchCollection.GetEnumerator();
						goto IL_0350;
						IL_0350:
						if (enumerator2.MoveNext())
						{
							match2 = (Match)enumerator2.Current;
							goto IL_02cf;
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
						goto IL_0368;
						IL_0416:
						num2 = 53;
						num6++;
						goto IL_041f;
						IL_0368:
						num2 = 42;
						if (enumerator2 is IDisposable)
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
							(enumerator2 as IDisposable).Dispose();
						}
						goto IL_038a;
						IL_02cf:
						num2 = 35;
						match3 = match2;
						goto IL_02d6;
						IL_02d6:
						num2 = 36;
						array2[num6] = match3.Groups[1].ToString();
						goto IL_02f4;
						IL_038a:
						num2 = 43;
						num7 = Information.UBound(array2);
						goto IL_0399;
						IL_0399:
						num2 = 44;
						num8 = num7;
						num5 = 0;
						goto IL_0438;
						IL_0438:
						if (num5 <= num8)
						{
							goto IL_03a8;
						}
						goto IL_0441;
						IL_0441:
						num2 = 55;
						matchCollection = Regex.Matches(text2, pattern);
						goto IL_0451;
						IL_0451:
						num2 = 56;
						array3 = Regex.Split(text2, pattern);
						goto IL_0461;
						IL_0461:
						num2 = 57;
						text2 = "";
						goto IL_046b;
						IL_046b:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0472;
						IL_0472:
						num2 = 59;
						num6 = 0;
						goto IL_0478;
						IL_0478:
						num2 = 60;
						num5 = 1;
						goto IL_047e;
						IL_047e:
						num2 = 61;
						range.get_Offset((object)1, (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_04b7;
						IL_04b7:
						num2 = 62;
						array4 = array3;
						num9 = 0;
						goto IL_09eb;
						IL_09eb:
						if (num9 < array4.Length)
						{
							text3 = array4[num9];
							goto IL_04cd;
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
						goto IL_0a00;
						IL_02f4:
						num2 = 37;
						array[num6] = match3.Index;
						goto IL_0305;
						IL_0a00:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0a07;
						IL_0a07:
						num2 = 103;
						range2 = range.get_Offset((object)1, (object)0);
						goto IL_0a21;
						IL_0a21:
						num2 = 104;
						range2.get_Offset((object)num5, (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0a5b;
						IL_0a5b:
						num2 = 105;
						range2.get_Offset((object)num5, (object)0).Formula = VH.A(48936) + text2;
						goto IL_0a88;
						IL_0a88:
						num2 = 106;
						if (Operators.ConditionalCompareObjectEqual(range2.get_Offset((object)num5, (object)0).Formula, "", TextCompare: false))
						{
							goto IL_0ab7;
						}
						goto IL_0abd;
						IL_0ab7:
						num2 = 107;
						flag = true;
						goto IL_0abd;
						IL_0abd:
						num2 = 108;
						if (Information.IsError(range2.get_Offset((object)num5, (object)0)) & !Information.IsError(range))
						{
							goto IL_0aea;
						}
						goto IL_0af0;
						IL_0aea:
						num2 = 109;
						flag = true;
						goto IL_0af0;
						IL_0af0:
						num2 = 110;
						range2.get_Offset((object)num5, (object)(-1)).Formula = VH.A(48936) + range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(48936) + range2.get_Offset((object)num5, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0b99;
						IL_0b99:
						num2 = 111;
						if (!Conversions.ToBoolean(range2.get_Offset((object)num5, (object)(-1)).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))))
						{
							goto IL_0bce;
						}
						goto IL_0bd4;
						IL_0bce:
						num2 = 112;
						flag = true;
						goto IL_0bd4;
						IL_0bd4:
						num2 = 113;
						if (flag)
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
							goto IL_0be8;
						}
						goto IL_0c5e;
						IL_030e:
						num2 = 39;
						text2 = Strings.Replace(text2, match3.ToString(), application.WorksheetFunction.Rept(VH.A(43203), match3.Length));
						goto IL_034a;
						IL_0be8:
						num2 = 114;
						Interaction.MsgBox(VH.A(155841), MsgBoxStyle.Exclamation, VH.A(40448));
						goto IL_0c0b;
						IL_0c0b:
						num2 = 115;
						((_Worksheet)worksheet).get_Range((object)range.get_Offset((object)1, (object)0), (object)range.get_Offset((object)(num5 + 1), (object)0)).EntireRow.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0e90;
						IL_0c5e:
						num2 = 117;
						range3 = Helpers.GetLabelCell(range);
						goto IL_0c6a;
						IL_0c6a:
						num2 = 118;
						if (range3 == null)
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
							goto IL_0c7b;
						}
						goto IL_0cb1;
						IL_034a:
						match3 = null;
						goto IL_034d;
						IL_0c7b:
						num2 = 119;
						range2.get_Offset((object)num5, (object)(-1)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(117068));
						goto IL_0d1e;
						IL_0cb1:
						num2 = 121;
						range2.get_Offset((object)num5, (object)(range3.Column - range2.Column)).Formula = VH.A(48936) + range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0d1e;
						IL_0d1e:
						num2 = 122;
						range2.get_Offset((object)(num5 + 1), (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0d5c;
						IL_0d5c:
						range2 = null;
						goto IL_0d5f;
						IL_0d5f:
						num2 = 124;
						AutoColor.Range(((_Worksheet)worksheet).get_Range((object)range.get_Offset((object)2, (object)0), (object)range.get_Offset((object)(num5 + 1), (object)0)));
						goto IL_0d9d;
						IL_0d9d:
						num2 = 125;
						range4 = ((_Worksheet)worksheet).get_Range((object)range.get_Offset((object)1, (object)(-1)), (object)range.get_Offset((object)(num5 + 2), (object)0));
						goto IL_0dd5;
						IL_0dd5:
						num2 = 126;
						range4.EntireRow.Select();
						goto IL_0de6;
						IL_0de6:
						num2 = 127;
						range4.Rows.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0e25;
						IL_0e25:
						num2 = 128;
						range5 = range.get_Offset((object)(num5 + 1), (object)0);
						goto IL_0e43;
						IL_0e43:
						num2 = 129;
						range5.Activate();
						goto IL_0e53;
						IL_0e53:
						num2 = 130;
						range5.NumberFormat = RuntimeHelpers.GetObjectValue(range.NumberFormat);
						goto IL_0e6e;
						IL_0e6e:
						num2 = 131;
						application.SendKeys(VH.A(49606), false);
						goto IL_0e8d;
						IL_0e8d:
						range5 = null;
						goto IL_0e90;
						IL_0e90:
						num2 = 133;
						application3 = application;
						goto IL_0e9a;
						IL_0e9a:
						num2 = 134;
						application3.ScreenUpdating = true;
						goto IL_0ea8;
						IL_0ea8:
						num2 = 135;
						application3.EnableEvents = true;
						goto IL_0eb6;
						IL_0eb6:
						num2 = 136;
						application3.DisplayAlerts = true;
						goto IL_0ec4;
						IL_0ec4:
						application3 = null;
						goto IL_0ec7;
						IL_0ec7:
						num2 = 138;
						application = null;
						goto IL_0ed0;
						IL_0ed0:
						num2 = 139;
						range4 = null;
						goto IL_0ed8;
						IL_0ed8:
						num2 = 140;
						range = null;
						goto IL_0ee1;
						IL_0ee1:
						num2 = 141;
						range3 = null;
						goto IL_0eea;
						IL_0eea:
						num2 = 142;
						worksheet = null;
						break;
						IL_0305:
						num2 = 38;
						num6++;
						goto IL_030e;
						IL_04cd:
						num2 = 63;
						if (Operators.CompareString(text3, "", TextCompare: false) != 0)
						{
							goto IL_04e4;
						}
						goto IL_09e2;
						IL_04e4:
						num2 = 65;
						range.get_Offset((object)1, (object)0).get_Offset((object)num5, (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0533;
						IL_0533:
						num2 = 66;
						text4 = "";
						goto IL_053d;
						IL_053d:
						num2 = 67;
						text4 = matchCollection[num5 - 1].ToString();
						goto IL_0556;
						IL_0556:
						num2 = 68;
						range6 = range.get_Offset((object)1, (object)0);
						goto IL_0570;
						IL_0570:
						num2 = 69;
						text2 = text2 + range6.get_Offset((object)num5, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + text4;
						goto IL_05c0;
						IL_05c0:
						num2 = 70;
						if (Versioned.IsNumeric(text3))
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
							goto IL_05db;
						}
						goto IL_0665;
						IL_034d:
						num2 = 41;
						goto IL_0350;
						IL_05db:
						num2 = 71;
						range6.get_Offset((object)num5, (object)0).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)text3);
						goto IL_0607;
						IL_0607:
						num2 = 72;
						range6.get_Offset((object)num5, (object)0).NumberFormat = "";
						goto IL_062a;
						IL_062a:
						num2 = 73;
						range6.get_Offset((object)num5, (object)(-1)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(155819));
						goto IL_09d6;
						IL_0665:
						num2 = 75;
						if (Regex.IsMatch(text3, VH.A(155830)))
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
							goto IL_068c;
						}
						goto IL_0723;
						IL_03e7:
						num2 = 50;
						text = Conversions.ToString(array[num6]);
						goto IL_03f8;
						IL_068c:
						num2 = 76;
						range6.get_Offset((object)num5, (object)0).Formula = VH.A(48936) + array2[num6];
						goto IL_06be;
						IL_06be:
						num2 = 77;
						num6++;
						goto IL_06c7;
						IL_06c7:
						num2 = 78;
						range6.get_Offset((object)num5, (object)0).NumberFormat = "";
						goto IL_06ea;
						IL_06ea:
						num2 = 79;
						range6.get_Offset((object)num5, (object)(-1)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(117068));
						goto IL_09d6;
						IL_0723:
						num2 = 81;
						range6.get_Offset((object)num5, (object)0).Formula = VH.A(48936) + text3;
						goto IL_0756;
						IL_0756:
						num2 = 82;
						text3 = Strings.Replace(text3, VH.A(39851), "");
						goto IL_0778;
						IL_0778:
						num2 = 83;
						if (Strings.InStr(text3, VH.A(7827)) == 0)
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
							goto IL_079b;
						}
						goto IL_07e3;
						IL_01a5:
						num2 = 22;
						match4 = match;
						goto IL_01ac;
						IL_079b:
						num2 = 84;
						text3 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), VH.A(7827)), text3));
						goto IL_07e3;
						IL_07e3:
						num2 = 85;
						array5 = Strings.Split(text3, VH.A(7827));
						goto IL_07ff;
						IL_07ff:
						num2 = 86;
						instance = application.ActiveWorkbook.Sheets[array5[0]];
						memberName = VH.A(41315);
						array6 = new object[1];
						reference = ref array5[1];
						array6[0] = reference;
						array7 = array6;
						obj = new bool[1] { true };
						array8 = obj;
						obj2 = NewLateBinding.LateGet(instance, null, memberName, array6, null, null, obj);
						if (array8[0])
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
							reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array7[0]), typeof(string));
						}
						range4 = (Range)obj2;
						goto IL_0890;
						IL_01bd:
						num2 = 24;
						array[num6] = match4.Index;
						goto IL_01ce;
						IL_01ce:
						num2 = 25;
						num6++;
						goto IL_01d7;
						IL_01ac:
						num2 = 23;
						array2[num6] = match4.ToString();
						goto IL_01bd;
						IL_0890:
						num2 = 87;
						if (range4 == null)
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
							goto IL_08a0;
						}
						goto IL_08d9;
						IL_01d7:
						num2 = 26;
						text2 = Strings.Replace(text2, match4.ToString(), application.WorksheetFunction.Rept(VH.A(43203), match4.Length));
						goto IL_0213;
						IL_08a0:
						num2 = 88;
						range6.get_Offset((object)num5, (object)(-1)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(117068));
						goto IL_099e;
						IL_08d9:
						num2 = 90;
						range3 = Helpers.GetLabelCell(range4);
						goto IL_08e6;
						IL_08e6:
						num2 = 91;
						if (range3 == null)
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
							goto IL_08f7;
						}
						goto IL_092f;
						IL_0213:
						match4 = null;
						goto IL_0216;
						IL_08f7:
						num2 = 92;
						range6.get_Offset((object)num5, (object)(-1)).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(117068));
						goto IL_099e;
						IL_092f:
						num2 = 94;
						range6.get_Offset((object)num5, (object)(range3.Column - range6.Column)).Formula = VH.A(48936) + range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0998;
						IL_0998:
						num2 = 95;
						range3 = null;
						goto IL_099e;
						IL_099e:
						num2 = 96;
						if (range4 != null)
						{
							goto IL_09a4;
						}
						goto IL_09d6;
						IL_09a4:
						num2 = 97;
						range6.get_Offset((object)num5, (object)0).NumberFormat = RuntimeHelpers.GetObjectValue(range4.NumberFormat);
						goto IL_09d1;
						IL_09d1:
						num2 = 98;
						range4 = null;
						goto IL_09d6;
						IL_09d6:
						range6 = null;
						goto IL_09d9;
						IL_09d9:
						num2 = 100;
						num5++;
						goto IL_09e2;
						IL_09e2:
						num2 = 101;
						num9++;
						goto IL_09eb;
						IL_03a8:
						num2 = 45;
						num10 = num5 + 1;
						num11 = num7;
						num6 = num10;
						goto IL_041f;
						IL_041f:
						if (num6 <= num11)
						{
							goto IL_03b7;
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
						goto IL_042f;
						IL_0216:
						num2 = 28;
						goto IL_0219;
						IL_042f:
						num2 = 54;
						num5++;
						goto IL_0438;
						end_IL_0000_3:
						break;
					}
					num2 = 143;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(155898));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 4449;
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
}
