using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class ContributionAnalysis
{
	public static void Go()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		string[] array = default(string[]);
		int num5 = default(int);
		Worksheet worksheet = default(Worksheet);
		string[] array2 = default(string[]);
		int num6 = default(int);
		Application application = default(Application);
		Regex regex = default(Regex);
		Range range = default(Range);
		Range range2 = default(Range);
		Range range3 = default(Range);
		string[] array3 = default(string[]);
		int num7 = default(int);
		int num8 = default(int);
		int num9 = default(int);
		int num10 = default(int);
		Range range4 = default(Range);
		object instance = default(object);
		int column = default(int);
		Range range5 = default(Range);
		Range range6 = default(Range);
		Range range7 = default(Range);
		Range range8 = default(Range);
		Range range9 = default(Range);
		Range range10 = default(Range);
		object instance2 = default(object);
		int num12 = default(int);
		int num13 = default(int);
		string text = default(string);
		int num14 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		Range range11 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					int num11;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						if (!Access.AllowExcelOperation((PlanType)4, (Restriction)2, false))
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
						goto IL_0024;
					case 3536:
						{
							num = num2;
							switch (num3)
							{
							case 2:
							case 3:
								break;
							case 1:
								goto IL_0c0c;
							default:
								goto end_IL_0000_2;
							}
							goto IL_0ba8;
						}
						IL_0c0c:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0024;
						case 4:
							goto IL_002b;
						case 5:
							goto IL_003c;
						case 6:
							goto IL_0052;
						case 7:
							goto IL_005a;
						case 8:
							goto IL_0078;
						case 9:
							goto IL_007f;
						case 10:
							goto IL_0089;
						case 11:
							goto IL_0093;
						case 12:
							goto IL_00a5;
						case 13:
							goto IL_00ce;
						case 15:
							goto IL_00e5;
						case 16:
							goto IL_0111;
						case 17:
							goto IL_011e;
						case 18:
							goto IL_012c;
						case 19:
							goto IL_0137;
						case 20:
							goto IL_013d;
						case 21:
							goto IL_0167;
						case 22:
							goto IL_017e;
						case 23:
							goto IL_01a6;
						case 24:
							goto IL_01c4;
						case 25:
							goto IL_01fe;
						case 26:
							goto IL_0215;
						case 27:
							goto IL_0237;
						case 28:
							goto IL_0272;
						case 29:
							goto IL_028e;
						case 30:
							goto IL_02be;
						case 31:
							goto IL_02c7;
						case 32:
							goto IL_02dd;
						case 34:
							goto IL_0304;
						case 35:
							goto IL_0320;
						case 36:
							goto IL_036c;
						case 33:
						case 37:
							goto IL_0375;
						case 39:
							goto IL_0396;
						case 40:
							goto IL_03b3;
						case 41:
							goto IL_03cf;
						case 42:
							goto IL_0419;
						case 38:
						case 43:
							goto IL_0422;
						case 44:
							goto IL_043e;
						case 45:
							goto IL_0449;
						case 46:
							goto IL_0458;
						case 47:
							goto IL_0467;
						case 48:
							goto IL_04a0;
						case 50:
							goto IL_04b9;
						case 51:
							goto IL_04cb;
						case 52:
							goto IL_0520;
						case 53:
							goto IL_052a;
						case 54:
							goto IL_0537;
						case 55:
							goto IL_0541;
						case 56:
							goto IL_055d;
						case 57:
							goto IL_0579;
						case 58:
							goto IL_05d5;
						case 59:
							goto IL_05e4;
						case 60:
							goto IL_0601;
						case 61:
							goto IL_0622;
						case 62:
							goto IL_06a3;
						case 63:
							goto IL_06b9;
						case 64:
							goto IL_06bc;
						case 65:
							goto IL_06c3;
						case 66:
							goto IL_06ca;
						case 67:
							goto IL_0717;
						case 68:
							goto IL_072b;
						case 69:
							goto IL_0754;
						case 70:
							goto IL_0762;
						case 71:
							goto IL_07a9;
						case 72:
							goto IL_080c;
						case 73:
							goto IL_0834;
						case 74:
							goto IL_0837;
						case 75:
							goto IL_083a;
						case 76:
							goto IL_0840;
						case 77:
							goto IL_085c;
						case 78:
							goto IL_0863;
						case 79:
							goto IL_0886;
						case 80:
							goto IL_0953;
						case 81:
							goto IL_097b;
						case 82:
							goto IL_09a9;
						case 83:
							goto IL_09d3;
						case 84:
							goto IL_09d6;
						case 85:
							goto IL_09d9;
						case 86:
							goto IL_09ff;
						case 87:
							goto IL_0a2b;
						case 88:
							goto IL_0a77;
						case 89:
							goto IL_0a84;
						case 90:
							goto IL_0a91;
						case 91:
							goto IL_0a94;
						case 92:
							goto IL_0b0a;
						case 93:
							goto IL_0b14;
						case 94:
							goto IL_0b21;
						case 95:
							goto IL_0b2e;
						case 96:
							goto IL_0b3b;
						case 14:
						case 49:
						case 97:
							goto IL_0ba8;
						case 98:
							goto IL_0bb2;
						case 99:
							goto IL_0bbc;
						case 100:
							goto IL_0bc2;
						case 101:
							goto IL_0bc8;
						case 102:
							goto IL_0bce;
						case 103:
							goto IL_0bd4;
						case 104:
							goto IL_0bda;
						case 105:
							goto IL_0be0;
						case 106:
							goto IL_0be6;
						case 107:
							goto IL_0beb;
						case 108:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 109:
							goto end_IL_0000;
						}
						goto default;
						IL_02c7:
						num2 = 31;
						goto IL_02ca;
						IL_0304:
						num2 = 34;
						array = (string[])Utils.CopyArray(array, new string[num5 + 1]);
						goto IL_0320;
						IL_0320:
						num2 = 35;
						array[num5] = ((_Worksheet)worksheet).get_Range((object)array2[num6], RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_036c;
						IL_0024:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_002b;
						IL_002b:
						num2 = 4;
						application = MH.A.Application;
						goto IL_003c;
						IL_003c:
						num2 = 5;
						regex = new Regex(VH.A(90056), RegexOptions.IgnoreCase);
						goto IL_0052;
						IL_0052:
						num2 = 6;
						Common.ClearVisualizations(application);
						goto IL_005a;
						IL_005a:
						num2 = 7;
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
							goto IL_0078;
						}
						goto IL_0bbc;
						IL_036c:
						num2 = 36;
						num5++;
						goto IL_0375;
						IL_0078:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_007f;
						IL_007f:
						num2 = 9;
						application.ScreenUpdating = false;
						goto IL_0089;
						IL_0089:
						num2 = 10;
						application.EnableEvents = false;
						goto IL_0093;
						IL_0093:
						num2 = 11;
						range = (Range)application.Selection;
						goto IL_00a5;
						IL_00a5:
						num2 = 12;
						if (Operators.ConditionalCompareObjectGreater(range.Rows.CountLarge, 1, TextCompare: false))
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
							goto IL_00ce;
						}
						goto IL_00e5;
						IL_0375:
						num2 = 37;
						num6++;
						goto IL_037e;
						IL_00ce:
						num2 = 13;
						Forms.WarningMessage(VH.A(90081));
						goto IL_0ba8;
						IL_00e5:
						num2 = 15;
						range2 = (Range)range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0111;
						IL_0111:
						num2 = 16;
						range3 = application.ActiveCell;
						goto IL_011e;
						IL_011e:
						num2 = 17;
						worksheet = range.Worksheet;
						goto IL_012c;
						IL_012c:
						num2 = 18;
						array = new string[1];
						goto IL_0137;
						IL_0137:
						num2 = 19;
						num5 = 0;
						goto IL_013d;
						IL_013d:
						num2 = 20;
						array3 = Strings.Split(Conversions.ToString(range2.Formula), VH.A(54459));
						goto IL_0167;
						IL_0167:
						num2 = 21;
						num7 = Information.UBound(array3);
						num8 = 0;
						goto IL_042b;
						IL_042b:
						if (num8 <= num7)
						{
							goto IL_017e;
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
						goto IL_043e;
						IL_0396:
						num2 = 39;
						if (Regex.IsMatch(array3[num8], VH.A(90192)))
						{
							goto IL_03b3;
						}
						goto IL_0422;
						IL_043e:
						num2 = 44;
						if (num5 > 0)
						{
							goto IL_0449;
						}
						goto IL_0ba8;
						IL_0449:
						num2 = 45;
						num8 = Information.UBound(array);
						goto IL_0458;
						IL_0458:
						num2 = 46;
						num9 = num8;
						num6 = 0;
						goto IL_0566;
						IL_0566:
						if (num6 <= num9)
						{
							goto IL_0467;
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
						goto IL_0579;
						IL_03b3:
						num2 = 40;
						array = (string[])Utils.CopyArray(array, new string[num5 + 1]);
						goto IL_03cf;
						IL_0579:
						num2 = 57;
						((_Worksheet)worksheet).get_Range((object)range2.get_Offset((object)1, (object)0), (object)range2.get_Offset((object)(num5 + 3), (object)0)).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_05d5;
						IL_05d5:
						num2 = 58;
						num10 = num8;
						num6 = 0;
						goto IL_0849;
						IL_0849:
						if (num6 <= num10)
						{
							goto IL_05e4;
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
						goto IL_085c;
						IL_03cf:
						num2 = 41;
						array[num5] = ((_Worksheet)worksheet).get_Range((object)array3[num8], RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0419;
						IL_085c:
						ProjectData.ClearProjectError();
						num3 = 3;
						goto IL_0863;
						IL_0863:
						num2 = 78;
						range4 = range.get_Offset((object)(num5 + 3), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0886;
						IL_0886:
						num2 = 79;
						range4.Formula = VH.A(54448) + range2.get_Offset((object)2, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2826) + range2.get_Offset((object)(num5 + 2), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
						goto IL_0953;
						IL_0953:
						num2 = 80;
						instance = worksheet.Cells[range4.Row, column];
						goto IL_097b;
						IL_097b:
						num2 = 81;
						NewLateBinding.LateSetComplex(instance, null, VH.A(41636), new object[1] { VH.A(57315) }, null, null, OptimisticSet: false, RValueBase: true);
						goto IL_09a9;
						IL_09a9:
						num2 = 82;
						NewLateBinding.LateSetComplex(instance, null, VH.A(90455), new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
						goto IL_09d3;
						IL_09d3:
						instance = null;
						goto IL_09d6;
						IL_09d6:
						range4 = null;
						goto IL_09d9;
						IL_09d9:
						num2 = 85;
						SumBar.Toggle(range2.get_Offset((object)(num5 + 2), RuntimeHelpers.GetObjectValue(Missing.Value)));
						goto IL_09ff;
						IL_09ff:
						num2 = 86;
						if (Operators.ConditionalCompareObjectGreater(range.Columns.CountLarge, 1, TextCompare: false))
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
							goto IL_0a2b;
						}
						goto IL_0b3b;
						IL_0419:
						num2 = 42;
						num5++;
						goto IL_0422;
						IL_0a2b:
						num2 = 87;
						range5 = ((_Worksheet)worksheet).get_Range((object)range2.get_Offset((object)2, (object)0), (object)range2.get_Offset((object)(num5 + 3), Operators.SubtractObject(range.Columns.CountLarge, 1)));
						goto IL_0a77;
						IL_0a77:
						num2 = 88;
						range5.Select();
						goto IL_0a84;
						IL_0a84:
						num2 = 89;
						range5.FillRight();
						goto IL_0a91;
						IL_0a91:
						range5 = null;
						goto IL_0a94;
						IL_0a94:
						num2 = 91;
						range6 = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[range2.get_Offset((object)2, (object)0).Row, column]), (object)range2.get_Offset((object)(num5 + 3), Operators.SubtractObject(range.Columns.CountLarge, 1)));
						goto IL_0b0a;
						IL_0b0a:
						num2 = 92;
						AutoColor.Range(range6);
						goto IL_0b14;
						IL_0b14:
						num2 = 93;
						range6.Select();
						goto IL_0b21;
						IL_0b21:
						num2 = 94;
						range.Select();
						goto IL_0b2e;
						IL_0b2e:
						num2 = 95;
						range3.Activate();
						goto IL_0b3b;
						IL_0b3b:
						num2 = 96;
						range2.get_Offset((object)1, (object)0).get_Resize((object)(num5 + 3), RuntimeHelpers.GetObjectValue(Missing.Value)).Rows.Group(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0ba8;
						IL_05e4:
						num2 = 59;
						range7 = ((_Worksheet)worksheet).get_Range((object)array[num6], RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0601;
						IL_0601:
						num2 = 60;
						range8 = range2.get_Offset((object)(num6 + 2), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0622;
						IL_0622:
						num2 = 61;
						range8.Formula = VH.A(48936) + range7.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75498) + range2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_06a3;
						IL_06a3:
						num2 = 62;
						range8.NumberFormat = VH.A(90396);
						goto IL_06b9;
						IL_06b9:
						range8 = null;
						goto IL_06bc;
						IL_06bc:
						num2 = 64;
						range9 = range7;
						goto IL_06c3;
						IL_06c3:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_06ca;
						IL_06ca:
						num2 = 66;
						range6 = ((_Application)range9.Application).get_Range(RuntimeHelpers.GetObjectValue(range9.Worksheet.Cells[range9.Row, 1]), (object)range7).SpecialCells(XlCellType.xlCellTypeConstants, XlSpecialCellsValue.xlTextValues);
						goto IL_0717;
						IL_0717:
						num2 = 67;
						if (range6 != null)
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
							goto IL_072b;
						}
						goto IL_0837;
						IL_0422:
						num2 = 43;
						num8++;
						goto IL_042b;
						IL_072b:
						num2 = 68;
						range10 = (Range)range6.get_Item((object)range6.Count, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0754;
						IL_0754:
						num2 = 69;
						column = range10.Column;
						goto IL_0762;
						IL_0762:
						num2 = 70;
						instance2 = range9.Worksheet.Cells[range2.get_Offset((object)(num6 + 2), RuntimeHelpers.GetObjectValue(Missing.Value)).Row, column];
						goto IL_07a9;
						IL_07a9:
						num2 = 71;
						NewLateBinding.LateSetComplex(instance2, null, VH.A(68956), new object[1] { VH.A(48936) + range10.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) }, null, null, OptimisticSet: false, RValueBase: true);
						goto IL_080c;
						IL_080c:
						num2 = 72;
						NewLateBinding.LateSetComplex(instance2, null, VH.A(90455), new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
						goto IL_0834;
						IL_0834:
						instance2 = null;
						goto IL_0837;
						IL_0837:
						range9 = null;
						goto IL_083a;
						IL_083a:
						num2 = 75;
						range6 = null;
						goto IL_0840;
						IL_0840:
						num2 = 76;
						num6++;
						goto IL_0849;
						IL_0467:
						num2 = 47;
						if (((_Worksheet)worksheet).get_Range((object)array[num6], RuntimeHelpers.GetObjectValue(Missing.Value)).Column != range2.Column)
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
							goto IL_04a0;
						}
						goto IL_04b9;
						IL_0bbc:
						num2 = 99;
						range3 = null;
						goto IL_0bc2;
						IL_04a0:
						num2 = 48;
						Forms.WarningMessage(VH.A(90249));
						goto IL_0ba8;
						IL_04b9:
						num2 = 50;
						num11 = num6 + 1;
						num12 = num8;
						num13 = num11;
						goto IL_054a;
						IL_054a:
						if (num13 <= num12)
						{
							goto IL_04cb;
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
						goto IL_055d;
						IL_0bc2:
						num2 = 100;
						range = null;
						goto IL_0bc8;
						IL_055d:
						num2 = 56;
						num6++;
						goto IL_0566;
						IL_04cb:
						num2 = 51;
						if (Conversions.ToInteger(Regex.Replace(array[num6], VH.A(90385), "")) > Conversions.ToInteger(Regex.Replace(array[num13], VH.A(90385), "")))
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
							goto IL_0520;
						}
						goto IL_0541;
						IL_0bc8:
						num2 = 101;
						range2 = null;
						goto IL_0bce;
						IL_0520:
						num2 = 52;
						text = array[num13];
						goto IL_052a;
						IL_052a:
						num2 = 53;
						array[num13] = array[num6];
						goto IL_0537;
						IL_0537:
						num2 = 54;
						array[num6] = text;
						goto IL_0541;
						IL_0541:
						num2 = 55;
						num13++;
						goto IL_054a;
						IL_0ba8:
						num2 = 97;
						application.EnableEvents = true;
						goto IL_0bb2;
						IL_0bb2:
						num2 = 98;
						application.ScreenUpdating = true;
						goto IL_0bbc;
						IL_017e:
						num2 = 22;
						array3[num8] = Strings.Replace(array3[num8], VH.A(48936), "");
						goto IL_01a6;
						IL_01a6:
						num2 = 23;
						if (regex.IsMatch(array3[num8]))
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
							goto IL_01c4;
						}
						goto IL_0396;
						IL_0bce:
						num2 = 102;
						range7 = null;
						goto IL_0bd4;
						IL_01c4:
						num2 = 24;
						array2 = Strings.Split(regex.Match(array3[num8]).Groups[1].ToString(), VH.A(2378));
						goto IL_01fe;
						IL_01fe:
						num2 = 25;
						num14 = Information.UBound(array2);
						num6 = 0;
						goto IL_037e;
						IL_037e:
						if (num6 <= num14)
						{
							goto IL_0215;
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
						goto IL_0422;
						IL_0bd4:
						num2 = 103;
						range10 = null;
						goto IL_0bda;
						IL_0215:
						num2 = 26;
						if (Strings.InStr(array2[num6], VH.A(2826)) > 0)
						{
							goto IL_0237;
						}
						goto IL_0304;
						IL_0237:
						num2 = 27;
						enumerator = ((_Worksheet)worksheet).get_Range((object)array2[num6], RuntimeHelpers.GetObjectValue(Missing.Value)).Cells.GetEnumerator();
						goto IL_02ca;
						IL_02ca:
						if (enumerator.MoveNext())
						{
							range11 = (Range)enumerator.Current;
							goto IL_0272;
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
						goto IL_02dd;
						IL_0bda:
						num2 = 104;
						range6 = null;
						goto IL_0be0;
						IL_02dd:
						num2 = 32;
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
						goto IL_0375;
						IL_0be6:
						num2 = 106;
						application = null;
						goto IL_0beb;
						IL_0beb:
						num2 = 107;
						regex = null;
						break;
						IL_0be0:
						num2 = 105;
						worksheet = null;
						goto IL_0be6;
						IL_0272:
						num2 = 28;
						array = (string[])Utils.CopyArray(array, new string[num5 + 1]);
						goto IL_028e;
						IL_028e:
						num2 = 29;
						array[num5] = range11.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_02be;
						IL_02be:
						num2 = 30;
						num5++;
						goto IL_02c7;
						end_IL_0000_3:
						break;
					}
					num2 = 108;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(90478));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 3536;
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
}
