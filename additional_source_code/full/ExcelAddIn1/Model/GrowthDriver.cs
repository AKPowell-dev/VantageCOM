using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Format;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class GrowthDriver
{
	public static void Add()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Application application = default(Application);
		Range range = default(Range);
		int num5 = default(int);
		Range range2 = default(Range);
		Range range3 = default(Range);
		Range range4 = default(Range);
		Range range5 = default(Range);
		Range range6 = default(Range);
		int num6 = default(int);
		double num7 = default(double);
		wpfGrowthDriver wpfGrowthDriver2 = default(wpfGrowthDriver);
		List<string> list = default(List<string>);
		IEnumerator enumerator = default(IEnumerator);
		List<string> list2 = default(List<string>);
		List<NumberFormatCycle.NumberFormat>.Enumerator enumerator2 = default(List<NumberFormatCycle.NumberFormat>.Enumerator);
		NumberFormatCycle.NumberFormat current = default(NumberFormatCycle.NumberFormat);
		string text = default(string);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					bool? flag;
					bool? isChecked;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
						{
							goto end_IL_0000;
						}
						goto IL_0011;
					case 3779:
						{
							num = num2;
							switch (num3)
							{
							case 2:
								break;
							case 1:
								goto IL_0ccf;
							default:
								goto end_IL_0000_2;
							}
							goto IL_0c7d;
						}
						IL_0ccf:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0011;
						case 4:
							goto IL_0019;
						case 5:
							goto IL_0020;
						case 6:
							goto IL_002f;
						case 7:
							goto IL_0056;
						case 8:
							goto IL_0067;
						case 9:
							goto IL_0071;
						case 10:
							goto IL_0092;
						case 12:
							goto IL_00ab;
						case 13:
							goto IL_00dc;
						case 14:
							goto IL_010c;
						case 15:
							goto IL_0112;
						case 16:
							goto IL_0122;
						case 17:
							goto IL_0136;
						case 18:
							goto IL_013c;
						case 19:
							goto IL_0176;
						case 20:
							goto IL_018f;
						case 22:
							goto IL_019a;
						case 23:
							goto IL_01e0;
						case 24:
							goto IL_01fd;
						case 21:
						case 25:
							goto IL_0204;
						case 26:
							goto IL_0215;
						case 27:
							goto IL_022a;
						case 28:
							goto IL_0245;
						case 30:
							goto IL_0264;
						case 32:
							goto IL_0282;
						case 29:
						case 31:
						case 33:
							goto IL_0289;
						case 34:
							goto IL_02b4;
						case 37:
							goto IL_02cb;
						case 38:
							goto IL_02d8;
						case 39:
							goto IL_02de;
						case 40:
							goto IL_0429;
						case 41:
							goto IL_043a;
						case 43:
							goto IL_0448;
						case 36:
						case 42:
						case 44:
							goto IL_0453;
						case 45:
							goto IL_04dd;
						case 46:
							goto IL_04eb;
						case 47:
							goto IL_04f1;
						case 48:
							goto IL_04f7;
						case 49:
							goto IL_0501;
						case 50:
							goto IL_050b;
						case 51:
							goto IL_052e;
						case 52:
							goto IL_0537;
						case 54:
							goto IL_0542;
						case 55:
							goto IL_0587;
						case 56:
							goto IL_05a2;
						case 57:
							goto IL_05c1;
						case 53:
						case 58:
							goto IL_05ca;
						case 59:
							goto IL_05e5;
						case 60:
							goto IL_0607;
						case 61:
							goto IL_0618;
						case 62:
							goto IL_0629;
						case 63:
							goto IL_062f;
						case 64:
							goto IL_0639;
						case 65:
							goto IL_0663;
						case 66:
							goto IL_0674;
						case 67:
							goto IL_0682;
						case 68:
							goto IL_0692;
						case 69:
							goto IL_06a5;
						case 70:
							goto IL_06b7;
						case 71:
							goto IL_06bd;
						case 72:
							goto IL_0703;
						case 73:
							goto IL_0715;
						case 74:
							goto IL_072c;
						case 75:
							goto IL_0737;
						case 77:
							goto IL_0774;
						case 78:
							goto IL_07a9;
						case 79:
							goto IL_07dd;
						case 80:
							goto IL_07e7;
						case 81:
							goto IL_07f1;
						case 82:
							goto IL_0846;
						case 83:
							goto IL_0881;
						case 84:
							goto IL_08b6;
						case 85:
							goto IL_08fe;
						case 86:
							goto IL_091a;
						case 87:
							goto IL_093d;
						case 88:
							goto IL_0940;
						case 89:
							goto IL_0946;
						case 90:
							goto IL_094d;
						case 91:
							goto IL_0970;
						case 93:
							goto IL_0999;
						case 94:
							goto IL_09a2;
						case 95:
							goto IL_09b7;
						case 96:
							goto IL_09be;
						case 97:
							goto IL_09d6;
						case 99:
							goto IL_0a90;
						case 100:
							goto IL_0ad8;
						case 101:
							goto IL_0afc;
						case 98:
						case 102:
							goto IL_0baf;
						case 103:
							goto IL_0bb2;
						case 92:
						case 104:
							goto IL_0bb9;
						case 105:
							goto IL_0bca;
						case 106:
							goto IL_0be2;
						case 107:
							goto IL_0be9;
						case 108:
							goto IL_0bfb;
						case 109:
							goto IL_0c22;
						case 110:
							goto IL_0c45;
						case 111:
							goto IL_0c5b;
						case 11:
						case 35:
						case 76:
						case 112:
							goto IL_0c7d;
						case 113:
							goto IL_0c87;
						case 114:
							goto IL_0c91;
						case 115:
							goto IL_0c97;
						case 116:
							goto IL_0c9d;
						case 117:
							goto IL_0ca3;
						case 118:
							goto IL_0ca9;
						case 119:
							goto IL_0caf;
						case 120:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 121:
							goto end_IL_0000;
						}
						goto default;
						IL_0011:
						num2 = 3;
						Common.ClearVisualizations();
						goto IL_0019;
						IL_0019:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0020;
						IL_0020:
						num2 = 5;
						application = MH.A.Application;
						goto IL_002f;
						IL_002f:
						num2 = 6;
						if (application.Selection is Range)
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
							goto IL_0056;
						}
						goto IL_0c91;
						IL_0ca3:
						num2 = 117;
						range = null;
						goto IL_0ca9;
						IL_0999:
						num2 = 93;
						num5++;
						goto IL_09a2;
						IL_09a2:
						num2 = 94;
						if (num5 > 1)
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
							goto IL_09b7;
						}
						goto IL_0bb2;
						IL_0056:
						num2 = 7;
						range2 = (Range)application.Selection;
						goto IL_0067;
						IL_0067:
						num2 = 8;
						range3 = application.ActiveCell;
						goto IL_0071;
						IL_0071:
						num2 = 9;
						if (Operators.ConditionalCompareObjectGreater(range2.Rows.CountLarge, 1, TextCompare: false))
						{
							goto IL_0092;
						}
						goto IL_00ab;
						IL_0092:
						num2 = 10;
						Forms.WarningMessage(VH.A(94625));
						goto IL_0c7d;
						IL_00ab:
						num2 = 12;
						range4 = Helpers.GetLabelCell((Range)range2.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
						goto IL_00dc;
						IL_00dc:
						num2 = 13;
						if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
							goto IL_010c;
						}
						goto IL_0453;
						IL_0ca9:
						num2 = 118;
						range5 = null;
						goto IL_0caf;
						IL_010c:
						num2 = 14;
						range5 = null;
						goto IL_0112;
						IL_0112:
						num2 = 15;
						range5 = A(range2, range4);
						goto IL_0122;
						IL_0122:
						num2 = 16;
						if (range5 == null)
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
							goto IL_0136;
						}
						goto IL_0282;
						IL_09b7:
						num2 = 95;
						range6 = range5;
						goto IL_09be;
						IL_0136:
						num2 = 17;
						num5 = 1;
						goto IL_013c;
						IL_013c:
						num2 = 18;
						if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.get_Offset((object)num5, (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
						{
							goto IL_0176;
						}
						goto IL_019a;
						IL_0176:
						num2 = 19;
						range5 = range2.get_Offset((object)num5, (object)0);
						goto IL_018f;
						IL_018f:
						num2 = 20;
						num6 = -num5;
						goto IL_0204;
						IL_019a:
						num2 = 22;
						if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range2.get_Offset((object)(-num5), (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
							goto IL_01e0;
						}
						goto IL_0204;
						IL_09be:
						num2 = 96;
						if (range6.Column < range3.Column)
						{
							goto IL_09d6;
						}
						goto IL_0a90;
						IL_01e0:
						num2 = 23;
						range5 = range2.get_Offset((object)(-num5), (object)0);
						goto IL_01fd;
						IL_01fd:
						num2 = 24;
						num6 = num5;
						goto IL_0204;
						IL_0204:
						num2 = 25;
						if (range5 != null)
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
							goto IL_0215;
						}
						goto IL_0264;
						IL_09d6:
						num2 = 97;
						range6.get_Offset((object)1, (object)0).Formula = VH.A(48936) + range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(75498) + range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(94843);
						goto IL_0baf;
						IL_0215:
						num2 = 26;
						range5 = RangeHelpers.CellsWithNumbers(range5.EntireRow);
						goto IL_022a;
						IL_022a:
						num2 = 27;
						range2 = range5.get_Offset((object)num6, (object)0);
						goto IL_0245;
						IL_0245:
						num2 = 28;
						range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)100);
						goto IL_0289;
						IL_0264:
						num2 = 30;
						num5++;
						if (num5 <= 100)
						{
							goto IL_013c;
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
						goto IL_0289;
						IL_0a90:
						num2 = 99;
						num7 = Conversions.ToDouble(Operators.SubtractObject(Operators.DivideObject(range6.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)), range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value))), 1));
						goto IL_0ad8;
						IL_0282:
						num2 = 32;
						range2 = range5;
						goto IL_0289;
						IL_0289:
						num2 = 33;
						if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
							goto IL_02b4;
						}
						goto IL_02cb;
						IL_0ad8:
						num2 = 100;
						range6.get_Offset((object)1, (object)0).Formula = num7;
						goto IL_0afc;
						IL_02b4:
						num2 = 34;
						Forms.WarningMessage(VH.A(94700));
						goto IL_0c7d;
						IL_02cb:
						num2 = 37;
						range2.Select();
						goto IL_02d8;
						IL_02d8:
						num2 = 38;
						range5 = null;
						goto IL_02de;
						IL_02de:
						num2 = 39;
						range5 = application.Intersect(range3, range2, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0429;
						IL_0429:
						num2 = 40;
						if (range5 == null)
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
							goto IL_043a;
						}
						goto IL_0448;
						IL_0afc:
						num2 = 101;
						range6.Formula = VH.A(48936) + range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(94848) + range6.get_Offset((object)1, (object)0).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(39904);
						goto IL_0baf;
						IL_043a:
						num2 = 41;
						range3 = A(range2);
						goto IL_0453;
						IL_0448:
						num2 = 43;
						range3.Activate();
						goto IL_0453;
						IL_0453:
						num2 = 44;
						if (Operators.ConditionalCompareObjectEqual(range3.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), NewLateBinding.LateGet(range2.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[0], null, null, null), TextCompare: false))
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
							goto IL_04dd;
						}
						goto IL_04eb;
						IL_0baf:
						range6 = null;
						goto IL_0bb2;
						IL_04dd:
						num2 = 45;
						range3 = A(range2);
						goto IL_04eb;
						IL_04eb:
						num2 = 46;
						num5 = 0;
						goto IL_04f1;
						IL_04f1:
						num2 = 47;
						num6 = 0;
						goto IL_04f7;
						IL_04f7:
						num2 = 48;
						wpfGrowthDriver2 = new wpfGrowthDriver();
						goto IL_0501;
						IL_0501:
						num2 = 49;
						list = new List<string>();
						goto IL_050b;
						IL_050b:
						num2 = 50;
						enumerator = range2.GetEnumerator();
						goto IL_05cd;
						IL_05cd:
						if (enumerator.MoveNext())
						{
							range5 = (Range)enumerator.Current;
							goto IL_052e;
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
						goto IL_05e5;
						IL_0bb2:
						num2 = 103;
						range = range5;
						goto IL_0bb9;
						IL_05e5:
						num2 = 59;
						if (enumerator is IDisposable)
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
							(enumerator as IDisposable).Dispose();
						}
						goto IL_0607;
						IL_0c7d:
						num2 = 112;
						application.EnableEvents = true;
						goto IL_0c87;
						IL_0c87:
						num2 = 113;
						application.ScreenUpdating = true;
						goto IL_0c91;
						IL_0607:
						num2 = 60;
						wpfGrowthDriver2.cbxColumn.ItemsSource = list;
						goto IL_0618;
						IL_0618:
						num2 = 61;
						wpfGrowthDriver2.cbxColumn.SelectedIndex = num6;
						goto IL_0629;
						IL_0629:
						num2 = 62;
						list = null;
						goto IL_062f;
						IL_062f:
						num2 = 63;
						list2 = new List<string>();
						goto IL_0639;
						IL_0639:
						num2 = 64;
						enumerator2 = KH.A.CyclePercent.Items.GetEnumerator();
						goto IL_0677;
						IL_0677:
						if (enumerator2.MoveNext())
						{
							current = enumerator2.Current;
							goto IL_0663;
						}
						goto IL_0682;
						IL_0682:
						num2 = 67;
						((IDisposable)enumerator2/*cast due to .constrained prefix*/).Dispose();
						goto IL_0692;
						IL_0692:
						num2 = 68;
						wpfGrowthDriver2.cbxNumFormat.ItemsSource = list2;
						goto IL_06a5;
						IL_06a5:
						num2 = 69;
						wpfGrowthDriver2.cbxNumFormat.SelectedIndex = 0;
						goto IL_06b7;
						IL_06b7:
						num2 = 70;
						list2 = null;
						goto IL_06bd;
						IL_06bd:
						num2 = 71;
						if (Regex.IsMatch(Conversions.ToString(range4.get_Offset((object)1, (object)0).Text), VH.A(94830), RegexOptions.IgnoreCase))
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
							goto IL_0703;
						}
						goto IL_072c;
						IL_0663:
						num2 = 65;
						list2.Add(current.Name);
						goto IL_0674;
						IL_0703:
						num2 = 72;
						wpfGrowthDriver2.chkGrowth.IsEnabled = true;
						goto IL_0715;
						IL_0715:
						num2 = 73;
						wpfGrowthDriver2.chkGrowth.IsChecked = true;
						goto IL_072c;
						IL_072c:
						num2 = 74;
						wpfGrowthDriver2.ShowDialog();
						goto IL_0737;
						IL_0737:
						num2 = 75;
						if (wpfGrowthDriver2.DialogResult.HasValue)
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
							if (wpfGrowthDriver2.DialogResult.Value)
							{
								goto IL_0774;
							}
						}
						goto IL_0c7d;
						IL_0674:
						num2 = 66;
						goto IL_0677;
						IL_0bb9:
						num2 = 104;
						goto IL_0bbc;
						IL_0774:
						num2 = 77;
						text = Strings.Replace(wpfGrowthDriver2.cbxColumn.SelectedItem.ToString(), VH.A(91534), "");
						goto IL_07a9;
						IL_07a9:
						num2 = 78;
						range3 = ((_Worksheet)range2.Worksheet).get_Range((object)(text + Conversions.ToString(range2.Row)), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_07dd;
						IL_07dd:
						num2 = 79;
						application.ScreenUpdating = false;
						goto IL_07e7;
						IL_07e7:
						num2 = 80;
						application.EnableEvents = false;
						goto IL_07f1;
						IL_07f1:
						num2 = 81;
						isChecked = wpfGrowthDriver2.chkGrowth.IsChecked;
						if (!isChecked.HasValue)
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
							flag = isChecked;
						}
						else
						{
							flag = isChecked != true;
						}
						isChecked = flag;
						if (isChecked == true)
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
							goto IL_0846;
						}
						goto IL_093d;
						IL_0537:
						num2 = 52;
						if (num5 != 1)
						{
							goto IL_0542;
						}
						goto IL_05ca;
						IL_0587:
						num2 = 55;
						list.Add(VH.A(91534) + text);
						goto IL_05a2;
						IL_0542:
						num2 = 54;
						text = Regex.Replace(range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), VH.A(94825), "");
						goto IL_0587;
						IL_052e:
						num2 = 51;
						num5++;
						goto IL_0537;
						IL_05a2:
						num2 = 56;
						if (range5.Column == range3.Column)
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
							goto IL_05c1;
						}
						goto IL_05ca;
						IL_0846:
						num2 = 82;
						range3.get_Offset((object)1, (object)0).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0881;
						IL_0881:
						num2 = 83;
						range4.get_Offset((object)1, (object)0).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)wpfGrowthDriver2.txtLabel.Text);
						goto IL_08b6;
						IL_08b6:
						num2 = 84;
						range2.get_Offset((object)1, (object)0).NumberFormat = KH.A.CyclePercent.Items[wpfGrowthDriver2.cbxNumFormat.SelectedIndex].Format;
						goto IL_08fe;
						IL_08fe:
						num2 = 85;
						if (wpfGrowthDriver2.chkIndent.IsChecked == true)
						{
							goto IL_091a;
						}
						goto IL_093d;
						IL_091a:
						num2 = 86;
						range4.get_Offset((object)1, (object)0).IndentLevel = 1;
						goto IL_093d;
						IL_093d:
						wpfGrowthDriver2 = null;
						goto IL_0940;
						IL_0940:
						num2 = 88;
						num5 = 0;
						goto IL_0946;
						IL_0946:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_094d;
						IL_094d:
						num2 = 90;
						enumerator3 = range2.GetEnumerator();
						goto IL_0bbc;
						IL_0bbc:
						if (enumerator3.MoveNext())
						{
							range5 = (Range)enumerator3.Current;
							goto IL_0970;
						}
						goto IL_0bca;
						IL_0bca:
						num2 = 105;
						if (enumerator3 is IDisposable)
						{
							(enumerator3 as IDisposable).Dispose();
						}
						goto IL_0be2;
						IL_05c1:
						num2 = 57;
						num6 = num5 - 2;
						goto IL_05ca;
						IL_0be2:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0be9;
						IL_0be9:
						num2 = 107;
						if (KH.A.AutoColorOnEntry)
						{
							goto IL_0bfb;
						}
						goto IL_0c7d;
						IL_0bfb:
						num2 = 108;
						enumerator4 = range2.Areas.GetEnumerator();
						goto IL_0c48;
						IL_0c48:
						if (enumerator4.MoveNext())
						{
							range5 = (Range)enumerator4.Current;
							goto IL_0c22;
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
						goto IL_0c5b;
						IL_05ca:
						num2 = 58;
						goto IL_05cd;
						IL_0c5b:
						num2 = 111;
						if (enumerator4 is IDisposable)
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
							(enumerator4 as IDisposable).Dispose();
						}
						goto IL_0c7d;
						IL_0c97:
						num2 = 115;
						range3 = null;
						goto IL_0c9d;
						IL_0c9d:
						num2 = 116;
						range4 = null;
						goto IL_0ca3;
						IL_0c91:
						num2 = 114;
						range2 = null;
						goto IL_0c97;
						IL_0c22:
						num2 = 109;
						AutoColor.Range(range5.get_Resize((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)));
						goto IL_0c45;
						IL_0c45:
						num2 = 110;
						goto IL_0c48;
						IL_0caf:
						num2 = 119;
						application = null;
						break;
						IL_0970:
						num2 = 91;
						if (!string.IsNullOrEmpty(Conversions.ToString(range5.Text)))
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
							goto IL_0999;
						}
						goto IL_0bb9;
						end_IL_0000_3:
						break;
					}
					num2 = 120;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(94857));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 3779;
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

	private static Range A(Range A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Areas areas = default(Areas);
		Range activeCell = default(Range);
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
				case 298:
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
							goto IL_0011;
						case 4:
							goto IL_0030;
						case 6:
							goto IL_0077;
						case 5:
						case 7:
							goto IL_00e1;
						case 8:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 9:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00e1:
					areas = null;
					break;
					IL_0007:
					num2 = 2;
					areas = A.Areas;
					goto IL_0011;
					IL_0011:
					num2 = 3;
					if (areas.Count == 2)
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
						goto IL_0030;
					}
					goto IL_0077;
					IL_0030:
					num2 = 4;
					NewLateBinding.LateCall(areas.get_Item(2).Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
					goto IL_00e1;
					IL_0077:
					num2 = 6;
					NewLateBinding.LateCall(areas.get_Item(areas.Count).Cells[RuntimeHelpers.GetObjectValue(areas.get_Item(areas.Count).Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
					goto IL_00e1;
					end_IL_0000_2:
					break;
				}
				num2 = 8;
				activeCell = A.Application.ActiveCell;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 298;
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
		return activeCell;
	}

	private static Range A(Range A, Range B = null)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range cell = default(Range);
		Range result = default(Range);
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
				case 316:
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
							goto IL_001f;
						case 5:
							goto IL_0079;
						case 4:
						case 6:
							goto IL_008a;
						case 7:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 8:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_008a:
					num2 = 6;
					A = ((_Worksheet)A.Worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(A.EntireRow.Cells[1, A.Application.Columns.Count], null, VH.A(75064), new object[1] { XlDirection.xlToLeft }, null, null, null)));
					break;
					IL_0007:
					num2 = 2;
					if (B == null)
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
						goto IL_001f;
					}
					goto IL_0079;
					IL_001f:
					num2 = 3;
					cell = (Range)NewLateBinding.LateGet(A.Worksheet.Cells[A.Row, 1], null, VH.A(75064), new object[1] { XlDirection.xlToRight }, null, null, null);
					goto IL_008a;
					IL_0079:
					num2 = 5;
					cell = B.get_End(XlDirection.xlToRight);
					goto IL_008a;
					end_IL_0000_2:
					break;
				}
				num2 = 7;
				result = RangeHelpers.CellsWithNumbers(A);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 316;
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
				switch (4)
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
}
