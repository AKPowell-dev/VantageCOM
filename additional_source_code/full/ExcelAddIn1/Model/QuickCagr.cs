using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class QuickCagr
{
	public static void Add()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Application application = default(Application);
		Range range = default(Range);
		Range range2 = default(Range);
		Range range3 = default(Range);
		Range range4 = default(Range);
		Range range5 = default(Range);
		bool flag = default(bool);
		bool flag2 = default(bool);
		IEnumerator enumerator = default(IEnumerator);
		Range range6 = default(Range);
		int num5 = default(int);
		int num6 = default(int);
		int num7 = default(int);
		int num8 = default(int);
		Range range7 = default(Range);
		int num9 = default(int);
		int num10 = default(int);
		wpfCagr wpfCagr2 = default(wpfCagr);
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
						if (!Access.AllowExcelOperation((PlanType)5, (Restriction)1, false))
						{
							goto end_IL_0000;
						}
						goto IL_0011;
					case 1943:
						{
							num = num2;
							switch (num3)
							{
							case 2:
								break;
							case 1:
								goto IL_05f7;
							default:
								goto end_IL_0000_2;
							}
							goto IL_058d;
						}
						IL_05f7:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0011;
						case 4:
							goto IL_0022;
						case 5:
							goto IL_002a;
						case 6:
							goto IL_0031;
						case 7:
							goto IL_0055;
						case 9:
							goto IL_005e;
						case 10:
							goto IL_0065;
						case 11:
							goto IL_006f;
						case 12:
							goto IL_007f;
						case 13:
							goto IL_008b;
						case 14:
							goto IL_0091;
						case 15:
							goto IL_0097;
						case 16:
							goto IL_009d;
						case 17:
							goto IL_00bc;
						case 18:
							goto IL_00cc;
						case 19:
							goto IL_00db;
						case 20:
							goto IL_00f4;
						case 21:
							goto IL_0124;
						case 22:
							goto IL_0135;
						case 24:
							goto IL_013e;
						case 23:
						case 25:
							goto IL_0145;
						case 27:
							goto IL_014d;
						case 28:
							goto IL_015e;
						case 29:
							goto IL_0167;
						case 26:
						case 31:
							goto IL_0179;
						case 30:
						case 32:
							goto IL_0195;
						case 33:
							goto IL_01b7;
						case 34:
							goto IL_01c9;
						case 35:
							goto IL_01d8;
						case 36:
							goto IL_01f1;
						case 37:
							goto IL_0217;
						case 38:
							goto IL_0228;
						case 40:
							goto IL_0231;
						case 39:
						case 41:
							goto IL_0238;
						case 43:
							goto IL_0240;
						case 44:
							goto IL_0251;
						case 45:
							goto IL_025a;
						case 42:
						case 47:
							goto IL_026c;
						case 46:
						case 48:
							goto IL_0288;
						case 49:
							goto IL_02a7;
						case 51:
							goto IL_02c2;
						case 50:
						case 52:
							goto IL_02c8;
						case 53:
							goto IL_02d5;
						case 54:
							goto IL_02df;
						case 55:
							goto IL_02f0;
						case 56:
							goto IL_032f;
						case 57:
							goto IL_033b;
						case 58:
							goto IL_0348;
						case 59:
							goto IL_0385;
						case 60:
							goto IL_0398;
						case 62:
							goto IL_03c0;
						case 64:
							goto IL_03c8;
						case 66:
							goto IL_03d0;
						case 61:
						case 63:
						case 65:
						case 67:
						case 68:
						case 70:
							goto IL_03d7;
						case 71:
							goto IL_03da;
						case 73:
							goto IL_03e4;
						case 74:
							goto IL_03eb;
						case 75:
							goto IL_03f2;
						case 76:
							goto IL_041c;
						case 77:
							goto IL_0455;
						case 78:
							goto IL_0458;
						case 79:
							goto IL_0462;
						case 80:
							goto IL_0517;
						case 81:
							goto IL_0521;
						case 82:
							goto IL_0527;
						case 83:
							goto IL_052d;
						case 84:
							goto IL_0533;
						case 69:
						case 85:
							goto IL_054c;
						case 86:
							goto IL_056e;
						case 87:
							goto IL_0579;
						case 88:
						case 99:
							goto IL_058d;
						case 89:
							goto IL_0594;
						case 90:
							goto IL_059f;
						case 91:
							goto IL_05a9;
						case 92:
							goto IL_05ae;
						case 93:
							goto IL_05b4;
						case 94:
							goto IL_05ba;
						case 95:
							goto IL_05c0;
						case 96:
							goto end_IL_0000_3;
						case 72:
						case 98:
							goto IL_05dc;
						default:
							goto end_IL_0000_2;
						case 2:
						case 8:
						case 97:
						case 100:
							goto end_IL_0000;
						}
						goto default;
						IL_0011:
						num2 = 3;
						application = MH.A.Application;
						goto IL_0022;
						IL_0022:
						num2 = 4;
						Common.ClearVisualizations(application);
						goto IL_002a;
						IL_002a:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0031;
						IL_0031:
						num2 = 6;
						if (!(application.Selection is Range))
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
							goto IL_0055;
						}
						goto IL_005e;
						IL_03eb:
						num2 = 74;
						range = range2;
						goto IL_03f2;
						IL_03f2:
						num2 = 75;
						range3 = (Range)range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_041c;
						IL_041c:
						num2 = 76;
						range4 = (Range)range.Cells[RuntimeHelpers.GetObjectValue(range.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0455;
						IL_0055:
						num2 = 7;
						application = null;
						goto end_IL_0000;
						IL_005e:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_0065;
						IL_0065:
						num2 = 10;
						application.ScreenUpdating = false;
						goto IL_006f;
						IL_006f:
						num2 = 11;
						range5 = (Range)application.Selection;
						goto IL_007f;
						IL_007f:
						num2 = 12;
						flag = JH.A(range5);
						goto IL_008b;
						IL_008b:
						num2 = 13;
						flag2 = false;
						goto IL_0091;
						IL_0091:
						num2 = 14;
						range3 = null;
						goto IL_0097;
						IL_0097:
						num2 = 15;
						range4 = null;
						goto IL_009d;
						IL_009d:
						num2 = 16;
						enumerator = range5.GetEnumerator();
						goto IL_0536;
						IL_0536:
						if (enumerator.MoveNext())
						{
							range6 = (Range)enumerator.Current;
							goto IL_00bc;
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
						goto IL_054c;
						IL_0455:
						range = null;
						goto IL_0458;
						IL_0458:
						num2 = 78;
						application.EnableEvents = false;
						goto IL_0462;
						IL_00bc:
						num2 = 17;
						num5 = -1 * (range6.Column - 1);
						goto IL_00cc;
						IL_00cc:
						num2 = 18;
						num6 = num5;
						num7 = -1;
						goto IL_0182;
						IL_0182:
						if (num7 >= num6)
						{
							goto IL_00db;
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
						goto IL_0195;
						IL_0462:
						num2 = 79;
						range6.Formula = VH.A(94884) + RelativeAddress(range4, (Worksheet)application.ActiveSheet, blnAbsolute: false) + VH.A(75498) + RelativeAddress(range3, (Worksheet)application.ActiveSheet, blnAbsolute: false) + VH.A(94889) + num8 + VH.A(54591) + RelativeAddress(range2, (Worksheet)application.ActiveSheet, blnAbsolute: false) + VH.A(94896);
						goto IL_0517;
						IL_00db:
						num2 = 19;
						range7 = range6.get_Offset((object)0, (object)num7);
						goto IL_00f4;
						IL_00f4:
						num2 = 20;
						if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range7.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
							goto IL_0124;
						}
						goto IL_014d;
						IL_0517:
						num2 = 80;
						application.EnableEvents = true;
						goto IL_0521;
						IL_0124:
						num2 = 21;
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
							goto IL_0135;
						}
						goto IL_013e;
						IL_0521:
						num2 = 81;
						flag2 = true;
						goto IL_0527;
						IL_0135:
						num2 = 22;
						range4 = range7;
						goto IL_0145;
						IL_013e:
						num2 = 24;
						range3 = range7;
						goto IL_0145;
						IL_0145:
						num2 = 25;
						num8 = 0;
						goto IL_0179;
						IL_014d:
						num2 = 27;
						if (range4 != null)
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
							goto IL_015e;
						}
						goto IL_0167;
						IL_0527:
						num2 = 82;
						range3 = null;
						goto IL_052d;
						IL_015e:
						num2 = 28;
						num8++;
						goto IL_0167;
						IL_0167:
						num2 = 29;
						if (num8 <= 1)
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
							goto IL_0179;
						}
						goto IL_0195;
						IL_052d:
						num2 = 83;
						range4 = null;
						goto IL_0533;
						IL_0179:
						num2 = 31;
						num7 += -1;
						goto IL_0182;
						IL_0195:
						num2 = 32;
						if (range4 != null)
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
							if (range3 != null)
							{
								goto IL_0288;
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
						goto IL_01b7;
						IL_0533:
						num2 = 84;
						goto IL_0536;
						IL_05dc:
						num2 = 98;
						Forms.WarningMessage(VH.A(94941));
						goto IL_058d;
						IL_054c:
						num2 = 85;
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
						goto IL_056e;
						IL_01b7:
						num2 = 33;
						num5 = -1 * (range6.Row - 1);
						goto IL_01c9;
						IL_01c9:
						num2 = 34;
						num9 = num5;
						num10 = -1;
						goto IL_0275;
						IL_0275:
						if (num10 >= num9)
						{
							goto IL_01d8;
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
						goto IL_0288;
						IL_05b4:
						num2 = 93;
						range5 = null;
						goto IL_05ba;
						IL_01d8:
						num2 = 35;
						range7 = range6.get_Offset((object)num10, (object)0);
						goto IL_01f1;
						IL_01f1:
						num2 = 36;
						if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range7.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
						{
							goto IL_0217;
						}
						goto IL_0240;
						IL_0217:
						num2 = 37;
						if (range4 == null)
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
							goto IL_0228;
						}
						goto IL_0231;
						IL_05ba:
						num2 = 94;
						range3 = null;
						goto IL_05c0;
						IL_0228:
						num2 = 38;
						range4 = range7;
						goto IL_0238;
						IL_0231:
						num2 = 40;
						range3 = range7;
						goto IL_0238;
						IL_0238:
						num2 = 41;
						num8 = 0;
						goto IL_026c;
						IL_0240:
						num2 = 43;
						if (range4 != null)
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
							goto IL_0251;
						}
						goto IL_025a;
						IL_056e:
						num2 = 86;
						if (flag && flag2)
						{
							goto IL_0579;
						}
						goto IL_058d;
						IL_0251:
						num2 = 44;
						num8++;
						goto IL_025a;
						IL_025a:
						num2 = 45;
						if (num8 <= 1)
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
							goto IL_026c;
						}
						goto IL_0288;
						IL_05c0:
						num2 = 95;
						range4 = null;
						break;
						IL_026c:
						num2 = 47;
						num10 += -1;
						goto IL_0275;
						IL_0288:
						num2 = 48;
						if (range4 != null)
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
							if (range3 != null)
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
								goto IL_02a7;
							}
						}
						goto IL_02c2;
						IL_0579:
						num2 = 87;
						JH.A(range5, VH.A(94911));
						goto IL_058d;
						IL_058d:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0594;
						IL_0594:
						num2 = 89;
						range5.Select();
						goto IL_059f;
						IL_02a7:
						num2 = 49;
						range2 = ((_Worksheet)range6.Worksheet).get_Range((object)range3, (object)range4);
						goto IL_02c8;
						IL_02c2:
						num2 = 51;
						range2 = null;
						goto IL_02c8;
						IL_02c8:
						num2 = 52;
						range6.Select();
						goto IL_02d5;
						IL_02d5:
						num2 = 53;
						wpfCagr2 = new wpfCagr();
						goto IL_02df;
						IL_02df:
						num2 = 54;
						if (range2 != null)
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
							goto IL_02f0;
						}
						goto IL_032f;
						IL_059f:
						num2 = 90;
						application.ScreenUpdating = true;
						goto IL_05a9;
						IL_02f0:
						num2 = 55;
						wpfCagr2.txtAddress.Text = range2.get_Address((object)true, (object)true, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_032f;
						IL_032f:
						num2 = 56;
						wpfCagr2.Tag = range2;
						goto IL_033b;
						IL_033b:
						num2 = 57;
						wpfCagr2.ShowDialog();
						goto IL_0348;
						IL_0348:
						num2 = 58;
						if (wpfCagr2.DialogResult.HasValue && wpfCagr2.DialogResult.Value)
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
							goto IL_0385;
						}
						goto IL_054c;
						IL_05a9:
						num2 = 91;
						application = null;
						goto IL_05ae;
						IL_05ae:
						num2 = 92;
						range7 = null;
						goto IL_05b4;
						IL_0385:
						num2 = 59;
						range2 = (Range)wpfCagr2.Tag;
						goto IL_0398;
						IL_0398:
						num2 = 60;
						switch (wpfCagr2.cbxInterval.SelectedIndex)
						{
						case 0:
							break;
						case 1:
							goto IL_03c8;
						case 2:
							goto IL_03d0;
						default:
							goto IL_03d7;
						}
						goto IL_03c0;
						IL_03d0:
						num2 = 66;
						num8 = 12;
						goto IL_03d7;
						IL_03c8:
						num2 = 64;
						num8 = 4;
						goto IL_03d7;
						IL_03c0:
						num2 = 62;
						num8 = 1;
						goto IL_03d7;
						IL_03d7:
						wpfCagr2 = null;
						goto IL_03da;
						IL_03da:
						num2 = 71;
						if (range2 != null)
						{
							goto IL_03e4;
						}
						goto IL_05dc;
						IL_03e4:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_03eb;
						end_IL_0000_3:
						break;
					}
					num2 = 96;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(94920));
					break;
				}
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1943;
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

	public static string RelativeAddress(Range rng, Worksheet RelativeToSheet, bool blnAbsolute)
	{
		if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(rng.Worksheet.Parent, null, VH.A(19019), new object[0], null, null, null), NewLateBinding.LateGet(RelativeToSheet.Parent, null, VH.A(19019), new object[0], null, null, null), TextCompare: false))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return rng.get_Address((object)blnAbsolute, (object)blnAbsolute, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
		}
		if (Operators.CompareString(rng.Worksheet.Name, RelativeToSheet.Name, TextCompare: false) != 0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					string name = rng.Worksheet.Name;
					string text = rng.get_Address((object)blnAbsolute, (object)blnAbsolute, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					if (Regex.IsMatch(name, VH.A(43285)))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								return name + VH.A(7827) + text;
							}
						}
					}
					return VH.A(39851) + name.Replace(VH.A(39851), VH.A(39854)) + VH.A(43343) + text;
				}
				}
			}
		}
		return rng.get_Address((object)blnAbsolute, (object)blnAbsolute, XlReferenceStyle.xlA1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
	}
}
