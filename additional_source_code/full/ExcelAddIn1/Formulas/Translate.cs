using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Translate
{
	private static string m_A = VH.A(157960);

	private static ctpTranslator m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("taskPane")]
	private static CustomTaskPane m_A;

	private static readonly string B = VH.A(157997);

	[CompilerGenerated]
	private static bool m_A;

	private static CustomTaskPane taskPane
	{
		[CompilerGenerated]
		get
		{
			return Translate.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Translate.m_A = value;
		}
	}

	private static bool TranslateFormulas
	{
		[CompilerGenerated]
		get
		{
			return Translate.m_A;
		}
		[CompilerGenerated]
		set
		{
			Translate.m_A = value;
		}
	}

	public static void OpenTranslator()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		int height = default(int);
		CustomTaskPane customTaskPane = default(CustomTaskPane);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num5;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 496:
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
							goto IL_000c;
						case 4:
							goto IL_001b;
						case 5:
							goto IL_0029;
						case 6:
							goto IL_0047;
						case 7:
							goto IL_005f;
						case 9:
							goto IL_006e;
						case 8:
						case 10:
							goto IL_007b;
						case 11:
							goto IL_0087;
						case 12:
							goto IL_00ac;
						case 13:
							goto IL_00c0;
						case 14:
							goto IL_00cc;
						case 15:
							goto IL_00e2;
						case 17:
							goto IL_012f;
						case 19:
							goto IL_013c;
						case 16:
						case 18:
						case 20:
						case 21:
							goto IL_0147;
						case 22:
							goto IL_0152;
						case 23:
							goto IL_015d;
						case 24:
							goto IL_0160;
						case 25:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 26:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0160:
					num2 = 24;
					taskPane.Visible = true;
					break;
					IL_0007:
					num2 = 2;
					height = 46;
					goto IL_000c;
					IL_000c:
					num2 = 3;
					_ = MH.A.Application;
					goto IL_001b;
					IL_001b:
					num2 = 4;
					taskPane = A();
					goto IL_0029;
					IL_0029:
					num2 = 5;
					if (taskPane != null)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0047;
					}
					goto IL_006e;
					IL_013c:
					num2 = 19;
					customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
					goto IL_0147;
					IL_012f:
					num2 = 17;
					customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
					goto IL_0147;
					IL_0147:
					num2 = 21;
					customTaskPane.Height = height;
					goto IL_0152;
					IL_0047:
					num2 = 6;
					Translate.m_A = taskPane.Control as ctpTranslator;
					goto IL_005f;
					IL_005f:
					num2 = 7;
					taskPane.Visible = true;
					goto IL_007b;
					IL_006e:
					num2 = 9;
					Translate.m_A = new ctpTranslator();
					goto IL_007b;
					IL_007b:
					num2 = 10;
					if (taskPane == null)
					{
						goto IL_0087;
					}
					goto IL_00c0;
					IL_0087:
					num2 = 11;
					taskPane = MH.A.CustomTaskPanes.Add(Translate.m_A, Translate.m_A);
					goto IL_00ac;
					IL_00ac:
					num2 = 12;
					Translate.m_A.CTP = taskPane;
					goto IL_00c0;
					IL_00c0:
					num2 = 13;
					customTaskPane = taskPane;
					goto IL_00cc;
					IL_00cc:
					num2 = 14;
					customTaskPane.VisibleChanged += A;
					goto IL_00e2;
					IL_00e2:
					num2 = 15;
					num5 = Conversions.ToInteger(KH.A.SettingsXml.DocumentElement.SelectSingleNode(B).InnerText);
					if (num5 == 0)
					{
						goto IL_012f;
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
					if (num5 == 1)
					{
						goto IL_013c;
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
					goto IL_0147;
					IL_0152:
					num2 = 22;
					customTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
					goto IL_015d;
					IL_015d:
					customTaskPane = null;
					goto IL_0160;
					end_IL_0000_2:
					break;
				}
				num2 = 25;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 496;
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
	}

	public static void CloseTranslator()
	{
		taskPane = A();
		taskPane.Visible = false;
	}

	private static void A(object A, EventArgs B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		CustomTaskPane customTaskPane = default(CustomTaskPane);
		bool visible = default(bool);
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
				case 169:
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
							goto IL_0010;
						case 4:
							goto IL_001c;
						case 5:
							goto IL_0035;
						case 6:
							goto IL_004e;
						case 7:
							goto IL_0057;
						case 8:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 9:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0057:
					num2 = 7;
					KH.A.InvalidateControl(VH.A(157910));
					break;
					IL_0007:
					num2 = 2;
					customTaskPane = A as CustomTaskPane;
					goto IL_0010;
					IL_0010:
					num2 = 3;
					visible = customTaskPane.Visible;
					goto IL_001c;
					IL_001c:
					num2 = 4;
					if (!visible)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0035;
					}
					goto IL_004e;
					IL_0035:
					num2 = 5;
					MH.A.CustomTaskPanes.Remove(customTaskPane);
					goto IL_004e;
					IL_004e:
					num2 = 6;
					TranslateFormulas = visible;
					goto IL_0057;
					end_IL_0000_2:
					break;
				}
				num2 = 8;
				customTaskPane = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 169;
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
			switch (1)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static CustomTaskPane A()
	{
		CustomTaskPaneCollection customTaskPanes = MH.A.CustomTaskPanes;
		checked
		{
			for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
			{
				if (Operators.CompareString(customTaskPanes[i].Title, Translate.m_A, TextCompare: false) != 0)
				{
					continue;
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
					return customTaskPanes[i];
				}
			}
			CustomTaskPane result = default(CustomTaskPane);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				customTaskPanes = null;
				return result;
			}
		}
	}

	public static void ToggleTranslate()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		int num5 = default(int);
		Application application = default(Application);
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
				case 429:
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
							goto IL_0035;
						case 4:
							goto IL_0044;
						case 5:
							goto IL_0056;
						case 6:
							goto IL_0074;
						case 8:
							goto IL_0084;
						case 9:
							goto IL_0093;
						case 11:
							goto IL_00c8;
						case 14:
							goto IL_00d2;
						case 16:
							goto IL_00d9;
						case 17:
							goto IL_00e3;
						case 19:
							goto IL_0118;
						case 7:
						case 10:
						case 12:
						case 13:
						case 15:
						case 18:
						case 20:
						case 21:
							goto IL_0120;
						case 22:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 23:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00d2:
					num2 = 14;
					if (num5 == 2)
					{
						goto IL_00d9;
					}
					goto IL_0118;
					IL_0007:
					num2 = 2;
					num5 = Conversions.ToInteger(KH.A.SettingsXml.DocumentElement.SelectSingleNode(B).InnerText);
					goto IL_0035;
					IL_0035:
					num2 = 3;
					TranslateFormulas = !TranslateFormulas;
					goto IL_0044;
					IL_0044:
					num2 = 4;
					application = MH.A.Application;
					goto IL_0056;
					IL_0056:
					num2 = 5;
					if (!TranslateFormulas)
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
						goto IL_0074;
					}
					goto IL_00d2;
					IL_00d9:
					num2 = 16;
					A((object)null, (Range)null);
					goto IL_00e3;
					IL_00e3:
					num2 = 17;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
					goto IL_0120;
					IL_0118:
					num2 = 19;
					OpenTranslator();
					goto IL_0120;
					IL_0074:
					num2 = 6;
					if (num5 == 2)
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
						goto IL_0084;
					}
					goto IL_00c8;
					IL_0120:
					application = null;
					break;
					IL_0084:
					num2 = 8;
					application.StatusBar = false;
					goto IL_0093;
					IL_0093:
					num2 = 9;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
					goto IL_0120;
					IL_00c8:
					num2 = 11;
					CloseTranslator();
					goto IL_0120;
					end_IL_0000_2:
					break;
				}
				num2 = 22;
				KH.A.InvalidateControl(VH.A(157910));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 429;
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
	}

	public static bool IsTranslating()
	{
		return TranslateFormulas;
	}

	private static void A(object A, Range B)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Application application = default(Application);
		int num = default(int);
		int num3 = default(int);
		int num5 = default(int);
		string pattern = default(string);
		Application application2 = default(Application);
		Range range = default(Range);
		string text = default(string);
		MatchCollection matchCollection = default(MatchCollection);
		int num6 = default(int);
		string text2 = default(string);
		string[] array = default(string[]);
		string[] array2 = default(string[]);
		string text3 = default(string);
		Application application3 = default(Application);
		string[] array3 = default(string[]);
		Range range2 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num4;
				object instance;
				string memberName;
				object[] array4;
				ref string reference;
				object[] array5;
				bool[] obj;
				bool[] array6;
				object obj2;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					application = MH.A.Application;
					goto IL_0012;
				case 1196:
					{
						num = num2;
						switch (num3)
						{
						case 2:
							break;
						case 1:
							goto IL_03d0;
						default:
							goto end_IL_0000;
						}
						goto IL_037b;
					}
					IL_03d0:
					num4 = num + 1;
					num = 0;
					switch (num4)
					{
					case 1:
						break;
					case 2:
						goto IL_0012;
					case 3:
						goto IL_0017;
					case 4:
						goto IL_0025;
					case 5:
						goto IL_002c;
					case 6:
						goto IL_0032;
					case 7:
						goto IL_003f;
					case 9:
						goto IL_0069;
					case 11:
						goto IL_0094;
					case 12:
						goto IL_009f;
					case 13:
						goto IL_00aa;
					case 14:
						goto IL_00b5;
					case 15:
						goto IL_00b8;
					case 16:
						goto IL_00d5;
					case 17:
						goto IL_0100;
					case 18:
						goto IL_011f;
					case 19:
						goto IL_0135;
					case 21:
						goto IL_014f;
					case 22:
						goto IL_015d;
					case 20:
					case 23:
						goto IL_016d;
					case 24:
						goto IL_0174;
					case 25:
						goto IL_018a;
					case 26:
						goto IL_019b;
					case 27:
						goto IL_01bb;
					case 28:
						goto IL_01de;
					case 29:
						goto IL_0226;
					case 30:
						goto IL_0242;
					case 31:
						goto IL_02cd;
					case 32:
						goto IL_02dd;
					case 34:
						goto IL_02f9;
					case 35:
						goto IL_031b;
					case 33:
					case 36:
						goto IL_0320;
					case 37:
						goto IL_0348;
					case 38:
						goto IL_0351;
					case 39:
						goto IL_036f;
					case 40:
					case 50:
						goto IL_037b;
					case 41:
						goto IL_0382;
					case 42:
						goto IL_038d;
					case 43:
						goto IL_0398;
					case 44:
						goto IL_03a3;
					case 45:
						goto IL_03a6;
					case 46:
						goto IL_03ac;
					case 47:
						goto end_IL_0000_2;
					case 8:
					case 10:
					case 49:
						goto IL_03b9;
					default:
						goto end_IL_0000;
					case 48:
					case 51:
						goto end_IL_0000_3;
					}
					goto default;
					IL_0012:
					num2 = 2;
					num5 = 0;
					goto IL_0017;
					IL_0017:
					num2 = 3;
					pattern = VH.A(155686);
					goto IL_0025;
					IL_0025:
					ProjectData.ClearProjectError();
					num3 = 2;
					goto IL_002c;
					IL_002c:
					num2 = 5;
					application2 = application;
					goto IL_0032;
					IL_0032:
					num2 = 6;
					range = application2.ActiveCell;
					goto IL_003f;
					IL_003f:
					num2 = 7;
					if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
					{
						goto IL_0069;
					}
					goto IL_03b9;
					IL_0069:
					num2 = 9;
					if (!Conversions.ToBoolean(range.HasArray))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0094;
					}
					goto IL_03b9;
					IL_0320:
					num2 = 36;
					text = text + VH.A(41385) + matchCollection[num5].ToString();
					goto IL_0348;
					IL_0348:
					num2 = 37;
					num5 = checked(num5 + 1);
					goto IL_0351;
					IL_0351:
					num2 = 38;
					num6 = checked(num6 + 1);
					goto IL_035a;
					IL_0094:
					num2 = 11;
					application2.ScreenUpdating = false;
					goto IL_009f;
					IL_009f:
					num2 = 12;
					application2.EnableEvents = false;
					goto IL_00aa;
					IL_00aa:
					num2 = 13;
					application2.DisplayAlerts = false;
					goto IL_00b5;
					IL_00b5:
					application2 = null;
					goto IL_00b8;
					IL_00b8:
					num2 = 15;
					text = Translate.A(range) + VH.A(157943);
					goto IL_00d5;
					IL_00d5:
					num2 = 16;
					text2 = Regex.Replace(Conversions.ToString(range.Formula), VH.A(157948), "");
					goto IL_0100;
					IL_0100:
					num2 = 17;
					text2 = Regex.Replace(text2, VH.A(157953), "");
					goto IL_011f;
					IL_011f:
					num2 = 18;
					if (Versioned.IsNumeric(text2))
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
						goto IL_0135;
					}
					goto IL_014f;
					IL_03b9:
					num2 = 49;
					application.StatusBar = false;
					goto IL_037b;
					IL_0135:
					num2 = 19;
					array = Regex.Split(text2, VH.A(150544));
					goto IL_016d;
					IL_014f:
					num2 = 21;
					matchCollection = Regex.Matches(text2, pattern);
					goto IL_015d;
					IL_015d:
					num2 = 22;
					array = Regex.Split(text2, pattern);
					goto IL_016d;
					IL_016d:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0174;
					IL_0174:
					num2 = 24;
					array2 = array;
					num6 = 0;
					goto IL_035a;
					IL_035a:
					if (num6 < array2.Length)
					{
						text3 = array2[num6];
						goto IL_018a;
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
					goto IL_036f;
					IL_037b:
					num2 = 40;
					application3 = application;
					goto IL_0382;
					IL_036f:
					num2 = 39;
					application.StatusBar = text;
					goto IL_037b;
					IL_0382:
					num2 = 41;
					application3.ScreenUpdating = true;
					goto IL_038d;
					IL_018a:
					num2 = 25;
					if (!Versioned.IsNumeric(text3))
					{
						goto IL_019b;
					}
					goto IL_02cd;
					IL_019b:
					num2 = 26;
					text3 = Strings.Replace(text3, VH.A(39851), "");
					goto IL_01bb;
					IL_01bb:
					num2 = 27;
					if (Strings.InStr(text3, VH.A(7827)) == 0)
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
						goto IL_01de;
					}
					goto IL_0226;
					IL_038d:
					num2 = 42;
					application3.EnableEvents = true;
					goto IL_0398;
					IL_01de:
					num2 = 28;
					text3 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), VH.A(7827)), text3));
					goto IL_0226;
					IL_0226:
					num2 = 29;
					array3 = Strings.Split(text3, VH.A(7827));
					goto IL_0242;
					IL_0242:
					num2 = 30;
					instance = application.ActiveWorkbook.Sheets[array3[0]];
					memberName = VH.A(41315);
					array4 = new object[1];
					reference = ref array3[1];
					array4[0] = reference;
					array5 = array4;
					obj = new bool[1] { true };
					array6 = obj;
					obj2 = NewLateBinding.LateGet(instance, null, memberName, array4, null, null, obj);
					if (array6[0])
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
						reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array5[0]), typeof(string));
					}
					range2 = (Range)obj2;
					goto IL_02cd;
					IL_03a3:
					application3 = null;
					goto IL_03a6;
					IL_03a6:
					num2 = 45;
					application = null;
					goto IL_03ac;
					IL_0398:
					num2 = 43;
					application3.DisplayAlerts = true;
					goto IL_03a3;
					IL_02cd:
					num2 = 31;
					if (range2 == null)
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
						goto IL_02dd;
					}
					goto IL_02f9;
					IL_03ac:
					num2 = 46;
					range2 = null;
					break;
					IL_02dd:
					num2 = 32;
					text = text + VH.A(41385) + text3;
					goto IL_0320;
					IL_02f9:
					num2 = 34;
					text = text + VH.A(41385) + Translate.A(range2);
					goto IL_031b;
					IL_031b:
					num2 = 35;
					range2 = null;
					goto IL_0320;
					end_IL_0000_2:
					break;
				}
				num2 = 47;
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 1196;
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
			switch (1)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private static string A(Range A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string result = default(string);
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
					break;
				case 58:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
							goto end_IL_0000_3;
						}
						goto default;
					}
					end_IL_0000_2:
					break;
				}
				num2 = 2;
				result = Conversions.ToString(Helpers.GetLabelCell(A).Text);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 58;
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
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
