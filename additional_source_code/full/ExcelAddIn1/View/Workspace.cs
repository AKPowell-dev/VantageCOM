using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class Workspace
{
	private static readonly string m_A = VH.A(175640);

	private static readonly string m_B = VH.A(175619);

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static bool m_B;

	[CompilerGenerated]
	private static bool C;

	[CompilerGenerated]
	private static bool D;

	[CompilerGenerated]
	private static bool E;

	public static bool Maximized
	{
		[CompilerGenerated]
		get
		{
			return Workspace.m_A;
		}
		[CompilerGenerated]
		set
		{
			Workspace.m_A = value;
		}
	} = false;

	private static bool DisplayFormulaBar
	{
		[CompilerGenerated]
		get
		{
			return Workspace.m_B;
		}
		[CompilerGenerated]
		set
		{
			Workspace.m_B = value;
		}
	} = true;

	private static bool DisplayStatusBar
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	} = true;

	private static bool DisplayHeadings
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	} = true;

	private static bool DisplayWorkbookTabs
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	} = true;

	public static void Maximize(bool blnRequireAuthentication)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		bool flag = default(bool);
		bool pressedMso = default(bool);
		Application application = default(Application);
		bool flag2 = default(bool);
		bool flag3 = default(bool);
		Application application2 = default(Application);
		XmlNode xmlNode = default(XmlNode);
		bool flag4 = default(bool);
		bool flag5 = default(bool);
		bool flag6 = default(bool);
		bool flag7 = default(bool);
		bool flag8 = default(bool);
		bool pressedMso2 = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (blnRequireAuthentication)
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
						if (!Licensing.AllowAdvancedViewOperation())
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
							goto IL_002b;
						}
					}
					goto IL_0037;
				case 1505:
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
							goto IL_002b;
						case 4:
							goto IL_0037;
						case 5:
							goto IL_003e;
						case 6:
							goto IL_004f;
						case 7:
							goto IL_0077;
						case 8:
							goto IL_009e;
						case 9:
							goto IL_00c3;
						case 10:
							goto IL_00e9;
						case 11:
							goto IL_010f;
						case 12:
							goto IL_0135;
						case 13:
							goto IL_015b;
						case 14:
							goto IL_0181;
						case 15:
							goto IL_01a3;
						case 16:
							goto IL_01a6;
						case 17:
							goto IL_01ac;
						case 18:
							goto IL_01b7;
						case 19:
							goto IL_01d1;
						case 20:
							goto IL_01e9;
						case 21:
							goto IL_01f6;
						case 22:
							goto IL_0207;
						case 23:
							goto IL_0218;
						case 24:
							goto IL_022e;
						case 25:
							goto IL_0242;
						case 26:
							goto IL_0253;
						case 27:
							goto IL_025e;
						case 28:
							goto IL_026f;
						case 29:
							goto IL_027a;
						case 30:
							goto IL_028b;
						case 31:
							goto IL_029b;
						case 32:
							goto IL_02ac;
						case 33:
							goto IL_02bc;
						case 34:
							goto IL_02d1;
						case 36:
							goto IL_02dc;
						case 37:
							goto IL_02f1;
						case 35:
						case 38:
							goto IL_02fa;
						case 39:
							goto IL_0301;
						case 40:
							goto IL_0309;
						case 41:
							goto IL_032e;
						case 42:
							goto IL_0342;
						case 43:
							goto IL_0373;
						case 44:
							goto IL_03a6;
						case 45:
							goto IL_03ba;
						case 47:
							goto IL_03d5;
						case 48:
							goto IL_03e6;
						case 49:
							goto IL_03f7;
						case 50:
							goto IL_03fe;
						case 51:
							goto IL_040f;
						case 52:
							goto IL_0420;
						case 53:
							goto IL_0434;
						case 54:
							goto IL_0445;
						case 55:
							goto IL_045b;
						case 56:
							goto IL_047a;
						case 57:
							goto IL_0483;
						case 58:
							goto IL_0498;
						case 46:
						case 59:
							goto IL_04a1;
						case 60:
							goto IL_04ac;
						case 61:
							goto IL_04af;
						case 62:
							goto IL_04b7;
						case 63:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 64:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0483:
					num2 = 57;
					if (flag && pressedMso)
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
						goto IL_0498;
					}
					goto IL_04a1;
					IL_03e6:
					num2 = 48;
					application.DisplayFormulaBar = DisplayFormulaBar;
					goto IL_03f7;
					IL_03f7:
					num2 = 49;
					if (flag2)
					{
						goto IL_03fe;
					}
					goto IL_040f;
					IL_040f:
					num2 = 51;
					if (flag3)
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
						goto IL_0420;
					}
					goto IL_0434;
					IL_03fe:
					num2 = 50;
					application.DisplayStatusBar = DisplayStatusBar;
					goto IL_040f;
					IL_04af:
					num2 = 61;
					A();
					goto IL_04b7;
					IL_002b:
					num2 = 2;
					A();
					goto end_IL_0000_3;
					IL_0037:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_003e;
					IL_003e:
					num2 = 5;
					application2 = MH.A.Application;
					goto IL_004f;
					IL_004f:
					num2 = 6;
					xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(168237));
					goto IL_0077;
					IL_0077:
					num2 = 7;
					flag4 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175619)).InnerText);
					goto IL_009e;
					IL_009e:
					num2 = 8;
					flag = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175640)).InnerText);
					goto IL_00c3;
					IL_00c3:
					num2 = 9;
					flag5 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175669)).InnerText);
					goto IL_00e9;
					IL_00e9:
					num2 = 10;
					flag6 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175698)).InnerText);
					goto IL_010f;
					IL_010f:
					num2 = 11;
					flag2 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175727)).InnerText);
					goto IL_0135;
					IL_0135:
					num2 = 12;
					flag3 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175754)).InnerText);
					goto IL_015b;
					IL_015b:
					num2 = 13;
					flag7 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175777)).InnerText);
					goto IL_0181;
					IL_0181:
					num2 = 14;
					flag8 = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(175794)).InnerText);
					goto IL_01a3;
					IL_01a3:
					xmlNode = null;
					goto IL_01a6;
					IL_01a6:
					num2 = 16;
					application = application2;
					goto IL_01ac;
					IL_01ac:
					num2 = 17;
					application.ScreenUpdating = false;
					goto IL_01b7;
					IL_01b7:
					num2 = 18;
					pressedMso = application.CommandBars.GetPressedMso(Workspace.m_A);
					goto IL_01d1;
					IL_01d1:
					num2 = 19;
					pressedMso2 = application.CommandBars.GetPressedMso(Workspace.m_B);
					goto IL_01e9;
					IL_01e9:
					num2 = 20;
					if (!Maximized)
					{
						goto IL_01f6;
					}
					goto IL_03d5;
					IL_01f6:
					num2 = 21;
					DisplayFormulaBar = application.DisplayFormulaBar;
					goto IL_0207;
					IL_0207:
					num2 = 22;
					DisplayStatusBar = application.DisplayStatusBar;
					goto IL_0218;
					IL_0218:
					num2 = 23;
					DisplayHeadings = application.ActiveWindow.DisplayHeadings;
					goto IL_022e;
					IL_022e:
					num2 = 24;
					DisplayWorkbookTabs = application.ActiveWindow.DisplayWorkbookTabs;
					goto IL_0242;
					IL_0242:
					num2 = 25;
					if (flag6)
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
						goto IL_0253;
					}
					goto IL_025e;
					IL_0420:
					num2 = 52;
					application.ActiveWindow.DisplayHeadings = DisplayHeadings;
					goto IL_0434;
					IL_0253:
					num2 = 26;
					application.DisplayFormulaBar = false;
					goto IL_025e;
					IL_025e:
					num2 = 27;
					if (flag2)
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
						goto IL_026f;
					}
					goto IL_027a;
					IL_0434:
					num2 = 53;
					if (flag7)
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
						goto IL_0445;
					}
					goto IL_045b;
					IL_026f:
					num2 = 28;
					application.DisplayStatusBar = false;
					goto IL_027a;
					IL_027a:
					num2 = 29;
					if (flag3)
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
						goto IL_028b;
					}
					goto IL_029b;
					IL_04b7:
					num2 = 62;
					Maximized = !Maximized;
					break;
					IL_028b:
					num2 = 30;
					application.ActiveWindow.DisplayHeadings = false;
					goto IL_029b;
					IL_029b:
					num2 = 31;
					if (flag7)
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
						goto IL_02ac;
					}
					goto IL_02bc;
					IL_0445:
					num2 = 54;
					application.ActiveWindow.DisplayWorkbookTabs = DisplayWorkbookTabs;
					goto IL_045b;
					IL_02ac:
					num2 = 32;
					application.ActiveWindow.DisplayWorkbookTabs = false;
					goto IL_02bc;
					IL_02bc:
					num2 = 33;
					if (flag4 && !pressedMso2)
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
						goto IL_02d1;
					}
					goto IL_02dc;
					IL_045b:
					num2 = 55;
					if (flag4)
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
						if (pressedMso2)
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
							goto IL_047a;
						}
					}
					goto IL_0483;
					IL_0498:
					num2 = 58;
					B(application2);
					goto IL_04a1;
					IL_02d1:
					num2 = 34;
					A(application2);
					goto IL_02fa;
					IL_02dc:
					num2 = 36;
					if (flag && !pressedMso)
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
						goto IL_02f1;
					}
					goto IL_02fa;
					IL_04a1:
					num2 = 59;
					application.ScreenUpdating = true;
					goto IL_04ac;
					IL_04ac:
					application = null;
					goto IL_04af;
					IL_02f1:
					num2 = 37;
					B(application2);
					goto IL_02fa;
					IL_02fa:
					num2 = 38;
					if (flag8)
					{
						goto IL_0301;
					}
					goto IL_032e;
					IL_0301:
					num2 = 39;
					SoftDisable.RemoveTaskPanes();
					goto IL_0309;
					IL_0309:
					num2 = 40;
					application.CommandBars[VH.A(52475)].Visible = false;
					goto IL_032e;
					IL_032e:
					num2 = 41;
					if (flag5)
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
						goto IL_0342;
					}
					goto IL_03ba;
					IL_047a:
					num2 = 56;
					A(application2);
					goto IL_0483;
					IL_0342:
					num2 = 42;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175821)).AddEventHandler(application, new AppEvents_NewWorkbookEventHandler(A));
					goto IL_0373;
					IL_0373:
					num2 = 43;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175844)).AddEventHandler(application, new AppEvents_WorkbookOpenEventHandler(A));
					goto IL_03a6;
					IL_03a6:
					num2 = 44;
					application.ActiveWindow.WindowState = XlWindowState.xlMaximized;
					goto IL_03ba;
					IL_03ba:
					num2 = 45;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175869));
					goto IL_04a1;
					IL_03d5:
					num2 = 47;
					if (flag6)
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
						goto IL_03e6;
					}
					goto IL_03f7;
					end_IL_0000_2:
					break;
				}
				num2 = 63;
				application2 = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1505;
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

	private static void A(Application A)
	{
		A.CommandBars.ExecuteMso(Workspace.m_B);
	}

	private static void B(Application A)
	{
		A.CommandBars.ExecuteMso(Workspace.m_A);
	}

	private static void A()
	{
		KH.A.InvalidateControl(VH.A(168220));
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175821)).RemoveEventHandler(application, new AppEvents_NewWorkbookEventHandler(Workspace.A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175844)).RemoveEventHandler(application, new AppEvents_WorkbookOpenEventHandler(Workspace.A));
		application.WindowState = XlWindowState.xlMaximized;
		application = null;
	}

	public static bool MaximizeOnStartUp(XmlDocument xmlDoc)
	{
		return Conversions.ToBoolean(xmlDoc.DocumentElement.SelectSingleNode(VH.A(175896)).InnerText);
	}
}
