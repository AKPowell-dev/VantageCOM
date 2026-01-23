using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Keyboard;

public sealed class Shortcuts
{
	public struct RibbonTips
	{
		public string ScreenTip;

		public string SuperTip;
	}

	private static Dictionary<string, string> m_A;

	public static Dictionary<string, RibbonTips> dictLookup2;

	[CompilerGenerated]
	private static Dictionary<string, Shortcut> m_A;

	public static Dictionary<string, Shortcut> ShortcutsDictionary
	{
		[CompilerGenerated]
		get
		{
			return Shortcuts.m_A;
		}
		[CompilerGenerated]
		set
		{
			Shortcuts.m_A = value;
		}
	} = null;

	public static void Load()
	{
		Application application = MH.A.Application;
		XmlNodeList xmlNodeList;
		try
		{
			xmlNodeList = A();
			if (xmlNodeList != null)
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = xmlNodeList.GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						XmlNode xmlNode2 = xmlNode;
						try
						{
							if (Operators.CompareString(xmlNode2.Attributes[VH.A(161660)].Value, "", TextCompare: false) == 0)
							{
								application.OnKey(ConvertKeystroke(xmlNode2.Attributes[VH.A(161707)].Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							else if (Operators.CompareString(xmlNode2.Attributes[VH.A(161707)].Value, "", TextCompare: false) != 0)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									if (1 == 0)
									{
										/*OpCode not supported: LdMemberToken*/;
									}
									A(xmlNode, application);
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
						xmlNode2 = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0109;
						}
						continue;
						end_IL_0109:
						break;
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
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
		xmlNodeList = null;
	}

	private static void A(XmlNode A, Application B)
	{
		B.OnKey(ConvertKeystroke(A.Attributes[VH.A(161707)].Value), clsUtilities.MacabacusXlamPath() + VH.A(7827) + A.Attributes[VH.A(161660)].Value);
	}

	public static void Remove()
	{
		Application application = MH.A.Application;
		XmlNodeList xmlNodeList;
		try
		{
			xmlNodeList = A();
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = xmlNodeList.GetEnumerator();
				while (enumerator.MoveNext())
				{
					string value = ((XmlNode)enumerator.Current).Attributes[VH.A(161707)].Value;
					if (Operators.CompareString(value, "", TextCompare: false) == 0)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					application.OnKey(ConvertKeystroke(value), RuntimeHelpers.GetObjectValue(Missing.Value));
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		xmlNodeList = null;
		application = null;
	}

	public static string HotkeyScreenTip(string idMso)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string screenTip = default(string);
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
				case 53:
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
				screenTip = A(idMso).ScreenTip;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 53;
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
		return screenTip;
	}

	public static string SuperTip(string idMso)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string superTip = default(string);
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
				case 53:
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
				superTip = A(idMso).SuperTip;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 53;
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
		return superTip;
	}

	private static RibbonTips A(string A)
	{
		//IL_1004: Unknown result type (might be due to invalid IL or missing references)
		//IL_1009: Unknown result type (might be due to invalid IL or missing references)
		//IL_100b: Unknown result type (might be due to invalid IL or missing references)
		RibbonTips ribbonTips = default(RibbonTips);
		try
		{
			if (Shortcuts.m_A == null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Shortcuts.m_A = new Dictionary<string, string>();
					Dictionary<string, string> a = Shortcuts.m_A;
					a.Add(VH.A(163347), VH.A(163370));
					a.Add(VH.A(163393), VH.A(163420));
					a.Add(VH.A(163447), VH.A(163472));
					a.Add(VH.A(163497), VH.A(163524));
					a.Add(VH.A(163551), VH.A(163570));
					a.Add(VH.A(163589), VH.A(163612));
					a.Add(VH.A(163635), VH.A(163656));
					a.Add(VH.A(163677), VH.A(163696));
					a.Add(VH.A(163727), VH.A(163756));
					a.Add(VH.A(163785), VH.A(163814));
					a.Add(VH.A(163843), VH.A(163872));
					a.Add(VH.A(163891), VH.A(163924));
					a.Add(VH.A(163957), VH.A(163988));
					a.Add(VH.A(164019), VH.A(164050));
					a.Add(VH.A(164093), VH.A(164120));
					a.Add(VH.A(164147), VH.A(164172));
					a.Add(VH.A(164197), VH.A(164216));
					a.Add(VH.A(164235), VH.A(164260));
					a.Add(VH.A(164285), VH.A(164306));
					a.Add(VH.A(164327), VH.A(164350));
					a.Add(VH.A(164373), "");
					a.Add(VH.A(164394), VH.A(164419));
					a.Add(VH.A(164454), VH.A(164475));
					a.Add(VH.A(164506), VH.A(164527));
					a.Add(VH.A(164548), VH.A(164565));
					a.Add(VH.A(164584), VH.A(164605));
					a.Add(VH.A(164626), VH.A(164649));
					a.Add(VH.A(164672), VH.A(164689));
					a.Add(VH.A(164706), VH.A(164723));
					a.Add(VH.A(164740), VH.A(164757));
					a.Add(VH.A(164784), VH.A(164801));
					a.Add(VH.A(164828), VH.A(164853));
					a.Add(VH.A(164876), VH.A(164901));
					a.Add(VH.A(164924), VH.A(164947));
					a.Add(VH.A(164980), VH.A(165005));
					a.Add(VH.A(165040), VH.A(165040));
					a.Add(VH.A(165069), VH.A(165069));
					a.Add(VH.A(165096), VH.A(151804));
					a.Add(VH.A(165125), VH.A(165144));
					a.Add(VH.A(165163), VH.A(165182));
					a.Add(VH.A(165201), VH.A(861));
					a.Add(VH.A(165222), VH.A(165235));
					a.Add(VH.A(165248), VH.A(165248));
					a.Add(VH.A(165275), VH.A(165275));
					a.Add(VH.A(165304), VH.A(150594));
					a.Add(VH.A(165321), VH.A(165344));
					a.Add(VH.A(165367), VH.A(165390));
					a.Add(VH.A(165413), VH.A(165436));
					a.Add(VH.A(165459), VH.A(165482));
					a.Add(VH.A(165505), VH.A(165528));
					a.Add(VH.A(165551), VH.A(165574));
					a.Add(VH.A(165597), VH.A(165620));
					a.Add(VH.A(165643), VH.A(165666));
					a.Add(VH.A(165689), VH.A(165714));
					a.Add(VH.A(165751), VH.A(165714));
					a.Add(VH.A(165788), VH.A(165817));
					a.Add(VH.A(165846), VH.A(165881));
					a.Add(VH.A(165916), VH.A(165931));
					a.Add(VH.A(165958), VH.A(165975));
					a.Add(VH.A(166002), VH.A(166025));
					a.Add(VH.A(1968), VH.A(166056));
					a.Add(VH.A(1983), VH.A(166097));
					a.Add(VH.A(166138), VH.A(166165));
					a.Add(VH.A(166200), VH.A(166227));
					a.Add(VH.A(166262), VH.A(1630));
					a.Add(VH.A(166285), VH.A(1507));
					a.Add(VH.A(166308), VH.A(166308));
					a.Add(VH.A(166333), VH.A(166333));
					a.Add(VH.A(99807), VH.A(166356));
					a.Add(VH.A(166385), VH.A(166385));
					a.Add(VH.A(166416), VH.A(166441));
					a.Add(VH.A(166466), VH.A(166489));
					a.Add(VH.A(166512), VH.A(497));
					a.Add(VH.A(166527), VH.A(166544));
					a.Add(VH.A(166561), VH.A(156702));
					a.Add(VH.A(195), VH.A(195));
					a.Add(VH.A(1599), VH.A(1599));
					a.Add(VH.A(166584), VH.A(166611));
					a.Add(VH.A(166638), VH.A(166661));
					a.Add(VH.A(166684), VH.A(166684));
					a.Add(VH.A(166705), VH.A(166705));
					a.Add(VH.A(166734), VH.A(166755));
					a.Add(VH.A(166784), VH.A(166807));
					a.Add(VH.A(166836), VH.A(166857));
					a.Add(VH.A(166878), VH.A(166907));
					a.Add(VH.A(166936), VH.A(166969));
					a.Add(VH.A(166996), VH.A(167017));
					a.Add(VH.A(167038), VH.A(167057));
					a.Add(VH.A(167078), VH.A(167099));
					a.Add(VH.A(167120), VH.A(167143));
					a.Add(VH.A(167166), VH.A(167179));
					a.Add(VH.A(167198), VH.A(167211));
					a.Add(VH.A(167230), VH.A(167243));
					a.Add(VH.A(167262), VH.A(167275));
					a.Add(VH.A(167294), VH.A(167307));
					a.Add(VH.A(167324), VH.A(167337));
					a.Add(VH.A(167354), VH.A(167371));
					a.Add(VH.A(167392), VH.A(167409));
					a.Add(VH.A(167430), VH.A(167445));
					a.Add(VH.A(167460), VH.A(167475));
					a.Add(VH.A(167490), VH.A(167509));
					a.Add(VH.A(167528), VH.A(167547));
					a.Add(VH.A(167566), VH.A(167579));
					a.Add(VH.A(167600), VH.A(167613));
					a.Add(VH.A(167634), VH.A(167647));
					a.Add(VH.A(167672), VH.A(167685));
					a.Add(VH.A(167710), VH.A(167710));
					a.Add(VH.A(167731), VH.A(167750));
					a.Add(VH.A(167769), VH.A(167788));
					a.Add(VH.A(167807), VH.A(167824));
					a.Add(VH.A(167841), VH.A(167860));
					a.Add(VH.A(167883), VH.A(167904));
					a.Add(VH.A(167925), VH.A(167944));
					a.Add(VH.A(167963), VH.A(167980));
					a.Add(VH.A(167999), VH.A(168016));
					a.Add(VH.A(125628), VH.A(125628));
					a.Add(VH.A(168035), VH.A(168035));
					a.Add(VH.A(168062), VH.A(168062));
					a.Add(VH.A(168091), VH.A(168106));
					a.Add(VH.A(168121), VH.A(168136));
					a.Add(VH.A(168151), VH.A(168160));
					a.Add(VH.A(793), VH.A(793));
					a.Add(VH.A(168191), VH.A(168191));
					a.Add(VH.A(168220), VH.A(168237));
					a.Add(VH.A(168272), VH.A(168285));
					a.Add(VH.A(168298), VH.A(168313));
					a.Add(VH.A(168328), VH.A(168328));
					a.Add(VH.A(168347), VH.A(168328));
					a.Add(VH.A(168368), VH.A(168368));
					a.Add(VH.A(168393), VH.A(168393));
					a.Add(VH.A(168416), VH.A(168416));
					a.Add(VH.A(168439), VH.A(168439));
					a.Add(VH.A(168452), VH.A(168479));
					a.Add(VH.A(168506), VH.A(168543));
					a.Add(VH.A(168574), VH.A(168599));
					_ = null;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		string text = Shortcuts.m_A[A];
		if (dictLookup2 == null)
		{
			dictLookup2 = new Dictionary<string, RibbonTips>();
		}
		if (!dictLookup2.ContainsKey(text))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					string value = KH.A.SettingsXml.SelectSingleNode(ShortcutXpath(text)).Attributes[VH.A(161707)].Value;
					Shortcut val = ShortcutsDictionary[text];
					string friendlyName = val.FriendlyName;
					string description = val.Description;
					_ = null;
					clsShortcuts.PrepKeystroke(ref value);
					ribbonTips.ScreenTip = friendlyName;
					ribbonTips.SuperTip = clsShortcuts.SuperTip(description, value);
					dictLookup2.Add(text, ribbonTips);
					return ribbonTips;
				}
				}
			}
		}
		return dictLookup2[text];
	}

	public static string ConvertKeystroke(string stroke)
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
					goto IL_0007;
				case 725:
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
							goto IL_0015;
						case 4:
							goto IL_003b;
						case 5:
							goto IL_005d;
						case 6:
							goto IL_0083;
						case 7:
							goto IL_00a7;
						case 8:
							goto IL_00cd;
						case 9:
							goto IL_00f3;
						case 10:
							goto IL_011a;
						case 11:
							goto IL_0141;
						case 12:
							goto IL_0166;
						case 13:
							goto IL_0189;
						case 14:
							goto IL_01ac;
						case 15:
							goto IL_01d1;
						case 16:
							goto IL_01f6;
						case 17:
							goto IL_021d;
						case 18:
							goto IL_0244;
						case 19:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 20:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0244:
					num2 = 18;
					stroke = stroke.Replace(VH.A(43340), VH.A(168895));
					break;
					IL_0007:
					num2 = 2;
					stroke = Strings.LCase(stroke);
					goto IL_0015;
					IL_0015:
					num2 = 3;
					stroke = stroke.Replace(VH.A(168628), VH.A(41312));
					goto IL_003b;
					IL_003b:
					num2 = 4;
					stroke = stroke.Replace(VH.A(168639), VH.A(168648));
					goto IL_005d;
					IL_005d:
					num2 = 5;
					stroke = stroke.Replace(VH.A(168651), VH.A(54459));
					goto IL_0083;
					IL_0083:
					num2 = 6;
					stroke = stroke.Replace(VH.A(168664), VH.A(168673));
					goto IL_00a7;
					IL_00a7:
					num2 = 7;
					stroke = stroke.Replace(VH.A(168686), VH.A(168695));
					goto IL_00cd;
					IL_00cd:
					num2 = 8;
					stroke = stroke.Replace(VH.A(168708), VH.A(168713));
					goto IL_00f3;
					IL_00f3:
					num2 = 9;
					stroke = stroke.Replace(VH.A(168722), VH.A(168731));
					goto IL_011a;
					IL_011a:
					num2 = 10;
					stroke = stroke.Replace(VH.A(94462), VH.A(168744));
					goto IL_0141;
					IL_0141:
					num2 = 11;
					stroke = stroke.Replace(VH.A(94471), VH.A(168757));
					goto IL_0166;
					IL_0166:
					num2 = 12;
					stroke = stroke.Replace(VH.A(168772), VH.A(168781));
					goto IL_0189;
					IL_0189:
					num2 = 13;
					stroke = stroke.Replace(VH.A(168794), VH.A(168801));
					goto IL_01ac;
					IL_01ac:
					num2 = 14;
					stroke = stroke.Replace(VH.A(168812), VH.A(168819));
					goto IL_01d1;
					IL_01d1:
					num2 = 15;
					stroke = stroke.Replace(VH.A(168836), VH.A(168843));
					goto IL_01f6;
					IL_01f6:
					num2 = 16;
					stroke = Regex.Replace(stroke, VH.A(168854), VH.A(168877));
					goto IL_021d;
					IL_021d:
					num2 = 17;
					stroke = stroke.Replace(VH.A(7120), VH.A(168888));
					goto IL_0244;
					end_IL_0000_2:
					break;
				}
				num2 = 19;
				result = stroke;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 725;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void OverrideHotkeys()
	{
		clsShortcuts.OverrideHotkeys((Action)Remove, (Action)Load);
	}

	public static void ResetShortcuts()
	{
		dictLookup2 = null;
		ShortcutManager.Refresh();
	}

	public static void OpenHotkeysSheet()
	{
		//IL_02bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0305: Unknown result type (might be due to invalid IL or missing references)
		//IL_0391: Unknown result type (might be due to invalid IL or missing references)
		//IL_0413: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		Application application = default(Application);
		int num2 = default(int);
		int num = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				Worksheet worksheet;
				switch (try0000_dispatch)
				{
				default:
				{
					int num3 = 1;
					application = MH.A.Application;
					ProjectData.ClearProjectError();
					num2 = 2;
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					application.PrintCommunication = false;
					worksheet = (Worksheet)application.Workbooks.Add(RuntimeHelpers.GetObjectValue(Missing.Value)).ActiveSheet;
					Worksheet worksheet2 = worksheet;
					worksheet2.Name = VH.A(168902);
					NewLateBinding.LateSetComplex(worksheet2.Cells[num3, 1], null, VH.A(41636), new object[1] { VH.A(19019) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet2.Cells[num3, 2], null, VH.A(41636), new object[1] { VH.A(168941) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet2.Cells[num3, 3], null, VH.A(41636), new object[1] { VH.A(161726) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet2.Cells[num3, 4], null, VH.A(41636), new object[1] { VH.A(163225) }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(worksheet2.Cells[num3, 5], null, VH.A(41636), new object[1] { VH.A(168960) }, null, null, OptimisticSet: false, RValueBase: true);
					Range range = ((_Worksheet)worksheet2).get_Range(RuntimeHelpers.GetObjectValue(worksheet2.Cells[num3, 1]), RuntimeHelpers.GetObjectValue(worksheet2.Cells[num3, 5]));
					range.Font.Bold = true;
					range.Borders[XlBordersIndex.xlEdgeBottom].Color = Information.RGB(0, 0, 0);
					_ = null;
					num3 = checked(num3 + 1);
					IEnumerator enumerator = A().GetEnumerator();
					while (enumerator.MoveNext())
					{
						XmlNode xmlNode = (XmlNode)enumerator.Current;
						string value = xmlNode.Attributes[VH.A(161660)].Value;
						if (!ShortcutsDictionary.TryGetValue(value, out var value2))
						{
							continue;
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (value2.Utility <= 0)
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
							break;
						}
						object instance = worksheet2.Cells[num3, 1];
						NewLateBinding.LateSetComplex(instance, null, VH.A(41636), new object[1] { value2.FriendlyName }, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(instance, null, VH.A(60565), new object[3]
						{
							0,
							1,
							xmlNode.Attributes[VH.A(161707)].Value
						}, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(instance, null, VH.A(60565), new object[3]
						{
							0,
							2,
							Strings.UCase(value2.Category)
						}, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(instance, null, VH.A(60565), new object[3]
						{
							0,
							3,
							value2.Utility.ToString()
						}, null, null, OptimisticSet: false, RValueBase: true);
						NewLateBinding.LateSetComplex(instance, null, VH.A(60565), new object[3] { 0, 4, value2.Description }, null, null, OptimisticSet: false, RValueBase: true);
						instance = null;
						if (num3 % 2 == 0)
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
							((_Worksheet)worksheet2).get_Range(RuntimeHelpers.GetObjectValue(worksheet2.Cells[num3, 1]), RuntimeHelpers.GetObjectValue(worksheet2.Cells[num3, 5])).Interior.Color = Information.RGB(228, 233, 241);
						}
						num3 = checked(num3 + 1);
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
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					object instance2 = worksheet2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
					NewLateBinding.LateCall(instance2, null, VH.A(168983), new object[0], null, null, null, IgnoreReturn: true);
					NewLateBinding.LateSetComplex(instance2, null, VH.A(151729), new object[1] { Operators.AddObject(NewLateBinding.LateGet(instance2, null, VH.A(151729), new object[0], null, null, null), 5) }, null, null, OptimisticSet: false, RValueBase: true);
					instance2 = null;
					object instance3 = worksheet2.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)];
					NewLateBinding.LateCall(instance3, null, VH.A(168983), new object[0], null, null, null, IgnoreReturn: true);
					NewLateBinding.LateSetComplex(instance3, null, VH.A(151729), new object[1] { Operators.AddObject(NewLateBinding.LateGet(instance3, null, VH.A(151729), new object[0], null, null, null), 2) }, null, null, OptimisticSet: false, RValueBase: true);
					instance3 = null;
					object instance4 = worksheet2.Columns[3, RuntimeHelpers.GetObjectValue(Missing.Value)];
					NewLateBinding.LateCall(instance4, null, VH.A(168983), new object[0], null, null, null, IgnoreReturn: true);
					NewLateBinding.LateSetComplex(instance4, null, VH.A(151729), new object[1] { Operators.AddObject(NewLateBinding.LateGet(instance4, null, VH.A(151729), new object[0], null, null, null), 2) }, null, null, OptimisticSet: false, RValueBase: true);
					instance4 = null;
					object instance5 = worksheet2.Columns[4, RuntimeHelpers.GetObjectValue(Missing.Value)];
					NewLateBinding.LateSetComplex(instance5, null, VH.A(151729), new object[1] { 10 }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(instance5, null, VH.A(168998), new object[1] { XlHAlign.xlHAlignCenter }, null, null, OptimisticSet: false, RValueBase: true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance5, null, VH.A(62391), new object[1] { 1 }, null, null, null), null, VH.A(168998), new object[1] { XlHAlign.xlHAlignLeft }, null, null, OptimisticSet: false, RValueBase: true);
					instance5 = null;
					object instance6 = worksheet2.Columns[5, RuntimeHelpers.GetObjectValue(Missing.Value)];
					NewLateBinding.LateCall(instance6, null, VH.A(168983), new object[0], null, null, null, IgnoreReturn: true);
					NewLateBinding.LateSetComplex(instance6, null, VH.A(151729), new object[1] { Operators.AddObject(NewLateBinding.LateGet(instance6, null, VH.A(151729), new object[0], null, null, null), 2) }, null, null, OptimisticSet: false, RValueBase: true);
					instance6 = null;
					IEnumerator enumerator2 = ((_Worksheet)worksheet2).get_Range(RuntimeHelpers.GetObjectValue(worksheet2.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet2.Cells[1, 5])).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Range range2 = (Range)enumerator2.Current;
						int column = range2.Column;
						if (column == 5)
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
							range2.AutoFilter(range2.Column, RuntimeHelpers.GetObjectValue(Missing.Value), XlAutoFilterOperator.xlAnd, RuntimeHelpers.GetObjectValue(Missing.Value), false);
						}
						else
						{
							range2.AutoFilter(range2.Column, RuntimeHelpers.GetObjectValue(Missing.Value), XlAutoFilterOperator.xlAnd, RuntimeHelpers.GetObjectValue(Missing.Value), true);
						}
						range2 = null;
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
					if (enumerator2 is IDisposable)
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
						(enumerator2 as IDisposable).Dispose();
					}
					worksheet2.Rows.RowHeight = 13.5;
					PageSetup pageSetup = worksheet2.PageSetup;
					pageSetup.PrintArea = worksheet.UsedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					pageSetup.Zoom = false;
					pageSetup.FitToPagesTall = 1;
					pageSetup.FitToPagesWide = 1;
					_ = null;
					worksheet2 = null;
					Window activeWindow = application.ActiveWindow;
					activeWindow.DisplayGridlines = false;
					activeWindow.View = XlWindowView.xlPageBreakPreview;
					activeWindow.Zoom = 85;
					_ = null;
					break;
				}
				case 2512:
					num = -1;
					switch (num2)
					{
					case 2:
						break;
					default:
						goto end_IL_0000;
					}
					break;
				}
				application.ScreenUpdating = true;
				application.EnableEvents = true;
				application.PrintCommunication = true;
				application = null;
				worksheet = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num2 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 2512;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
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

	public static void OpenHotkeysPdf()
	{
		string text = string.Empty;
		try
		{
			text = Path.Combine(clsFile.GetDownloadFolder(), VH.A(169037));
			File.WriteAllBytes(text, J.Macabacus_Shortcuts_Cheatsheet);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			string text2;
			if (ex2 is IOException)
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
				if (clsFile.FileIsInUse(text))
				{
					text2 = VH.A(169219);
					goto IL_0096;
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
			}
			text2 = string.Format(VH.A(169106), VH.A(7803), ex2.Message);
			goto IL_0096;
			IL_0096:
			Forms.ErrorMessage(text2);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
			return;
		}
		try
		{
			clsUtilities.RunFile(text);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(string.Format(VH.A(169389), VH.A(7803), ex4.Message));
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}

	public static void BuildDictionary()
	{
		ShortcutsDictionary = clsShortcuts.BuildShortcutsDictionary(J.DefaultShortcuts);
	}

	public static string ShortcutXpath(string strVal)
	{
		return VH.A(169500) + strVal + VH.A(38059);
	}

	private static XmlNodeList A()
	{
		return KH.A.SettingsXml.DocumentElement.SelectNodes(VH.A(169569));
	}
}
