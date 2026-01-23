using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Keyboard;

public sealed class Shortcuts
{
	public struct RibbonTips
	{
		public string ScreenTip;

		public string SuperTip;
	}

	private static readonly string m_A = XC.A(3286);

	private static readonly string B = XC.A(3319);

	private static Dictionary<string, string> m_A;

	public static Dictionary<string, RibbonTips> dictLookup2;

	[CompilerGenerated]
	private static Dictionary<string, Shortcut> m_A;

	public static Dictionary<string, Shortcut> Dictionary
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
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		object objectValue = RuntimeHelpers.GetObjectValue(application.CustomizationContext);
		Template macabacusTemplate = GetMacabacusTemplate(application);
		XmlNodeList xmlNodeList;
		try
		{
			xmlNodeList = A();
			if (xmlNodeList != null)
			{
				IEnumerator enumerator = default(IEnumerator);
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
					enumerator = xmlNodeList.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							XmlNode obj = (XmlNode)enumerator.Current;
							string value = obj.Attributes[XC.A(3734)].Value;
							string value2 = obj.Attributes[XC.A(3678)].Value;
							_ = null;
							try
							{
								if (value2.Length == 0)
								{
									Microsoft.Office.Interop.Word.Application application2 = application;
									int keyCode = ConvertKeystroke(value, application);
									object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
									((_Application)application2).get_FindKey(keyCode, ref KeyCode).Clear();
									continue;
								}
								if (value.Length <= 0)
								{
									continue;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									AssignShortcut((WdKey)ConvertKeystroke(value, application), value2, application);
									break;
								}
								continue;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0115;
							}
							continue;
							end_IL_0115:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					break;
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
		if (macabacusTemplate != null)
		{
			macabacusTemplate.Saved = true;
			macabacusTemplate = null;
			application.CustomizationContext = RuntimeHelpers.GetObjectValue(objectValue);
		}
		if (Conversion.Val(PC.A.Application.Version) < 15.0)
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
			application.ScreenRefresh();
		}
		application = null;
		xmlNodeList = null;
	}

	public static Template GetMacabacusTemplate(Microsoft.Office.Interop.Word.Application wdApp)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = wdApp.Templates.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Template template = (Template)enumerator.Current;
				if (Operators.CompareString(template.Name, clsUtilities.DOTM_FILE_NAME, TextCompare: false) != 0)
				{
					continue;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					wdApp.CustomizationContext = template;
					return template;
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_005b;
				}
				continue;
				end_IL_005b:
				break;
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
		return null;
	}

	public static void AssignShortcut(WdKey key, string strMacro, Microsoft.Office.Interop.Word.Application wdApp)
	{
		KeyBindings keyBindings = wdApp.KeyBindings;
		string command = XC.A(5371) + strMacro;
		object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
		object CommandParameter = RuntimeHelpers.GetObjectValue(Missing.Value);
		keyBindings.Add(WdKeyCategory.wdKeyCategoryCommand, command, (int)key, ref KeyCode, ref CommandParameter);
	}

	public static void Remove()
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
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
					string value = ((XmlNode)enumerator.Current).Attributes[XC.A(3734)].Value;
					if (value.Length <= 0)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Microsoft.Office.Interop.Word.Application application2 = application;
					int keyCode = ConvertKeystroke(value, application);
					object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
					((_Application)application2).get_FindKey(keyCode, ref KeyCode).Clear();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0097;
					}
					continue;
					end_IL_0097:
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
		//IL_03c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c8: Unknown result type (might be due to invalid IL or missing references)
		RibbonTips ribbonTips = default(RibbonTips);
		try
		{
			if (Shortcuts.m_A == null)
			{
				while (true)
				{
					switch (3)
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
					a.Add(XC.A(5412), XC.A(5412));
					a.Add(XC.A(5441), XC.A(5441));
					a.Add(XC.A(5470), XC.A(5470));
					a.Add(XC.A(5503), XC.A(5503));
					a.Add(XC.A(5526), XC.A(5526));
					a.Add(XC.A(5549), XC.A(5549));
					a.Add(XC.A(5572), XC.A(5572));
					a.Add(XC.A(5595), XC.A(5595));
					a.Add(XC.A(5618), XC.A(5618));
					a.Add(XC.A(5641), XC.A(5660));
					a.Add(XC.A(5685), XC.A(5706));
					a.Add(XC.A(5727), XC.A(5744));
					a.Add(XC.A(5763), XC.A(5784));
					a.Add(XC.A(5805), XC.A(5828));
					a.Add(XC.A(5851), XC.A(5851));
					a.Add(XC.A(5880), XC.A(5880));
					a.Add(XC.A(5907), XC.A(5926));
					a.Add(XC.A(5945), XC.A(5962));
					a.Add(XC.A(5979), XC.A(5979));
					a.Add(XC.A(5992), XC.A(5992));
					a.Add(XC.A(6007), XC.A(6007));
					a.Add(XC.A(6022), XC.A(6022));
					a.Add(XC.A(6035), XC.A(6035));
					a.Add(XC.A(6048), XC.A(6048));
					a.Add(XC.A(6071), XC.A(6071));
					a.Add(XC.A(6092), XC.A(6092));
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
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			dictLookup2 = new Dictionary<string, RibbonTips>();
		}
		if (!dictLookup2.ContainsKey(text))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					string value = NC.A.SettingsXml.SelectSingleNode(ShortcutXpath(text)).Attributes[XC.A(3734)].Value;
					Shortcut val = Dictionary[text];
					string friendlyName = val.FriendlyName;
					string description = val.Description;
					_ = null;
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

	public static int ConvertKeystroke(string stroke, Microsoft.Office.Interop.Word.Application wdApp)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		WdKey wdKey = default(WdKey);
		string text = default(string);
		int num6 = default(int);
		int result = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num5;
				int num7;
				int num8;
				int num9;
				int num10;
				object Arg3;
				object Arg2;
				object Arg;
				int num11;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 4405:
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
							goto IL_001a;
						case 5:
							goto IL_0046;
						case 7:
							goto IL_0051;
						case 8:
							goto IL_0074;
						case 10:
							goto IL_007f;
						case 11:
							goto IL_00a3;
						case 13:
							goto IL_00af;
						case 14:
							goto IL_00c5;
						case 16:
							goto IL_00d1;
						case 17:
							goto IL_00f5;
						case 19:
							goto IL_0101;
						case 20:
							goto IL_0125;
						case 22:
							goto IL_0131;
						case 23:
							goto IL_0155;
						case 25:
							goto IL_0161;
						case 26:
							goto IL_0179;
						case 28:
							goto IL_0185;
						case 29:
							goto IL_01a9;
						case 31:
							goto IL_01b5;
						case 32:
							goto IL_01d7;
						case 34:
							goto IL_01e3;
						case 35:
							goto IL_0207;
						case 37:
							goto IL_0213;
						case 38:
							goto IL_0237;
						case 40:
							goto IL_0243;
						case 41:
							goto IL_0267;
						case 43:
							goto IL_0273;
						case 44:
							goto IL_028d;
						case 46:
							goto IL_0299;
						case 47:
							goto IL_02b3;
						case 49:
							goto IL_02bf;
						case 50:
							goto IL_02e3;
						case 52:
							goto IL_02ef;
						case 53:
							goto IL_0313;
						case 55:
							goto IL_031f;
						case 56:
							goto IL_0337;
						case 58:
							goto IL_0343;
						case 59:
							goto IL_0367;
						case 61:
							goto IL_0373;
						case 62:
							goto IL_0397;
						case 64:
							goto IL_03a3;
						case 65:
							goto IL_03bd;
						case 67:
							goto IL_03c9;
						case 68:
							goto IL_03e1;
						case 70:
							goto IL_03ed;
						case 71:
							goto IL_0411;
						case 73:
							goto IL_041d;
						case 74:
							goto IL_043d;
						case 76:
							goto IL_0449;
						case 77:
							goto IL_046d;
						case 79:
							goto IL_0479;
						case 80:
							goto IL_049b;
						case 82:
							goto IL_04a7;
						case 83:
							goto IL_04cb;
						case 85:
							goto IL_04d7;
						case 86:
							goto IL_04f9;
						case 88:
							goto IL_0505;
						case 89:
							goto IL_0525;
						case 91:
							goto IL_0531;
						case 92:
							goto IL_054b;
						case 94:
							goto IL_0557;
						case 95:
							goto IL_0571;
						case 97:
							goto IL_057d;
						case 98:
							goto IL_0595;
						case 100:
							goto IL_05a1;
						case 101:
							goto IL_05c5;
						case 103:
							goto IL_05d1;
						case 104:
							goto IL_05f5;
						case 106:
							goto IL_0601;
						case 107:
							goto IL_0623;
						case 109:
							goto IL_062f;
						case 110:
							goto IL_0653;
						case 112:
							goto IL_065f;
						case 113:
							goto IL_0681;
						case 115:
							goto IL_068d;
						case 116:
							goto IL_06a5;
						case 118:
							goto IL_06b1;
						case 119:
							goto IL_06d3;
						case 121:
							goto IL_06df;
						case 122:
							goto IL_0703;
						case 124:
							goto IL_070f;
						case 125:
							goto IL_0725;
						case 127:
							goto IL_0731;
						case 128:
							goto IL_0749;
						case 130:
							goto IL_0758;
						case 131:
							goto IL_077d;
						case 133:
							goto IL_078c;
						case 134:
							goto IL_07b1;
						case 136:
							goto IL_07c0;
						case 137:
							goto IL_07e7;
						case 139:
							goto IL_07f6;
						case 140:
							goto IL_081b;
						case 142:
							goto IL_082a;
						case 143:
							goto IL_084f;
						case 145:
							goto IL_085e;
						case 146:
							goto IL_0879;
						case 148:
							goto IL_0888;
						case 149:
							goto IL_08ad;
						case 151:
							goto IL_08bc;
						case 152:
							goto IL_08d7;
						case 154:
							goto IL_08e6;
						case 155:
							goto IL_090d;
						case 157:
							goto IL_091c;
						case 158:
							goto IL_0937;
						case 160:
							goto IL_0946;
						case 161:
							goto IL_096b;
						case 163:
							goto IL_097a;
						case 164:
							goto IL_099f;
						case 166:
							goto IL_09ae;
						case 167:
							goto IL_09c7;
						case 169:
							goto IL_09d9;
						case 170:
							goto IL_09fe;
						case 172:
							goto IL_0a10;
						case 173:
							goto IL_0a2b;
						case 175:
							goto IL_0a3d;
						case 176:
							goto IL_0a5a;
						case 178:
							goto IL_0a6c;
						case 179:
							goto IL_0a87;
						case 181:
							goto IL_0a99;
						case 182:
							goto IL_0ac0;
						case 184:
							goto IL_0acf;
						case 185:
							goto IL_0af4;
						case 187:
							goto IL_0b03;
						case 188:
							goto IL_0b28;
						case 6:
						case 9:
						case 12:
						case 15:
						case 18:
						case 21:
						case 24:
						case 27:
						case 30:
						case 33:
						case 36:
						case 39:
						case 42:
						case 45:
						case 48:
						case 51:
						case 54:
						case 57:
						case 60:
						case 63:
						case 66:
						case 69:
						case 72:
						case 75:
						case 78:
						case 81:
						case 84:
						case 87:
						case 90:
						case 93:
						case 96:
						case 99:
						case 102:
						case 105:
						case 108:
						case 111:
						case 114:
						case 117:
						case 120:
						case 123:
						case 126:
						case 129:
						case 132:
						case 135:
						case 138:
						case 141:
						case 144:
						case 147:
						case 150:
						case 153:
						case 156:
						case 159:
						case 162:
						case 165:
						case 168:
						case 171:
						case 174:
						case 177:
						case 180:
						case 183:
						case 186:
						case 189:
							goto IL_0b35;
						case 190:
							goto IL_0b55;
						case 191:
							goto IL_0b7f;
						case 192:
							goto IL_0ba6;
						case 194:
							goto IL_0bf2;
						case 196:
							goto IL_0c3e;
						case 197:
							goto IL_0c61;
						case 199:
							goto IL_0caf;
						case 201:
							goto IL_0cfd;
						case 202:
							goto IL_0d27;
						case 203:
							goto IL_0d44;
						case 205:
							goto IL_0d8f;
						case 193:
						case 195:
						case 198:
						case 200:
						case 204:
						case 206:
							goto IL_0dd4;
						case 207:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 208:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_08d7:
					num2 = 152;
					wdKey = WdKey.wdKeyF8;
					goto IL_0b35;
					IL_0007:
					num2 = 2;
					stroke = stroke.ToLower();
					goto IL_0015;
					IL_0015:
					num2 = 3;
					text = stroke;
					goto IL_001a;
					IL_001a:
					num2 = 4;
					if (text.EndsWith(XC.A(6113)))
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
						goto IL_0046;
					}
					goto IL_0051;
					IL_08e6:
					num2 = 154;
					if (text.Contains(XC.A(4451)))
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
						goto IL_090d;
					}
					goto IL_091c;
					IL_0bf2:
					num2 = 194;
					Arg = WdKey.wdKeyAlt;
					Arg2 = wdKey;
					Arg3 = RuntimeHelpers.GetObjectValue(Missing.Value);
					num5 = wdApp.BuildKeyCode(WdKey.wdKeyControl, ref Arg, ref Arg2, ref Arg3);
					wdKey = (WdKey)Conversions.ToInteger(Arg2);
					num6 = num5;
					goto IL_0dd4;
					IL_090d:
					num2 = 155;
					wdKey = WdKey.wdKeyF9;
					goto IL_0b35;
					IL_0046:
					num2 = 5;
					wdKey = WdKey.wdKey0;
					goto IL_0b35;
					IL_0051:
					num2 = 7;
					if (text.EndsWith(XC.A(6118)))
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
						goto IL_0074;
					}
					goto IL_007f;
					IL_091c:
					num2 = 157;
					if (text.Contains(XC.A(6356)))
					{
						goto IL_0937;
					}
					goto IL_0946;
					IL_0074:
					num2 = 8;
					wdKey = WdKey.wdKey1;
					goto IL_0b35;
					IL_007f:
					num2 = 10;
					if (text.EndsWith(XC.A(6123)))
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
						goto IL_00a3;
					}
					goto IL_00af;
					IL_0937:
					num2 = 158;
					wdKey = WdKey.wdKeyF10;
					goto IL_0b35;
					IL_00a3:
					num2 = 11;
					wdKey = WdKey.wdKey2;
					goto IL_0b35;
					IL_00af:
					num2 = 13;
					if (text.EndsWith(XC.A(6128)))
					{
						goto IL_00c5;
					}
					goto IL_00d1;
					IL_00c5:
					num2 = 14;
					wdKey = WdKey.wdKey3;
					goto IL_0b35;
					IL_00d1:
					num2 = 16;
					if (text.EndsWith(XC.A(6133)))
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
						goto IL_00f5;
					}
					goto IL_0101;
					IL_0946:
					num2 = 160;
					if (text.Contains(XC.A(6363)))
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
						goto IL_096b;
					}
					goto IL_097a;
					IL_00f5:
					num2 = 17;
					wdKey = WdKey.wdKey4;
					goto IL_0b35;
					IL_0101:
					num2 = 19;
					if (text.EndsWith(XC.A(6138)))
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
						goto IL_0125;
					}
					goto IL_0131;
					IL_0c3e:
					num2 = 196;
					if (text.Contains(XC.A(6407)))
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
						goto IL_0c61;
					}
					goto IL_0caf;
					IL_0125:
					num2 = 20;
					wdKey = WdKey.wdKey5;
					goto IL_0b35;
					IL_0131:
					num2 = 22;
					if (text.EndsWith(XC.A(6143)))
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
						goto IL_0155;
					}
					goto IL_0161;
					IL_096b:
					num2 = 161;
					wdKey = WdKey.wdKeyF11;
					goto IL_0b35;
					IL_0155:
					num2 = 23;
					wdKey = WdKey.wdKey6;
					goto IL_0b35;
					IL_0161:
					num2 = 25;
					if (text.EndsWith(XC.A(6148)))
					{
						goto IL_0179;
					}
					goto IL_0185;
					IL_0179:
					num2 = 26;
					wdKey = WdKey.wdKey7;
					goto IL_0b35;
					IL_0185:
					num2 = 28;
					if (text.EndsWith(XC.A(6153)))
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
						goto IL_01a9;
					}
					goto IL_01b5;
					IL_097a:
					num2 = 163;
					if (text.Contains(XC.A(4471)))
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
						goto IL_099f;
					}
					goto IL_09ae;
					IL_01a9:
					num2 = 29;
					wdKey = WdKey.wdKey8;
					goto IL_0b35;
					IL_01b5:
					num2 = 31;
					if (text.EndsWith(XC.A(6158)))
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
						goto IL_01d7;
					}
					goto IL_01e3;
					IL_0d8f:
					num2 = 205;
					Arg = wdKey;
					Arg2 = RuntimeHelpers.GetObjectValue(Missing.Value);
					Arg3 = RuntimeHelpers.GetObjectValue(Missing.Value);
					num7 = wdApp.BuildKeyCode(WdKey.wdKeyAlt, ref Arg, ref Arg2, ref Arg3);
					wdKey = (WdKey)Conversions.ToInteger(Arg);
					num6 = num7;
					goto IL_0dd4;
					IL_01d7:
					num2 = 32;
					wdKey = WdKey.wdKey9;
					goto IL_0b35;
					IL_01e3:
					num2 = 34;
					if (text.EndsWith(XC.A(6163)))
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
						goto IL_0207;
					}
					goto IL_0213;
					IL_099f:
					num2 = 164;
					wdKey = WdKey.wdKeyF12;
					goto IL_0b35;
					IL_0207:
					num2 = 35;
					wdKey = WdKey.wdKeyA;
					goto IL_0b35;
					IL_0213:
					num2 = 37;
					if (text.EndsWith(XC.A(6168)))
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
						goto IL_0237;
					}
					goto IL_0243;
					IL_09ae:
					num2 = 166;
					if (text.EndsWith(XC.A(6370)))
					{
						goto IL_09c7;
					}
					goto IL_09d9;
					IL_0237:
					num2 = 38;
					wdKey = WdKey.wdKeyB;
					goto IL_0b35;
					IL_0243:
					num2 = 40;
					if (text.EndsWith(XC.A(6173)))
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
						goto IL_0267;
					}
					goto IL_0273;
					IL_09c7:
					num2 = 167;
					wdKey = WdKey.wdKeyComma;
					goto IL_0b35;
					IL_0267:
					num2 = 41;
					wdKey = WdKey.wdKeyC;
					goto IL_0b35;
					IL_0273:
					num2 = 43;
					if (text.EndsWith(XC.A(6178)))
					{
						goto IL_028d;
					}
					goto IL_0299;
					IL_028d:
					num2 = 44;
					wdKey = WdKey.wdKeyD;
					goto IL_0b35;
					IL_0299:
					num2 = 46;
					if (text.EndsWith(XC.A(6183)))
					{
						goto IL_02b3;
					}
					goto IL_02bf;
					IL_02b3:
					num2 = 47;
					wdKey = WdKey.wdKeyE;
					goto IL_0b35;
					IL_02bf:
					num2 = 49;
					if (text.EndsWith(XC.A(6188)))
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
						goto IL_02e3;
					}
					goto IL_02ef;
					IL_09d9:
					num2 = 169;
					if (text.EndsWith(XC.A(4860)))
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
						goto IL_09fe;
					}
					goto IL_0a10;
					IL_02e3:
					num2 = 50;
					wdKey = WdKey.wdKeyF;
					goto IL_0b35;
					IL_02ef:
					num2 = 52;
					if (text.EndsWith(XC.A(6193)))
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
						goto IL_0313;
					}
					goto IL_031f;
					IL_0c61:
					num2 = 197;
					Arg3 = WdKey.wdKeyShift;
					Arg2 = wdKey;
					Arg = RuntimeHelpers.GetObjectValue(Missing.Value);
					num8 = wdApp.BuildKeyCode(WdKey.wdKeyControl, ref Arg3, ref Arg2, ref Arg);
					wdKey = (WdKey)Conversions.ToInteger(Arg2);
					num6 = num8;
					goto IL_0dd4;
					IL_0313:
					num2 = 53;
					wdKey = WdKey.wdKeyG;
					goto IL_0b35;
					IL_031f:
					num2 = 55;
					if (text.EndsWith(XC.A(6198)))
					{
						goto IL_0337;
					}
					goto IL_0343;
					IL_0337:
					num2 = 56;
					wdKey = WdKey.wdKeyH;
					goto IL_0b35;
					IL_0343:
					num2 = 58;
					if (text.EndsWith(XC.A(6203)))
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
						goto IL_0367;
					}
					goto IL_0373;
					IL_09fe:
					num2 = 170;
					wdKey = WdKey.wdKeyPeriod;
					goto IL_0b35;
					IL_0367:
					num2 = 59;
					wdKey = WdKey.wdKeyI;
					goto IL_0b35;
					IL_0373:
					num2 = 61;
					if (text.EndsWith(XC.A(6208)))
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
						goto IL_0397;
					}
					goto IL_03a3;
					IL_0a10:
					num2 = 172;
					if (text.EndsWith(XC.A(6373)))
					{
						goto IL_0a2b;
					}
					goto IL_0a3d;
					IL_0397:
					num2 = 62;
					wdKey = WdKey.wdKeyJ;
					goto IL_0b35;
					IL_03a3:
					num2 = 64;
					if (text.EndsWith(XC.A(6213)))
					{
						goto IL_03bd;
					}
					goto IL_03c9;
					IL_03bd:
					num2 = 65;
					wdKey = WdKey.wdKeyK;
					goto IL_0b35;
					IL_03c9:
					num2 = 67;
					if (text.EndsWith(XC.A(6218)))
					{
						goto IL_03e1;
					}
					goto IL_03ed;
					IL_03e1:
					num2 = 68;
					wdKey = WdKey.wdKeyL;
					goto IL_0b35;
					IL_03ed:
					num2 = 70;
					if (text.EndsWith(XC.A(6223)))
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
						goto IL_0411;
					}
					goto IL_041d;
					IL_0a2b:
					num2 = 173;
					wdKey = WdKey.wdKeySemiColon;
					goto IL_0b35;
					IL_0411:
					num2 = 71;
					wdKey = WdKey.wdKeyM;
					goto IL_0b35;
					IL_041d:
					num2 = 73;
					if (text.EndsWith(XC.A(6228)))
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
						goto IL_043d;
					}
					goto IL_0449;
					IL_0a3d:
					num2 = 175;
					if (text.EndsWith(XC.A(6376)))
					{
						goto IL_0a5a;
					}
					goto IL_0a6c;
					IL_043d:
					num2 = 74;
					wdKey = WdKey.wdKeyN;
					goto IL_0b35;
					IL_0449:
					num2 = 76;
					if (text.EndsWith(XC.A(6233)))
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
						goto IL_046d;
					}
					goto IL_0479;
					IL_0a5a:
					num2 = 176;
					wdKey = WdKey.wdKeySingleQuote;
					goto IL_0b35;
					IL_046d:
					num2 = 77;
					wdKey = WdKey.wdKeyO;
					goto IL_0b35;
					IL_0479:
					num2 = 79;
					if (text.EndsWith(XC.A(6238)))
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
						goto IL_049b;
					}
					goto IL_04a7;
					IL_0a6c:
					num2 = 178;
					if (text.EndsWith(XC.A(6379)))
					{
						goto IL_0a87;
					}
					goto IL_0a99;
					IL_049b:
					num2 = 80;
					wdKey = WdKey.wdKeyP;
					goto IL_0b35;
					IL_04a7:
					num2 = 82;
					if (text.EndsWith(XC.A(6243)))
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
						goto IL_04cb;
					}
					goto IL_04d7;
					IL_0a87:
					num2 = 179;
					wdKey = WdKey.wdKeyOpenSquareBrace;
					goto IL_0b35;
					IL_04cb:
					num2 = 83;
					wdKey = WdKey.wdKeyQ;
					goto IL_0b35;
					IL_04d7:
					num2 = 85;
					if (text.EndsWith(XC.A(6248)))
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
						goto IL_04f9;
					}
					goto IL_0505;
					IL_0a99:
					num2 = 181;
					if (text.EndsWith(XC.A(6382)))
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
						goto IL_0ac0;
					}
					goto IL_0acf;
					IL_04f9:
					num2 = 86;
					wdKey = WdKey.wdKeyR;
					goto IL_0b35;
					IL_0505:
					num2 = 88;
					if (text.EndsWith(XC.A(6253)))
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
						goto IL_0525;
					}
					goto IL_0531;
					IL_0caf:
					num2 = 199;
					Arg = wdKey;
					Arg2 = RuntimeHelpers.GetObjectValue(Missing.Value);
					Arg3 = RuntimeHelpers.GetObjectValue(Missing.Value);
					num9 = wdApp.BuildKeyCode(WdKey.wdKeyControl, ref Arg, ref Arg2, ref Arg3);
					wdKey = (WdKey)Conversions.ToInteger(Arg);
					num6 = num9;
					goto IL_0dd4;
					IL_0525:
					num2 = 89;
					wdKey = WdKey.wdKeyS;
					goto IL_0b35;
					IL_0531:
					num2 = 91;
					if (text.EndsWith(XC.A(6258)))
					{
						goto IL_054b;
					}
					goto IL_0557;
					IL_054b:
					num2 = 92;
					wdKey = WdKey.wdKeyT;
					goto IL_0b35;
					IL_0557:
					num2 = 94;
					if (text.EndsWith(XC.A(6263)))
					{
						goto IL_0571;
					}
					goto IL_057d;
					IL_0571:
					num2 = 95;
					wdKey = WdKey.wdKeyU;
					goto IL_0b35;
					IL_057d:
					num2 = 97;
					if (text.EndsWith(XC.A(6268)))
					{
						goto IL_0595;
					}
					goto IL_05a1;
					IL_0595:
					num2 = 98;
					wdKey = WdKey.wdKeyV;
					goto IL_0b35;
					IL_05a1:
					num2 = 100;
					if (text.EndsWith(XC.A(6273)))
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
						goto IL_05c5;
					}
					goto IL_05d1;
					IL_0ac0:
					num2 = 182;
					wdKey = WdKey.wdKeyCloseSquareBrace;
					goto IL_0b35;
					IL_05c5:
					num2 = 101;
					wdKey = WdKey.wdKeyW;
					goto IL_0b35;
					IL_05d1:
					num2 = 103;
					if (text.EndsWith(XC.A(6278)))
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
						goto IL_05f5;
					}
					goto IL_0601;
					IL_0acf:
					num2 = 184;
					if (text.EndsWith(XC.A(6385)))
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
						goto IL_0af4;
					}
					goto IL_0b03;
					IL_05f5:
					num2 = 104;
					wdKey = WdKey.wdKeyX;
					goto IL_0b35;
					IL_0601:
					num2 = 106;
					if (text.EndsWith(XC.A(6283)))
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
						goto IL_0623;
					}
					goto IL_062f;
					IL_0cfd:
					num2 = 201;
					if (text.Contains(XC.A(6400)))
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
						goto IL_0d27;
					}
					goto IL_0dd4;
					IL_0623:
					num2 = 107;
					wdKey = WdKey.wdKeyY;
					goto IL_0b35;
					IL_062f:
					num2 = 109;
					if (text.EndsWith(XC.A(6288)))
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
						goto IL_0653;
					}
					goto IL_065f;
					IL_0af4:
					num2 = 185;
					wdKey = WdKey.wdKeyEquals;
					goto IL_0b35;
					IL_0653:
					num2 = 110;
					wdKey = WdKey.wdKeyZ;
					goto IL_0b35;
					IL_065f:
					num2 = 112;
					if (text.Contains(XC.A(6293)))
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
						goto IL_0681;
					}
					goto IL_068d;
					IL_0b03:
					num2 = 187;
					if (text.EndsWith(XC.A(6388)))
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
						goto IL_0b28;
					}
					goto IL_0b35;
					IL_0681:
					num2 = 113;
					wdKey = WdKey.wdKeyPageUp;
					goto IL_0b35;
					IL_068d:
					num2 = 115;
					if (text.Contains(XC.A(6302)))
					{
						goto IL_06a5;
					}
					goto IL_06b1;
					IL_06a5:
					num2 = 116;
					wdKey = WdKey.wdKeyPageDown;
					goto IL_0b35;
					IL_06b1:
					num2 = 118;
					if (text.Contains(XC.A(6311)))
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
						goto IL_06d3;
					}
					goto IL_06df;
					IL_0dd4:
					text = null;
					break;
					IL_06d3:
					num2 = 119;
					wdKey = WdKey.wdKeyHome;
					goto IL_0b35;
					IL_06df:
					num2 = 121;
					if (text.Contains(XC.A(6320)))
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
						goto IL_0703;
					}
					goto IL_070f;
					IL_0b28:
					num2 = 188;
					wdKey = WdKey.wdKeyHyphen;
					goto IL_0b35;
					IL_0703:
					num2 = 122;
					wdKey = WdKey.wdKeyEnd;
					goto IL_0b35;
					IL_070f:
					num2 = 124;
					if (text.Contains(XC.A(6327)))
					{
						goto IL_0725;
					}
					goto IL_0731;
					IL_0725:
					num2 = 125;
					wdKey = WdKey.wdKeyInsert;
					goto IL_0b35;
					IL_0731:
					num2 = 127;
					if (text.Contains(XC.A(6334)))
					{
						goto IL_0749;
					}
					goto IL_0758;
					IL_0749:
					num2 = 128;
					wdKey = WdKey.wdKeyDelete;
					goto IL_0b35;
					IL_0758:
					num2 = 130;
					if (text.Contains(XC.A(4409)))
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
						goto IL_077d;
					}
					goto IL_078c;
					IL_0b35:
					num2 = 189;
					if (text.Contains(XC.A(6391)))
					{
						goto IL_0b55;
					}
					goto IL_0cfd;
					IL_077d:
					num2 = 131;
					wdKey = WdKey.wdKeyF1;
					goto IL_0b35;
					IL_078c:
					num2 = 133;
					if (text.Contains(XC.A(4414)))
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
						goto IL_07b1;
					}
					goto IL_07c0;
					IL_0b55:
					num2 = 190;
					if (text.Contains(XC.A(6400)))
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
						goto IL_0b7f;
					}
					goto IL_0c3e;
					IL_07b1:
					num2 = 134;
					wdKey = WdKey.wdKeyF2;
					goto IL_0b35;
					IL_07c0:
					num2 = 136;
					if (text.Contains(XC.A(6341)))
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
						goto IL_07e7;
					}
					goto IL_07f6;
					IL_0d27:
					num2 = 202;
					if (text.Contains(XC.A(6407)))
					{
						goto IL_0d44;
					}
					goto IL_0d8f;
					IL_07e7:
					num2 = 137;
					wdKey = WdKey.wdKeyF3;
					goto IL_0b35;
					IL_07f6:
					num2 = 139;
					if (text.Contains(XC.A(4419)))
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
						goto IL_081b;
					}
					goto IL_082a;
					IL_0b7f:
					num2 = 191;
					if (text.Contains(XC.A(6407)))
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
						goto IL_0ba6;
					}
					goto IL_0bf2;
					IL_081b:
					num2 = 140;
					wdKey = WdKey.wdKeyF4;
					goto IL_0b35;
					IL_082a:
					num2 = 142;
					if (text.Contains(XC.A(4441)))
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
						goto IL_084f;
					}
					goto IL_085e;
					IL_0d44:
					num2 = 203;
					Arg3 = WdKey.wdKeyShift;
					Arg2 = wdKey;
					Arg = RuntimeHelpers.GetObjectValue(Missing.Value);
					num10 = wdApp.BuildKeyCode(WdKey.wdKeyAlt, ref Arg3, ref Arg2, ref Arg);
					wdKey = (WdKey)Conversions.ToInteger(Arg2);
					num6 = num10;
					goto IL_0dd4;
					IL_084f:
					num2 = 143;
					wdKey = WdKey.wdKeyF5;
					goto IL_0b35;
					IL_085e:
					num2 = 145;
					if (text.Contains(XC.A(6346)))
					{
						goto IL_0879;
					}
					goto IL_0888;
					IL_0879:
					num2 = 146;
					wdKey = WdKey.wdKeyF6;
					goto IL_0b35;
					IL_0888:
					num2 = 148;
					if (text.Contains(XC.A(4446)))
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
						goto IL_08ad;
					}
					goto IL_08bc;
					IL_0ba6:
					num2 = 192;
					Arg3 = WdKey.wdKeyAlt;
					Arg2 = WdKey.wdKeyShift;
					Arg = wdKey;
					num11 = wdApp.BuildKeyCode(WdKey.wdKeyControl, ref Arg3, ref Arg2, ref Arg);
					wdKey = (WdKey)Conversions.ToInteger(Arg);
					num6 = num11;
					goto IL_0dd4;
					IL_08ad:
					num2 = 149;
					wdKey = WdKey.wdKeyF7;
					goto IL_0b35;
					IL_08bc:
					num2 = 151;
					if (text.Contains(XC.A(6351)))
					{
						goto IL_08d7;
					}
					goto IL_08e6;
					end_IL_0000_2:
					break;
				}
				num2 = 207;
				result = num6;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 4405;
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
				switch (1)
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

	public static void OverrideHotkeys()
	{
		clsShortcuts.OverrideHotkeys((Action)Remove, (Action)Load);
	}

	public static void ImportHotkeys()
	{
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.DefaultExt = XC.A(6418);
		openFileDialog.Filter = XC.A(6425);
		openFileDialog.Title = XC.A(6470);
		if (openFileDialog.ShowDialog() == DialogResult.OK)
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
			XmlDocument xmlDocument;
			try
			{
				xmlDocument = new XmlDocument();
				xmlDocument.Load(openFileDialog.FileName);
				clsShortcuts.SanitizeShortcutsXml(ref xmlDocument);
				if (xmlDocument.GetElementsByTagName(XC.A(6543)).Count != 1)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						throw new Exception();
					}
				}
				string outerXml = xmlDocument.OuterXml;
				outerXml = Regex.Replace(outerXml, XC.A(6580), [SpecialName] (Match A) => XC.A(18458) + A.Value.Substring(1).ToUpper());
				outerXml = outerXml.Replace(XC.A(6601), XC.A(6616));
				outerXml = outerXml.Replace(XC.A(6625), XC.A(6644));
				outerXml = outerXml.Replace(XC.A(6653), XC.A(6670));
				outerXml = outerXml.Replace(XC.A(6675), XC.A(6696));
				outerXml = outerXml.Replace(XC.A(6705), XC.A(6726));
				outerXml = outerXml.Replace(XC.A(6735), XC.A(6758));
				outerXml = outerXml.Replace(XC.A(6769), XC.A(6758));
				outerXml = outerXml.Replace(XC.A(6792), XC.A(6809));
				outerXml = outerXml.Replace(XC.A(6820), XC.A(6841));
				outerXml = outerXml.Replace(XC.A(6868), XC.A(6885));
				outerXml = outerXml.Replace(XC.A(6896), XC.A(6388));
				Remove();
				xmlDocument.LoadXml(outerXml);
				XmlDocument xmlDocument2 = new XmlDocument();
				xmlDocument2.LoadXml(M.DefaultShortcuts);
				clsShortcuts.UpdateShortcutSettings(ref xmlDocument, xmlDocument2, XC.A(6899));
				xmlDocument2 = null;
				Load();
				Forms.SuccessMessage(XC.A(6930));
				ShortcutManager.Refresh();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(XC.A(7011));
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			xmlDocument = null;
		}
		openFileDialog = null;
	}

	public static void Reset()
	{
		dictLookup2 = null;
		ShortcutManager.Refresh();
	}

	public static void BuildDictionary()
	{
		Dictionary = clsShortcuts.BuildShortcutsDictionary(M.DefaultShortcuts);
	}

	public static string ShortcutXpath(string strVal)
	{
		return XC.A(7082) + strVal + XC.A(7149);
	}

	private static XmlNodeList A()
	{
		return NC.A.SettingsXml.DocumentElement.SelectNodes(XC.A(7154));
	}
}
