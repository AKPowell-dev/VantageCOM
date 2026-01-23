using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Keyboard;

public sealed class DisabledKeys
{
	public static void ToggleDisabledKey(string id, bool bln)
	{
		XmlDocument settingsXml = NC.A.SettingsXml;
		XmlDocument xmlDocument = settingsXml;
		string left = id.ToUpper();
		WdKey wdKey = default(WdKey);
		if (Operators.CompareString(left, XC.A(3151), TextCompare: false) != 0)
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
			if (Operators.CompareString(left, XC.A(3156), TextCompare: false) != 0)
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
				if (Operators.CompareString(left, XC.A(3169), TextCompare: false) != 0)
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
				}
				else
				{
					wdKey = WdKey.wdKeyScrollLock;
					NC.A.DisableKeyScrollLock = bln;
					xmlDocument.GetElementsByTagName(XC.A(3236)).Item(0).InnerText = Conversions.ToString(bln);
				}
			}
			else
			{
				wdKey = WdKey.wdKeyInsert;
				NC.A.DisableKeyInsert = bln;
				xmlDocument.GetElementsByTagName(XC.A(3209)).Item(0).InnerText = Conversions.ToString(bln);
			}
		}
		else
		{
			wdKey = WdKey.wdKeyF1;
			NC.A.DisableKeyF1 = bln;
			xmlDocument.GetElementsByTagName(XC.A(3190)).Item(0).InnerText = Conversions.ToString(bln);
		}
		NC.A.SaveSettings(settingsXml);
		xmlDocument = null;
		if (bln)
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
			DisableKey(wdKey);
		}
		else
		{
			A(wdKey);
		}
		wdKey = (WdKey)0;
	}

	public static void DisableKey(WdKey key)
	{
		try
		{
			Application application = PC.A.Application;
			application.CustomizationContext = application.NormalTemplate;
			object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
			((_Application)application).get_FindKey((int)key, ref KeyCode).Disable();
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(WdKey A)
	{
		try
		{
			Application application = PC.A.Application;
			application.CustomizationContext = application.NormalTemplate;
			object KeyCode = RuntimeHelpers.GetObjectValue(Missing.Value);
			((_Application)application).get_FindKey((int)A, ref KeyCode).Clear();
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static bool GetKeyState(string id)
	{
		string left = id.ToUpper();
		if (Operators.CompareString(left, XC.A(3151), TextCompare: false) != 0)
		{
			bool result = default(bool);
			if (Operators.CompareString(left, XC.A(3156), TextCompare: false) != 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (Operators.CompareString(left, XC.A(3271), TextCompare: false) == 0)
						{
							return NC.A.DisableKeyNumLock;
						}
						if (Operators.CompareString(left, XC.A(3169), TextCompare: false) != 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									return result;
								}
							}
						}
						return NC.A.DisableKeyScrollLock;
					}
				}
			}
			return NC.A.DisableKeyInsert;
		}
		return NC.A.DisableKeyF1;
	}
}
