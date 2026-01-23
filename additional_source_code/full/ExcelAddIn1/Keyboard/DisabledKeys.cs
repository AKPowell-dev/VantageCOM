using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Keyboard;

public sealed class DisabledKeys
{
	private static bool m_A;

	private static bool m_B;

	private static bool C;

	private static bool D;

	private static readonly string m_A = VH.A(161284);

	private static readonly string m_B = VH.A(161289);

	private static readonly string C = VH.A(161302);

	private static readonly string D = VH.A(161317);

	public static bool DisableKeyF1
	{
		get
		{
			return DisabledKeys.m_A;
		}
		set
		{
			DisabledKeys.m_A = value;
			try
			{
				if (value)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							A(DisabledKeys.m_A);
							return;
						}
					}
				}
				B(DisabledKeys.m_A);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static bool DisableKeyInsert
	{
		get
		{
			return DisabledKeys.m_B;
		}
		set
		{
			DisabledKeys.m_B = value;
			try
			{
				if (value)
				{
					A(DisabledKeys.m_B);
				}
				else
				{
					B(DisabledKeys.m_B);
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static bool DisableKeyNumLock
	{
		get
		{
			return DisabledKeys.C;
		}
		set
		{
			DisabledKeys.C = value;
			try
			{
				if (value)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							A(C);
							return;
						}
					}
				}
				B(C);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static bool DisableKeyScrollLock
	{
		get
		{
			return DisabledKeys.D;
		}
		set
		{
			DisabledKeys.D = value;
			try
			{
				if (value)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							A(D);
							return;
						}
					}
				}
				B(D);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void ToggleDisabledKey(string id, bool bln)
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		XmlElement documentElement = settingsXml.DocumentElement;
		if (Operators.CompareString(id, DisabledKeys.m_A, TextCompare: false) == 0)
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
			DisableKeyF1 = bln;
			documentElement.SelectSingleNode(VH.A(161174)).InnerText = Conversions.ToString(bln);
		}
		else if (Operators.CompareString(id, DisabledKeys.m_B, TextCompare: false) == 0)
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
			DisableKeyInsert = bln;
			documentElement.SelectSingleNode(VH.A(161193)).InnerText = Conversions.ToString(bln);
		}
		else if (Operators.CompareString(id, C, TextCompare: false) == 0)
		{
			DisableKeyNumLock = bln;
			documentElement.SelectSingleNode(VH.A(161220)).InnerText = Conversions.ToString(bln);
		}
		else if (Operators.CompareString(id, D, TextCompare: false) == 0)
		{
			DisableKeyScrollLock = bln;
			documentElement.SelectSingleNode(VH.A(161249)).InnerText = Conversions.ToString(bln);
		}
		documentElement = null;
		KH.A.SaveSettings(settingsXml);
		settingsXml = null;
	}

	private static void A(string A)
	{
		MH.A.Application.OnKey(VH.A(19799) + A + VH.A(19802), "");
	}

	private static void B(string A)
	{
		MH.A.Application.OnKey(VH.A(19799) + A + VH.A(19802), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public static bool GetKeyState(string id)
	{
		if (Operators.CompareString(id, DisabledKeys.m_A, TextCompare: false) == 0)
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
					return DisableKeyF1;
				}
			}
		}
		if (Operators.CompareString(id, DisabledKeys.m_B, TextCompare: false) == 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return DisableKeyInsert;
				}
			}
		}
		if (Operators.CompareString(id, C, TextCompare: false) == 0)
		{
			return DisableKeyNumLock;
		}
		if (Operators.CompareString(id, D, TextCompare: false) == 0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return DisableKeyScrollLock;
				}
			}
		}
		bool result = default(bool);
		return result;
	}
}
