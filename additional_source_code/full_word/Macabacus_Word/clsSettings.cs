using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.UI;
using Macabacus_Word.DocBuilder;
using Macabacus_Word.Keyboard;
using Macabacus_Word.Library2.Versioning;
using Macabacus_Word.Links;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

public sealed class clsSettings
{
	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private List<Color> m_A;

	[CompilerGenerated]
	private List<Color> B;

	[CompilerGenerated]
	private List<Color> C;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private bool E;

	[CompilerGenerated]
	private bool F;

	public XmlDocument SettingsXml => Manage.GetSettings();

	public bool RebuildTables
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public bool ShowAllNames
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	public List<Color> FontColorCycle
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public List<Color> FillColorCycle
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public List<Color> BorderColorCycle
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	public bool DisableKeyF1
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
	}

	public bool DisableKeyInsert
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
	}

	public bool DisableKeyNumLock
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
	}

	public bool DisableKeyScrollLock
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	public clsSettings(XmlDocument xmlSettings)
	{
		A(xmlSettings);
	}

	public clsSettings()
	{
		A(Manage.GetXml(false));
	}

	private void A(XmlDocument A)
	{
		NC.A = null;
		XmlNodeList xmlNodeList;
		try
		{
			XmlElement documentElement = SettingsXml.DocumentElement;
			if (Operators.CompareString(documentElement.Name, XC.A(40364), TextCompare: false) != 0)
			{
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
					throw new Exception();
				}
			}
			RebuildTables = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(40399) + Constants.XML_REBUILD_TABLES).InnerText);
			ShowAllNames = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(40399) + Constants.XML_SHOW_ALL_NAMES).InnerText);
			Highlight.LoadColor(Conversions.ToInteger(documentElement.SelectSingleNode(Constants.XML_LINK_HIGHLIGHT_COLOR).InnerText));
			FontColorCycle = new List<Color>();
			xmlNodeList = documentElement.SelectNodes(XC.A(40426));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = xmlNodeList.GetEnumerator();
				while (enumerator.MoveNext())
				{
					XmlNode xmlNode = (XmlNode)enumerator.Current;
					FontColorCycle.Add(clsColors.RGB2Color(xmlNode.InnerText));
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			FillColorCycle = new List<Color>();
			xmlNodeList = documentElement.SelectNodes(XC.A(40479));
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = xmlNodeList.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					XmlNode xmlNode2 = (XmlNode)enumerator2.Current;
					FillColorCycle.Add(clsColors.RGB2Color(xmlNode2.InnerText));
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_01ac;
					}
					continue;
					end_IL_01ac:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			BorderColorCycle = new List<Color>();
			xmlNodeList = documentElement.SelectNodes(XC.A(40532));
			IEnumerator enumerator3 = xmlNodeList.GetEnumerator();
			try
			{
				while (enumerator3.MoveNext())
				{
					XmlNode xmlNode3 = (XmlNode)enumerator3.Current;
					BorderColorCycle.Add(clsColors.RGB2Color(xmlNode3.InnerText));
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_023a;
					}
					continue;
					end_IL_023a:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator3 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			DisableKeyF1 = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(3190)).InnerText);
			DisableKeyInsert = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(3209)).InnerText);
			DisableKeyNumLock = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(40589)).InnerText);
			DisableKeyScrollLock = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(3236)).InnerText);
			Base.AutoFieldPreview = Conversions.ToBoolean(documentElement.SelectSingleNode(XC.A(40618)).InnerText);
			Check.CheckOutdatedLibraryContent = Conversions.ToBoolean(documentElement.SelectSingleNode(Constants.XML_CHECK_OUTDATED_LIB_CONTENT).InnerText);
			documentElement = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(XC.A(40669));
			ProjectData.ClearProjectError();
		}
		xmlNodeList = null;
	}

	public void SaveSettings(XmlDocument xmlSettings)
	{
		Manage.Save(xmlSettings, true);
	}

	public static void SettingsExport()
	{
		Manage.Export();
	}

	public static bool SettingsImport()
	{
		Shortcuts.Remove();
		bool num = Manage.Import();
		if (num)
		{
			A(XC.A(40851));
		}
		Shortcuts.Load();
		return num;
	}

	public static void SettingsReset()
	{
		Shortcuts.Remove();
		if (Manage.Reset())
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
			N.Settings.Reset();
			A(XC.A(40914));
		}
		Shortcuts.Load();
	}

	private static void A(string A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
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
				case 100:
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
							goto IL_000f;
						case 4:
							goto IL_001b;
						case 5:
							goto IL_0022;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0022:
					num2 = 5;
					NC.A.Invalidate();
					break;
					IL_0007:
					num2 = 2;
					NC.A = null;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					NC.A = new clsSettings();
					goto IL_001b;
					IL_001b:
					num2 = 4;
					Shortcuts.Reset();
					goto IL_0022;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				Forms.SuccessMessage(A);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 100;
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
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}
