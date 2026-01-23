using System;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1;

public sealed class clsSettings
{
	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private bool E;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private bool F;

	[CompilerGenerated]
	private bool G;

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
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public bool SlideNumbersStartAtOne
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

	public bool SequentialSlideNumbers
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

	public bool OverrideSectionActions
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

	public string SectionSubsectionSeparator
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

	public bool ShowBulletPunctuation
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

	public bool TextLinkCompatibilityMode
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
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
		KG.A = null;
		try
		{
			XmlElement documentElement = A.DocumentElement;
			if (Operators.CompareString(documentElement.Name, AH.A(153172), TextCompare: false) != 0)
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
						throw new Exception();
					}
				}
			}
			RebuildTables = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153207) + Constants.XML_REBUILD_TABLES).InnerText);
			ShowAllNames = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153207) + Constants.XML_SHOW_ALL_NAMES).InnerText);
			SlideNumbersStartAtOne = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153234)).InnerText);
			SequentialSlideNumbers = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153297)).InnerText);
			OverrideSectionActions = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153366)).InnerText);
			SectionSubsectionSeparator = documentElement.SelectSingleNode(AH.A(153437)).InnerText;
			ShowBulletPunctuation = Conversions.ToBoolean(documentElement.SelectSingleNode(AH.A(153504)).InnerText);
			TextLinkCompatibilityMode = Conversions.ToBoolean(documentElement.SelectSingleNode(Constants.XML_TEXT_LINK_COMPAT).InnerText);
			Check.CheckOutdatedLibraryContent = Conversions.ToBoolean(documentElement.SelectSingleNode(Constants.XML_CHECK_OUTDATED_LIB_CONTENT).InnerText);
			Create.DefaultTemplateId = documentElement.SelectSingleNode(Constants.XML_DEFAULT_TEMPLATE_ID).InnerText;
			Highlight.LoadColor(Conversions.ToInteger(documentElement.SelectSingleNode(Constants.XML_LINK_HIGHLIGHT_COLOR).InnerText));
			documentElement = null;
			Rules.LoadOptions(A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(AH.A(153561));
			ProjectData.ClearProjectError();
		}
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
		bool num = Manage.Import();
		if (num)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(AH.A(153743));
		}
		return num;
	}

	public static void SettingsReset()
	{
		if (!Manage.Reset())
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			PB.Settings.Reset();
			A(AH.A(153806));
			return;
		}
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
				case 89:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 6:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_001b:
					num2 = 4;
					KG.A.Invalidate();
					break;
					IL_0007:
					num2 = 2;
					KG.A = null;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					KG.A = new clsSettings();
					goto IL_001b;
					end_IL_0000_2:
					break;
				}
				num2 = 5;
				Forms.SuccessMessage(A);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 89;
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
			switch (5)
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
