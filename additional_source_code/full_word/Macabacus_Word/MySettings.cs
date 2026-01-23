using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

[CompilerGenerated]
[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.10.0.0")]
[EditorBrowsable(EditorBrowsableState.Advanced)]
internal sealed class MySettings : ApplicationSettingsBase
{
	private static MySettings defaultInstance = (MySettings)SettingsBase.Synchronized(new MySettings());

	public static MySettings Default => defaultInstance;

	[DefaultSettingValue("Red")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public Color LastFontColor
	{
		get
		{
			object obj = this[XC.A(42798)];
			if (obj == null)
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
						return default(Color);
					}
				}
			}
			return (Color)obj;
		}
		set
		{
			this[XC.A(42798)] = value;
		}
	}

	[DefaultSettingValue("255, 128, 0")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public Color LastFillColor
	{
		get
		{
			object obj = this[XC.A(42825)];
			if (obj == null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return default(Color);
					}
				}
			}
			return (Color)obj;
		}
		set
		{
			this[XC.A(42825)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("DodgerBlue")]
	[DebuggerNonUserCode]
	public Color LastBorderColor
	{
		get
		{
			object obj = this[XC.A(42852)];
			if (obj == null)
			{
				return default(Color);
			}
			return (Color)obj;
		}
		set
		{
			this[XC.A(42852)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ImportMatchDestinationWidth
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(42883)]);
		}
		set
		{
			this[XC.A(42883)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool ImportMatchDestinationHeight
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(42938)]);
		}
		set
		{
			this[XC.A(42938)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowPersonal
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(42995)]);
		}
		set
		{
			this[XC.A(42995)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowShared
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43046)]);
		}
		set
		{
			this[XC.A(43046)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowShapes
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43093)]);
		}
		set
		{
			this[XC.A(43093)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ContentInsertShowImages
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43140)]);
		}
		set
		{
			this[XC.A(43140)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("430")]
	[DebuggerNonUserCode]
	public int ContentInsertPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[XC.A(43187)]);
		}
		set
		{
			this[XC.A(43187)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowCharts
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43232)]);
		}
		set
		{
			this[XC.A(43232)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowText
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43279)]);
		}
		set
		{
			this[XC.A(43279)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ProofingShowErrors
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43322)]);
		}
		set
		{
			this[XC.A(43322)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ProofingShowWarnings
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43359)]);
		}
		set
		{
			this[XC.A(43359)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ProofingShowMessages
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43400)]);
		}
		set
		{
			this[XC.A(43400)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowPreview
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43441)]);
		}
		set
		{
			this[XC.A(43441)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool LibraryPaneKeepSourceFormat
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43486)]);
		}
		set
		{
			this[XC.A(43486)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("430")]
	public int TaskPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[XC.A(43541)]);
		}
		set
		{
			this[XC.A(43541)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ContentInsertShowPublic
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43568)]);
		}
		set
		{
			this[XC.A(43568)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool LibraryPaneShowStars
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43615)]);
		}
		set
		{
			this[XC.A(43615)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowDocs
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43656)]);
		}
		set
		{
			this[XC.A(43656)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowPDFs
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43699)]);
		}
		set
		{
			this[XC.A(43699)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("")]
	[UserScopedSetting]
	public string ContentInsertExcludeImageTypes
	{
		get
		{
			return Conversions.ToString(this[XC.A(43742)]);
		}
		set
		{
			this[XC.A(43742)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowImageTypeBadge
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43803)]);
		}
		set
		{
			this[XC.A(43803)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShow3rdParty
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43862)]);
		}
		set
		{
			this[XC.A(43862)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentShowOnlyOutdated
	{
		get
		{
			return Conversions.ToBoolean(this[XC.A(43913)]);
		}
		set
		{
			this[XC.A(43913)] = value;
		}
	}
}
