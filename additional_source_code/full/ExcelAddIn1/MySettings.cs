using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[EditorBrowsable(EditorBrowsableState.Advanced)]
[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.10.0.0")]
[CompilerGenerated]
internal sealed class MySettings : ApplicationSettingsBase
{
	private static MySettings defaultInstance = (MySettings)SettingsBase.Synchronized(new MySettings());

	public static MySettings Default => defaultInstance;

	[DefaultSettingValue("Blue")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public Color LastFontColor
	{
		get
		{
			object obj = this[VH.A(207712)];
			if (obj == null)
			{
				return default(Color);
			}
			return (Color)obj;
		}
		set
		{
			this[VH.A(207712)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("Yellow")]
	[UserScopedSetting]
	public Color LastFillColor
	{
		get
		{
			object obj = this[VH.A(207739)];
			if (obj == null)
			{
				while (true)
				{
					switch (6)
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
			this[VH.A(207739)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("Black")]
	[DebuggerNonUserCode]
	public Color LastBorderColor
	{
		get
		{
			object obj = this[VH.A(207766)];
			if (obj == null)
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
						return default(Color);
					}
				}
			}
			return (Color)obj;
		}
		set
		{
			this[VH.A(207766)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("4")]
	public int AuditDockPosition
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(207797)]);
		}
		set
		{
			this[VH.A(207797)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("350")]
	[UserScopedSetting]
	public int AuditFormWidth
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(207832)]);
		}
		set
		{
			this[VH.A(207832)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("250")]
	public int AuditFormHeight
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(207861)]);
		}
		set
		{
			this[VH.A(207861)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("2")]
	public int ExplorerPanePosn
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(207892)]);
		}
		set
		{
			this[VH.A(207892)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("350")]
	public int ExplorerPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(207925)]);
		}
		set
		{
			this[VH.A(207925)] = value;
		}
	}

	[DefaultSettingValue("Links")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public string PasteTranspose
	{
		get
		{
			return Conversions.ToString(this[VH.A(166755)]);
		}
		set
		{
			this[VH.A(166755)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool AutoTracePrecedents
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(207960)]);
		}
		set
		{
			this[VH.A(207960)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool AutoTraceDependents
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(207999)]);
		}
		set
		{
			this[VH.A(207999)] = value;
		}
	}

	[DefaultSettingValue("0")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int AuditFormTop
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(208038)]);
		}
		set
		{
			this[VH.A(208038)] = value;
		}
	}

	[DefaultSettingValue("0")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int AuditFormLeft
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(208063)]);
		}
		set
		{
			this[VH.A(208063)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool NameScrubberShowDependents
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208090)]);
		}
		set
		{
			this[VH.A(208090)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool AuditEvaluateFormulas
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208143)]);
		}
		set
		{
			this[VH.A(208143)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AuditUnhideRowsColumns
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208186)]);
		}
		set
		{
			this[VH.A(208186)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool AuditTraceArrows
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208231)]);
		}
		set
		{
			this[VH.A(208231)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool AuditFormulaWrap
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208264)]);
		}
		set
		{
			this[VH.A(208264)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("42")]
	[UserScopedSetting]
	public int AuditFormulaTextBoxHeight
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(208297)]);
		}
		set
		{
			this[VH.A(208297)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ExportMatchDestinationWidth
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208348)]);
		}
		set
		{
			this[VH.A(208348)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool ExportMatchDestinationHeight
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208403)]);
		}
		set
		{
			this[VH.A(208403)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool AuditFormMoveOnNavigate
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208460)]);
		}
		set
		{
			this[VH.A(208460)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool AdvancedFindFiltersValue
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208507)]);
		}
		set
		{
			this[VH.A(208507)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool AdvancedFindFiltersText
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208556)]);
		}
		set
		{
			this[VH.A(208556)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AdvancedFindFiltersDate
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208603)]);
		}
		set
		{
			this[VH.A(208603)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AdvancedFindFiltersFormat
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208650)]);
		}
		set
		{
			this[VH.A(208650)] = value;
		}
	}

	[DefaultSettingValue("0")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int AdvancedFindScope
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(208701)]);
		}
		set
		{
			this[VH.A(208701)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("0")]
	[DebuggerNonUserCode]
	public int AdvancedFindSelectMode
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(208736)]);
		}
		set
		{
			this[VH.A(208736)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool AdvancedFindLookInValues
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208781)]);
		}
		set
		{
			this[VH.A(208781)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool AdvancedFindLookInFormulas
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208830)]);
		}
		set
		{
			this[VH.A(208830)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool AdvancedFindLookInComments
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208883)]);
		}
		set
		{
			this[VH.A(208883)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool AdvancedFindLookInHyperlinks
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208936)]);
		}
		set
		{
			this[VH.A(208936)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool AdvancedFindLookInEmptyCells
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(208993)]);
		}
		set
		{
			this[VH.A(208993)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool AdvancedFindMatchCase
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209050)]);
		}
		set
		{
			this[VH.A(209050)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowPersonal
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209093)]);
		}
		set
		{
			this[VH.A(209093)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowShared
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209144)]);
		}
		set
		{
			this[VH.A(209144)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ContentInsertShowCharts
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209191)]);
		}
		set
		{
			this[VH.A(209191)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowShapes
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209238)]);
		}
		set
		{
			this[VH.A(209238)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowImages
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209285)]);
		}
		set
		{
			this[VH.A(209285)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowModules
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209332)]);
		}
		set
		{
			this[VH.A(209332)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("430")]
	public int ContentInsertPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(209381)]);
		}
		set
		{
			this[VH.A(209381)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("0")]
	public int ConformSizeBehavior
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(209426)]);
		}
		set
		{
			this[VH.A(209426)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ConformSizeShowGuide
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209465)]);
		}
		set
		{
			this[VH.A(209465)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool FootnotesInspectPrintAreas
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209506)]);
		}
		set
		{
			this[VH.A(209506)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool FootnotesSearchByRows
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209559)]);
		}
		set
		{
			this[VH.A(209559)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SimplifyIndirect
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209602)]);
		}
		set
		{
			this[VH.A(209602)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SimplifyChoose
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209635)]);
		}
		set
		{
			this[VH.A(209635)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool SimplifyOffset
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209664)]);
		}
		set
		{
			this[VH.A(209664)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool SimplifyHlookup
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209693)]);
		}
		set
		{
			this[VH.A(209693)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool SimplifyVlookup
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209724)]);
		}
		set
		{
			this[VH.A(209724)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SimplifyIndexMatch
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209755)]);
		}
		set
		{
			this[VH.A(209755)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("Red")]
	[DebuggerNonUserCode]
	public Color DiscussPenColor
	{
		get
		{
			object obj = this[VH.A(209792)];
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
			this[VH.A(209792)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("1")]
	[DebuggerNonUserCode]
	public int DiscussPenThickness
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(209823)]);
		}
		set
		{
			this[VH.A(209823)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("Yellow")]
	public Color DiscussHighlighterColor
	{
		get
		{
			object obj = this[VH.A(209862)];
			if (obj == null)
			{
				return default(Color);
			}
			return (Color)obj;
		}
		set
		{
			this[VH.A(209862)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool DiscussUsePen
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209909)]);
		}
		set
		{
			this[VH.A(209909)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool DiscussEmbedFiles
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209936)]);
		}
		set
		{
			this[VH.A(209936)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool SimplifyIf
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209971)]);
		}
		set
		{
			this[VH.A(209971)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool SimplifyMin
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(209992)]);
		}
		set
		{
			this[VH.A(209992)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool SimplifyMax
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210015)]);
		}
		set
		{
			this[VH.A(210015)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool AuditHighlightCells
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210038)]);
		}
		set
		{
			this[VH.A(210038)] = value;
		}
	}

	[DefaultSettingValue("1")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int MoveDataLabelsStep
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(210077)]);
		}
		set
		{
			this[VH.A(210077)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ContentInsertShowText
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210114)]);
		}
		set
		{
			this[VH.A(210114)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ContentInsertShowTables
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210157)]);
		}
		set
		{
			this[VH.A(210157)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool AuditOpenWorkbookLinks
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210204)]);
		}
		set
		{
			this[VH.A(210204)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool AdvancedFindLookInPrintAreas
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210249)]);
		}
		set
		{
			this[VH.A(210249)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool LibraryPaneShowPreview
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210306)]);
		}
		set
		{
			this[VH.A(210306)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool LibraryPaneKeepSourceFormat
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210351)]);
		}
		set
		{
			this[VH.A(210351)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("430")]
	[DebuggerNonUserCode]
	public int TaskPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[VH.A(210406)]);
		}
		set
		{
			this[VH.A(210406)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ExplorerPreviews
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210433)]);
		}
		set
		{
			this[VH.A(210433)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AuditGroupDependents
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210466)]);
		}
		set
		{
			this[VH.A(210466)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AuditEvaluateArguments
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210507)]);
		}
		set
		{
			this[VH.A(210507)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SimplifyXlookup
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210552)]);
		}
		set
		{
			this[VH.A(210552)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ContentInsertShowPublic
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210583)]);
		}
		set
		{
			this[VH.A(210583)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowStars
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210630)]);
		}
		set
		{
			this[VH.A(210630)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowModels
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210671)]);
		}
		set
		{
			this[VH.A(210671)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowPDFs
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210718)]);
		}
		set
		{
			this[VH.A(210718)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("")]
	public string ContentInsertExcludeImageTypes
	{
		get
		{
			return Conversions.ToString(this[VH.A(210761)]);
		}
		set
		{
			this[VH.A(210761)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowImageTypeBadge
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210822)]);
		}
		set
		{
			this[VH.A(210822)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShow3rdParty
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210881)]);
		}
		set
		{
			this[VH.A(210881)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool AuditShowExplanations
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210932)]);
		}
		set
		{
			this[VH.A(210932)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool SuperFindHighlightResults
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(210975)]);
		}
		set
		{
			this[VH.A(210975)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool AdvancedFindLookInCharts
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(211026)]);
		}
		set
		{
			this[VH.A(211026)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentShowOnlyOutdated
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(211075)]);
		}
		set
		{
			this[VH.A(211075)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool SimplifySumIf
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(211122)]);
		}
		set
		{
			this[VH.A(211122)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool SimplifySumIfs
	{
		get
		{
			return Conversions.ToBoolean(this[VH.A(211149)]);
		}
		set
		{
			this[VH.A(211149)] = value;
		}
	}
}
