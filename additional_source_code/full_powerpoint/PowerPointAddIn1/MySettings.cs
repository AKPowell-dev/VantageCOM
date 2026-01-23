using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

[GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.10.0.0")]
[EditorBrowsable(EditorBrowsableState.Advanced)]
[CompilerGenerated]
internal sealed class MySettings : ApplicationSettingsBase
{
	private static MySettings defaultInstance = (MySettings)SettingsBase.Synchronized(new MySettings());

	public static MySettings Default => defaultInstance;

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("Red")]
	public Color LastFontColor
	{
		get
		{
			object obj = this[AH.A(167156)];
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
			this[AH.A(167156)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("255, 128, 0")]
	[DebuggerNonUserCode]
	public Color LastFillColor
	{
		get
		{
			object obj = this[AH.A(167183)];
			if (obj == null)
			{
				return default(Color);
			}
			return (Color)obj;
		}
		set
		{
			this[AH.A(167183)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("DodgerBlue")]
	public Color LastBorderColor
	{
		get
		{
			object obj = this[AH.A(167210)];
			if (obj == null)
			{
				return default(Color);
			}
			return (Color)obj;
		}
		set
		{
			this[AH.A(167210)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("0")]
	[DebuggerNonUserCode]
	public int ExplorerPanePosn
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(167241)]);
		}
		set
		{
			this[AH.A(167241)] = value;
		}
	}

	[DefaultSettingValue("350")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public int ExplorerPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(167274)]);
		}
		set
		{
			this[AH.A(167274)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool ExplorerShowImages
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167309)]);
		}
		set
		{
			this[AH.A(167309)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ExplorerShowCharts
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167346)]);
		}
		set
		{
			this[AH.A(167346)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ExplorerShowLinkedShapes
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167383)]);
		}
		set
		{
			this[AH.A(167383)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool ExplorerShowHyperlinks
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167432)]);
		}
		set
		{
			this[AH.A(167432)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ExplorerShowTables
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167477)]);
		}
		set
		{
			this[AH.A(167477)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ExplorerShowEmbeddedExcel
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167514)]);
		}
		set
		{
			this[AH.A(167514)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ExplorerShowEmbeddedWord
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167565)]);
		}
		set
		{
			this[AH.A(167565)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool ExplorerShowComments
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167614)]);
		}
		set
		{
			this[AH.A(167614)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool TableOfContentsSubtitles
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167655)]);
		}
		set
		{
			this[AH.A(167655)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ConformLayoutSizesAndPosition
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167704)]);
		}
		set
		{
			this[AH.A(167704)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ConformLayoutFormats
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167763)]);
		}
		set
		{
			this[AH.A(167763)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ConformLayoutInsertMissingPlaceholders
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167804)]);
		}
		set
		{
			this[AH.A(167804)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool ConformLayoutDeleteSuperfluous
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167881)]);
		}
		set
		{
			this[AH.A(167881)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ImportMatchDestinationWidth
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167942)]);
		}
		set
		{
			this[AH.A(167942)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ImportMatchDestinationHeight
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(167997)]);
		}
		set
		{
			this[AH.A(167997)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("1")]
	public int PictureSizePrompt
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(168054)]);
		}
		set
		{
			this[AH.A(168054)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("")]
	[DebuggerNonUserCode]
	public string TemplateManifestPersonal
	{
		get
		{
			return Conversions.ToString(this[AH.A(168089)]);
		}
		set
		{
			this[AH.A(168089)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("")]
	public string ContentManifestPersonal
	{
		get
		{
			return Conversions.ToString(this[AH.A(168138)]);
		}
		set
		{
			this[AH.A(168138)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("")]
	[DebuggerNonUserCode]
	public string DefaultTemplateId
	{
		get
		{
			return Conversions.ToString(this[AH.A(168185)]);
		}
		set
		{
			this[AH.A(168185)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ProofingShowErrors
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168220)]);
		}
		set
		{
			this[AH.A(168220)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ProofingShowWarnings
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168257)]);
		}
		set
		{
			this[AH.A(168257)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ProofingShowMessages
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168298)]);
		}
		set
		{
			this[AH.A(168298)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool ContentInsertShowPersonal
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168339)]);
		}
		set
		{
			this[AH.A(168339)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowShared
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168390)]);
		}
		set
		{
			this[AH.A(168390)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ContentInsertShowSlides
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168437)]);
		}
		set
		{
			this[AH.A(168437)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool ContentInsertShowShapes
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168484)]);
		}
		set
		{
			this[AH.A(168484)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowImages
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168531)]);
		}
		set
		{
			this[AH.A(168531)] = value;
		}
	}

	[DefaultSettingValue("430")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public int ContentInsertPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(168578)]);
		}
		set
		{
			this[AH.A(168578)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowCharts
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168623)]);
		}
		set
		{
			this[AH.A(168623)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool SelectMatchWidth
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168670)]);
		}
		set
		{
			this[AH.A(168670)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SelectMatchHeight
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168703)]);
		}
		set
		{
			this[AH.A(168703)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool SelectMatchTop
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168738)]);
		}
		set
		{
			this[AH.A(168738)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool SelectMatchLeft
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168767)]);
		}
		set
		{
			this[AH.A(168767)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool SelectMatchShapeType
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168798)]);
		}
		set
		{
			this[AH.A(168798)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool SelectMatchFill
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168839)]);
		}
		set
		{
			this[AH.A(168839)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool SelectMatchFont
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168870)]);
		}
		set
		{
			this[AH.A(168870)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool SelectMatchBorder
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168901)]);
		}
		set
		{
			this[AH.A(168901)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SelectMatchZOrderAbove
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168936)]);
		}
		set
		{
			this[AH.A(168936)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool SelectMatchZOrderBelow
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(168981)]);
		}
		set
		{
			this[AH.A(168981)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool AirplaneMode
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(1938)]);
		}
		set
		{
			this[AH.A(1938)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("0.05")]
	[UserScopedSetting]
	public float AirplaneModeMinWidth
	{
		get
		{
			return Conversions.ToSingle(this[AH.A(169026)]);
		}
		set
		{
			this[AH.A(169026)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("0.05")]
	public float AirplaneModeMinHeight
	{
		get
		{
			return Conversions.ToSingle(this[AH.A(169067)]);
		}
		set
		{
			this[AH.A(169067)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("3")]
	public float AirplaneModeMaxWidth
	{
		get
		{
			return Conversions.ToSingle(this[AH.A(169110)]);
		}
		set
		{
			this[AH.A(169110)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("3")]
	public string AirplaneModeMaxHeight
	{
		get
		{
			return Conversions.ToString(this[AH.A(169151)]);
		}
		set
		{
			this[AH.A(169151)] = value;
		}
	}

	[DefaultSettingValue("200")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int AirplaneModePeek
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(169194)]);
		}
		set
		{
			this[AH.A(169194)] = value;
		}
	}

	[DefaultSettingValue("20")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public int ExplorerTreeNodeSpacing
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(169227)]);
		}
		set
		{
			this[AH.A(169227)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowText
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169274)]);
		}
		set
		{
			this[AH.A(169274)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool ContentInsertShowDecks
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169317)]);
		}
		set
		{
			this[AH.A(169317)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	[UserScopedSetting]
	public bool LibraryPaneShowPreview
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169362)]);
		}
		set
		{
			this[AH.A(169362)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool LibraryPaneKeepSourceFormat
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169407)]);
		}
		set
		{
			this[AH.A(169407)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("430")]
	public int TaskPaneWidth
	{
		get
		{
			return Conversions.ToInteger(this[AH.A(169462)]);
		}
		set
		{
			this[AH.A(169462)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ExplorerShowSmartArt
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169489)]);
		}
		set
		{
			this[AH.A(169489)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ExplorerShowAllPresentations
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169530)]);
		}
		set
		{
			this[AH.A(169530)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ExplorerPreviews
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169587)]);
		}
		set
		{
			this[AH.A(169587)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool ExplorerShowNotes
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169620)]);
		}
		set
		{
			this[AH.A(169620)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ExplorerShowMedia
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169655)]);
		}
		set
		{
			this[AH.A(169655)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ExplorerShowInk
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169690)]);
		}
		set
		{
			this[AH.A(169690)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowPublic
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169721)]);
		}
		set
		{
			this[AH.A(169721)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowStars
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169768)]);
		}
		set
		{
			this[AH.A(169768)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool ContentInsertShowPDFs
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169809)]);
		}
		set
		{
			this[AH.A(169809)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("")]
	public string ContentInsertExcludeImageTypes
	{
		get
		{
			return Conversions.ToString(this[AH.A(169852)]);
		}
		set
		{
			this[AH.A(169852)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool LibraryPaneShowImageTypeBadge
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169913)]);
		}
		set
		{
			this[AH.A(169913)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	public bool ContentInsertShowVideos
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(169972)]);
		}
		set
		{
			this[AH.A(169972)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("True")]
	public bool ContentInsertShow3rdParty
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170019)]);
		}
		set
		{
			this[AH.A(170019)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool SelectMatchAdjustments
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170070)]);
		}
		set
		{
			this[AH.A(170070)] = value;
		}
	}

	[UserScopedSetting]
	[DefaultSettingValue("False")]
	[DebuggerNonUserCode]
	public bool SelectMatchFreeformPoints
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170115)]);
		}
		set
		{
			this[AH.A(170115)] = value;
		}
	}

	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	[UserScopedSetting]
	public bool SelectMatchBottom
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170166)]);
		}
		set
		{
			this[AH.A(170166)] = value;
		}
	}

	[DebuggerNonUserCode]
	[UserScopedSetting]
	[DefaultSettingValue("False")]
	public bool SelectMatchRight
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170201)]);
		}
		set
		{
			this[AH.A(170201)] = value;
		}
	}

	[DefaultSettingValue("False")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool SelectMatchRotation
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170234)]);
		}
		set
		{
			this[AH.A(170234)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[DebuggerNonUserCode]
	[UserScopedSetting]
	public bool PaginateDuplex
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170273)]);
		}
		set
		{
			this[AH.A(170273)] = value;
		}
	}

	[DefaultSettingValue("True")]
	[UserScopedSetting]
	[DebuggerNonUserCode]
	public bool PaginateFlysheetsFront
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170302)]);
		}
		set
		{
			this[AH.A(170302)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool PaginateBindingsView
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170347)]);
		}
		set
		{
			this[AH.A(170347)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("True")]
	public bool LibraryPaneOfferArrange
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170388)]);
		}
		set
		{
			this[AH.A(170388)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("False")]
	public bool ContentShowOnlyOutdated
	{
		get
		{
			return Conversions.ToBoolean(this[AH.A(170435)]);
		}
		set
		{
			this[AH.A(170435)] = value;
		}
	}

	[UserScopedSetting]
	[DebuggerNonUserCode]
	[DefaultSettingValue("0")]
	public float LastGapSize
	{
		get
		{
			return Conversions.ToSingle(this[AH.A(170482)]);
		}
		set
		{
			this[AH.A(170482)] = value;
		}
	}
}
