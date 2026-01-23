using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("Name")]
[CompilerGenerated]
[TypeIdentifier]
[Guid("0002096B-0000-0000-C000-000000000046")]
public interface _Document
{
	[DispId(0)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(1)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_4();

	[DispId(3)]
	string Path
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap2_4();

	[DispId(9)]
	Comments Comments
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(9)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_9();

	[DispId(15)]
	Sections Sections
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(15)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_12();

	[DispId(29)]
	string FullName
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(29)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(30)]
	Revisions Revisions
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(30)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_2();

	[DispId(1101)]
	PageSetup PageSetup
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1101)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1101)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	[DispId(34)]
	Windows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(34)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_6();

	[DispId(40)]
	bool Saved
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(40)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(40)]
		[param: In]
		set;
	}

	void _VtblGap7_1();

	[DispId(42)]
	Window ActiveWindow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(42)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_2();

	[DispId(44)]
	bool ReadOnly
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(44)]
		get;
	}

	void _VtblGap9_14();

	[DispId(56)]
	StoryRanges StoryRanges
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(56)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_3();

	[DispId(60)]
	WdProtectionType ProtectionType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(60)]
		get;
	}

	void _VtblGap11_1();

	[DispId(62)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(62)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(63)]
	ListTemplates ListTemplates
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(63)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap12_3();

	[DispId(67)]
	object AttachedTemplate
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(67)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(67)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(68)]
	InlineShapes InlineShapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap13_72();

	[DispId(314)]
	bool TrackRevisions
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(314)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(314)]
		[param: In]
		set;
	}

	void _VtblGap14_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1105)]
	void Close([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveChanges, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OriginalFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object RouteDocument);

	void _VtblGap15_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(108)]
	void Save();

	void _VtblGap16_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2000)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Range([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object End);

	void _VtblGap17_118();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(376)]
	void SaveAs([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LockComments, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object EmbedTrueTypeFonts, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveNativePictureFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveFormsData, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveAsAOCELetter, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object InsertLineBreaks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AllowSubstitutions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LineEnding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddBiDiMarks);

	void _VtblGap18_61();

	[DispId(502)]
	bool TrackFormatting
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(502)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(502)]
		[param: In]
		set;
	}

	void _VtblGap19_7();

	[DispId(508)]
	ContentControls ContentControls
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(508)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap20_13();

	[DispId(521)]
	CustomXMLParts CustomXMLParts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(521)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap21_7();

	[DispId(527)]
	bool Final
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(527)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(527)]
		[param: In]
		set;
	}

	void _VtblGap22_19();

	[DispId(545)]
	OfficeTheme DocumentTheme
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(545)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap23_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(552)]
	void ExportAsFixedFormat([In][MarshalAs(UnmanagedType.BStr)] string OutputFileName, [In] WdExportFormat ExportFormat, [In] bool OpenAfterExport = false, [In] WdExportOptimizeFor OptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint, [In] WdExportRange Range = WdExportRange.wdExportAllDocument, [In] int From = 1, [In] int To = 1, [In] WdExportItem Item = WdExportItem.wdExportDocumentContent, [In] bool IncludeDocProps = false, [In] bool KeepIRM = true, [In] WdExportCreateBookmarks CreateBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks, [In] bool DocStructureTags = true, [In] bool BitmapMissingFonts = true, [In] bool UseISO19005_1 = false, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FixedFormatExtClassPtr);

	void _VtblGap24_16();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(568)]
	void SaveAs2([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LockComments, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object EmbedTrueTypeFonts, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveNativePictureFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveFormsData, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveAsAOCELetter, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object InsertLineBreaks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AllowSubstitutions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LineEnding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddBiDiMarks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object CompatibilityMode);

	void _VtblGap25_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(621)]
	void SaveCopyAs([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FileFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LockComments, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddToRecentFiles, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WritePassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object EmbedTrueTypeFonts, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveNativePictureFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveFormsData, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveAsAOCELetter, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Encoding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object InsertLineBreaks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AllowSubstitutions, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LineEnding, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object AddBiDiMarks, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object CompatibilityMode);
}
