using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("9149349D-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface _Presentation
{
	[DispId(2001)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2001)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_1();

	[DispId(2003)]
	Master SlideMaster
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	void ApplyTemplate([In][MarshalAs(UnmanagedType.BStr)] string FileName);

	void _VtblGap3_3();

	[DispId(2011)]
	Slides Slides
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2011)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2012)]
	PageSetup PageSetup
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2012)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_1();

	[DispId(2014)]
	ExtraColors ExtraColors
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_2();

	[DispId(2017)]
	DocumentWindows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2017)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2018)]
	Tags Tags
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2018)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_1();

	[DispId(2020)]
	object BuiltInDocumentProperties
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2020)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[DispId(2021)]
	object CustomDocumentProperties
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2021)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap7_1();

	[DispId(2023)]
	MsoTriState ReadOnly
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2023)]
		get;
	}

	[DispId(2024)]
	string FullName
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2024)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(2025)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2025)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(2026)]
	string Path
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2026)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(2027)]
	MsoTriState Saved
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2027)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2027)]
		[param: In]
		set;
	}

	void _VtblGap8_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2029)]
	[return: MarshalAs(UnmanagedType.Interface)]
	DocumentWindow NewWindow();

	void _VtblGap9_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2035)]
	void Save();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2036)]
	void SaveAs([In][MarshalAs(UnmanagedType.BStr)] string FileName, [In] PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, [In] MsoTriState EmbedTrueTypeFonts = MsoTriState.msoTriStateMixed);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2037)]
	void SaveCopyAs([In][MarshalAs(UnmanagedType.BStr)] string FileName, [In] PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, [In] MsoTriState EmbedTrueTypeFonts = MsoTriState.msoTriStateMixed);

	void _VtblGap10_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2039)]
	void Close();

	void _VtblGap11_31();

	[DispId(2063)]
	Designs Designs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2063)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap12_35();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2096)]
	void ExportAsFixedFormat([In][MarshalAs(UnmanagedType.BStr)] string Path, [In] PpFixedFormatType FixedFormatType, [In] PpFixedFormatIntent Intent = PpFixedFormatIntent.ppFixedFormatIntentScreen, [In] MsoTriState FrameSlides = MsoTriState.msoFalse, [In] PpPrintHandoutOrder HandoutOrder = PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, [In] PpPrintOutputType OutputType = PpPrintOutputType.ppPrintOutputSlides, [In] MsoTriState PrintHiddenSlides = MsoTriState.msoFalse, [In][MarshalAs(UnmanagedType.Interface)] PrintRange PrintRange = null, [In] PpPrintRangeType RangeType = PpPrintRangeType.ppPrintAll, [In][MarshalAs(UnmanagedType.BStr)] string SlideShowName = "", [In] bool IncludeDocProperties = false, [In] bool KeepIRMSettings = true, [In] bool DocStructureTags = true, [In] bool BitmapMissingFonts = true, [In] bool UseISO19005_1 = false, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ExternalExporter);

	void _VtblGap13_6();

	[DispId(2103)]
	CustomXMLParts CustomXMLParts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2103)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2104)]
	bool Final
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2104)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2104)]
		[param: In]
		set;
	}

	void _VtblGap14_7();

	[DispId(2111)]
	SectionProperties SectionProperties
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2111)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
