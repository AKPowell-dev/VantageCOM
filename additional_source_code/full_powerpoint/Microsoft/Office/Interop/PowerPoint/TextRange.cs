using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("9149348F-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
[CompilerGenerated]
[DefaultMember("Text")]
public interface TextRange : Collection
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap1_1();

	[DispId(11)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		get;
	}

	void _VtblGap2_2();

	[DispId(2003)]
	ActionSettings ActionSettings
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2004)]
	int Start
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
	}

	[DispId(2005)]
	int Length
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		get;
	}

	void _VtblGap3_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2010)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange Paragraphs([In] int Start = -1, [In] int Length = -1);

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2012)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange Words([In] int Start = -1, [In] int Length = -1);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2013)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange Characters([In] int Start = -1, [In] int Length = -1);

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2015)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange Runs([In] int Start = -1, [In] int Length = -1);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2016)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange TrimText();

	[DispId(0)]
	string Text
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2017)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange InsertAfter([In][MarshalAs(UnmanagedType.BStr)] string NewText = "");

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2018)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange InsertBefore([In][MarshalAs(UnmanagedType.BStr)] string NewText = "");

	void _VtblGap6_3();

	[DispId(2022)]
	Font Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2022)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2023)]
	ParagraphFormat ParagraphFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2023)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2024)]
	int IndentLevel
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2024)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2024)]
		[param: In]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2025)]
	void Select();

	void _VtblGap7_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2027)]
	void Copy();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2028)]
	void Delete();

	void _VtblGap8_11();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2039)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange PasteSpecial([In] PpPasteDataType DataType = PpPasteDataType.ppPasteDefault, [In] MsoTriState DisplayAsIcon = MsoTriState.msoFalse, [In][MarshalAs(UnmanagedType.BStr)] string IconFileName = "", [In] int IconIndex = 0, [In][MarshalAs(UnmanagedType.BStr)] string IconLabel = "", [In] MsoTriState Link = MsoTriState.msoFalse);
}
