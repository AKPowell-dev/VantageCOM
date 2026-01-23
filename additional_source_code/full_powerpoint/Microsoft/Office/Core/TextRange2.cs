using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[DefaultMember("Text")]
[Guid("000C0397-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface TextRange2 : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_2();

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

	[DispId(1)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 Item([In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	[DispId(3)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[DispId(4)]
	TextRange2 Paragraphs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(4)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(5)]
	TextRange2 Sentences
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(5)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(6)]
	TextRange2 Words
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(7)]
	TextRange2 Characters
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(7)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(8)]
	TextRange2 Lines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(8)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(9)]
	TextRange2 Runs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(9)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(10)]
	ParagraphFormat2 ParagraphFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(10)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(11)]
	Font2 Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(12)]
	int Length
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(12)]
		get;
	}

	[DispId(13)]
	int Start
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(13)]
		get;
	}

	[DispId(14)]
	float BoundLeft
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(14)]
		get;
	}

	[DispId(15)]
	float BoundTop
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(15)]
		get;
	}

	[DispId(16)]
	float BoundWidth
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(16)]
		get;
	}

	[DispId(17)]
	float BoundHeight
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(17)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(18)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 TrimText();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(19)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 InsertAfter([In][MarshalAs(UnmanagedType.BStr)] string NewText = "");

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(22)]
	void Select();

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(24)]
	void Copy();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(25)]
	void Delete();

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(27)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 PasteSpecial([In] MsoClipboardFormat Format);

	void _VtblGap5_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(31)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 Find([In][MarshalAs(UnmanagedType.BStr)] string FindWhat, [In] int After = 0, [In] MsoTriState MatchCase = MsoTriState.msoFalse, [In] MsoTriState WholeWords = MsoTriState.msoFalse);

	void _VtblGap6_2();

	[DispId(34)]
	MsoLanguageID LanguageID
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(34)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(34)]
		[param: In]
		set;
	}
}
