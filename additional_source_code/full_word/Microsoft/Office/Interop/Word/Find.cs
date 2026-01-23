using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("000209B0-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface Find
{
	void _VtblGap1_3();

	[DispId(10)]
	bool Forward
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(10)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(10)]
		[param: In]
		set;
	}

	void _VtblGap2_5();

	[DispId(14)]
	bool MatchCase
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(14)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(14)]
		[param: In]
		set;
	}

	[DispId(15)]
	bool MatchWildcards
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(15)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(15)]
		[param: In]
		set;
	}

	void _VtblGap3_2();

	[DispId(17)]
	bool MatchWholeWord
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(17)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(17)]
		[param: In]
		set;
	}

	void _VtblGap4_8();

	[DispId(22)]
	string Text
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(22)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(22)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap5_6();

	[DispId(27)]
	WdFindWrap Wrap
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(27)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(27)]
		[param: In]
		set;
	}

	void _VtblGap6_9();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(31)]
	void ClearFormatting();

	void _VtblGap7_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(444)]
	bool Execute([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object FindText, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchCase, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchWholeWord, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchWildcards, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchSoundsLike, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchAllWordForms, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Forward, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Wrap, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Format, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ReplaceWith, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Replace, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchKashida, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchDiacritics, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchAlefHamza, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object MatchControl);

	void _VtblGap8_12();

	[DispId(105)]
	bool MatchPrefix
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(105)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(105)]
		[param: In]
		set;
	}

	[DispId(106)]
	bool MatchSuffix
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(106)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(106)]
		[param: In]
		set;
	}
}
