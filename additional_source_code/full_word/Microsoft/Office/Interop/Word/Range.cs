using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("0002095E-0000-0000-C000-000000000046")]
[TypeIdentifier]
[DefaultMember("Text")]
[CompilerGenerated]
public interface Range
{
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

	void _VtblGap1_2();

	[DispId(3)]
	int Start
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[param: In]
		set;
	}

	[DispId(4)]
	int End
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(4)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(4)]
		[param: In]
		set;
	}

	[DispId(5)]
	Font Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(5)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(5)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	[DispId(6)]
	Range Duplicate
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(7)]
	WdStoryType StoryType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(7)]
		get;
	}

	[DispId(50)]
	Tables Tables
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(50)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(51)]
	Words Words
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(51)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_1();

	[DispId(53)]
	Characters Characters
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(53)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_5();

	[DispId(59)]
	Paragraphs Paragraphs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(59)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1100)]
	Borders Borders
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1100)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1100)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	void _VtblGap4_8();

	[DispId(68)]
	ListFormat ListFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_1();

	[DispId(1000)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1000)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_1();

	[DispId(1002)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1002)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap7_14();

	[DispId(153)]
	WdLanguageID LanguageID
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(153)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(153)]
		[param: In]
		set;
	}

	void _VtblGap8_8();

	[DispId(301)]
	WdColorIndex HighlightColorIndex
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(301)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(301)]
		[param: In]
		set;
	}

	void _VtblGap9_7();

	[DispId(262)]
	Find Find
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(262)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_2();

	[DispId(311)]
	ShapeRange ShapeRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(311)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_2();

	[DispId(313)]
	object Information
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(313)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap12_5();

	[DispId(319)]
	InlineShapes InlineShapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(319)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(320)]
	Range NextStoryRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(320)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap13_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(65535)]
	void Select();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(100)]
	void SetRange([In] int Start, [In] int End);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(101)]
	void Collapse([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Direction);

	void _VtblGap14_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(105)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Next([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(106)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Previous([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	void _VtblGap15_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(114)]
	int MoveEndWhile([In][MarshalAs(UnmanagedType.Struct)] ref object Cset, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	void _VtblGap16_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(120)]
	void Copy();

	void _VtblGap17_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(126)]
	bool InRange([In][MarshalAs(UnmanagedType.Interface)] Range Range);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(127)]
	int Delete([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	void _VtblGap18_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(129)]
	int Expand([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit);

	void _VtblGap19_11();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(171)]
	bool IsEqual([In][MarshalAs(UnmanagedType.Interface)] Range Range);

	void _VtblGap20_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(176)]
	void PasteSpecial([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object IconIndex, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Link, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Placement, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DisplayAsIcon, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DataType, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object IconFileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object IconLabel);

	void _VtblGap21_46();

	[DispId(405)]
	string ID
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(405)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(405)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap22_4();

	[DispId(409)]
	Document Document
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(409)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap23_19();

	[DispId(424)]
	ContentControls ContentControls
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(424)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
