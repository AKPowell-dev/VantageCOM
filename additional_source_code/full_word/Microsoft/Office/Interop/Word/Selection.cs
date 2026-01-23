using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("00020975-0000-0000-C000-000000000046")]
[TypeIdentifier]
[DefaultMember("Text")]
public interface Selection
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

	[DispId(2)]
	Range FormattedText
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

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

	void _VtblGap1_2();

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
	WdSelectionType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		get;
	}

	void _VtblGap2_3();

	[DispId(50)]
	Tables Tables
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(50)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_6();

	[DispId(57)]
	Cells Cells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(57)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_1();

	[DispId(59)]
	Paragraphs Paragraphs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(59)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_3();

	[DispId(64)]
	Fields Fields
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(64)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_6();

	[DispId(75)]
	Bookmarks Bookmarks
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(75)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_10();

	[DispId(306)]
	HeaderFooter HeaderFooter
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(306)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_4();

	[DispId(400)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(400)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(401)]
	object Information
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(401)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap9_12();

	[DispId(411)]
	InlineShapes InlineShapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(411)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_3();

	[DispId(1003)]
	Document Document
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1004)]
	ShapeRange ShapeRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(101)]
	void Collapse([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Direction);

	void _VtblGap12_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(104)]
	void InsertAfter([In][MarshalAs(UnmanagedType.BStr)] string Text);

	void _VtblGap13_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(110)]
	int MoveStart([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	void _VtblGap14_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(113)]
	int MoveStartWhile([In][MarshalAs(UnmanagedType.Struct)] ref object Cset, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(114)]
	int MoveEndWhile([In][MarshalAs(UnmanagedType.Struct)] ref object Cset, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count);

	void _VtblGap15_9();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(126)]
	bool InRange([In][MarshalAs(UnmanagedType.Interface)] Range Range);

	void _VtblGap16_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(129)]
	int Expand([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit);

	void _VtblGap17_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(164)]
	void InsertSymbol([In] int CharacterNumber, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Font, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unicode, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Bias);

	void _VtblGap18_8();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(173)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range GoTo([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object What, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Which, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Name);

	void _VtblGap19_9();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(500)]
	int MoveLeft([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Extend);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(501)]
	int MoveRight([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Unit, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Count, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Extend);

	void _VtblGap20_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(507)]
	void TypeText([In][MarshalAs(UnmanagedType.BStr)] string Text);

	void _VtblGap21_49();

	[DispId(1021)]
	ShapeRange ChildShapeRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1021)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap22_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1013)]
	void PasteAndFormat([In] WdRecoveryType Type);
}
