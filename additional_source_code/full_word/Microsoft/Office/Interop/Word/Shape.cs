using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("000209A0-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface Shape
{
	void _VtblGap1_4();

	[DispId(101)]
	MsoAutoShapeType AutoShapeType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[param: In]
		set;
	}

	void _VtblGap2_4();

	[DispId(107)]
	FillFormat Fill
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(107)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(108)]
	GroupShapes GroupItems
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(108)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(109)]
	float Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(109)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(109)]
		[param: In]
		set;
	}

	void _VtblGap3_1();

	[DispId(111)]
	float Left
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		[param: In]
		set;
	}

	[DispId(112)]
	LineFormat Line
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(112)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(113)]
	MsoTriState LockAspectRatio
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		[param: In]
		set;
	}

	[DispId(115)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(115)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(115)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap4_1();

	[DispId(117)]
	float Rotation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		[param: In]
		set;
	}

	void _VtblGap5_3();

	[DispId(121)]
	TextFrame TextFrame
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(121)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_1();

	[DispId(123)]
	float Top
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		[param: In]
		set;
	}

	[DispId(124)]
	MsoShapeType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(124)]
		get;
	}

	void _VtblGap7_4();

	[DispId(128)]
	float Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		[param: In]
		set;
	}

	[DispId(129)]
	int ZOrderPosition
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(129)]
		get;
	}

	void _VtblGap8_1();

	[DispId(300)]
	WdRelativeHorizontalPosition RelativeHorizontalPosition
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(300)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(300)]
		[param: In]
		set;
	}

	[DispId(301)]
	WdRelativeVerticalPosition RelativeVerticalPosition
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(301)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(301)]
		[param: In]
		set;
	}

	void _VtblGap9_2();

	[DispId(303)]
	WrapFormat WrapFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(303)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_1();

	[DispId(501)]
	Range Anchor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(501)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(11)]
	void Delete();

	void _VtblGap12_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(19)]
	void ScaleHeight([In] float Factor, [In] MsoTriState RelativeToOriginalSize, [In] MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(20)]
	void ScaleWidth([In] float Factor, [In] MsoTriState RelativeToOriginalSize, [In] MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(21)]
	void Select([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Replace);

	void _VtblGap13_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(24)]
	void ZOrder([In] MsoZOrderCmd ZOrderCmd);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(25)]
	[return: MarshalAs(UnmanagedType.Interface)]
	InlineShape ConvertToInlineShape();

	void _VtblGap14_2();

	[DispId(131)]
	string AlternativeText
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(131)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(131)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap15_16();

	[DispId(148)]
	MsoTriState HasChart
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		get;
	}

	[DispId(149)]
	Chart Chart
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(149)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(200)]
	float LeftRelative
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(200)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(200)]
		[param: In]
		set;
	}

	[DispId(201)]
	float TopRelative
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(201)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(201)]
		[param: In]
		set;
	}

	[DispId(202)]
	float WidthRelative
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(202)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(202)]
		[param: In]
		set;
	}

	[DispId(203)]
	float HeightRelative
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(203)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(203)]
		[param: In]
		set;
	}

	void _VtblGap16_5();

	[DispId(153)]
	GlowFormat Glow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(153)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap17_1();

	[DispId(155)]
	TextFrame2 TextFrame2
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(155)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(206)]
	MsoTriState HasSmartArt
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(206)]
		get;
	}

	[DispId(156)]
	SmartArt SmartArt
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(156)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
