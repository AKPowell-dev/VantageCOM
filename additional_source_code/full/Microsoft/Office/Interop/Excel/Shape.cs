using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00024439-0000-0000-C000-000000000046")]
[CompilerGenerated]
[InterfaceType(2)]
[TypeIdentifier]
public interface Shape
{
	[DispId(148)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	void Delete();

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1676)]
	void Flip([In] MsoFlipCmd FlipCmd);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1680)]
	void IncrementRotation([In] float Increment);

	void _VtblGap4_5();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(235)]
	void Select([Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);

	void _VtblGap5_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(622)]
	void ZOrder([In] MsoZOrderCmd ZOrderCmd);

	void _VtblGap6_1();

	[DispId(1692)]
	TextFrame TextFrame
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1692)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_6();

	[DispId(1663)]
	FillFormat Fill
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1663)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1698)]
	GroupShapes GroupItems
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1698)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(123)]
	float Height
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		set;
	}

	void _VtblGap8_1();

	[DispId(127)]
	float Left
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		set;
	}

	[DispId(817)]
	LineFormat Line
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(817)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1700)]
	MsoTriState LockAspectRatio
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1700)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1700)]
		set;
	}

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap9_7();

	[DispId(126)]
	float Top
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		set;
	}

	[DispId(108)]
	MsoShapeType Type
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(108)]
		get;
	}

	void _VtblGap10_2();

	[DispId(558)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		set;
	}

	[DispId(122)]
	float Width
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		set;
	}

	void _VtblGap11_9();

	[DispId(620)]
	Range TopLeftCell
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(620)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(615)]
	Range BottomRightCell
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(615)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(617)]
	XlPlacement Placement
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(617)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(617)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	void Copy();

	void _VtblGap12_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(213)]
	void CopyPicture([Optional][In][MarshalAs(UnmanagedType.Struct)] object Appearance, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Format);

	void _VtblGap13_4();

	[DispId(1891)]
	string AlternativeText
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1891)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1891)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap14_13();

	[DispId(7)]
	Chart Chart
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(7)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2658)]
	MsoTriState HasChart
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2658)]
		get;
	}
}
