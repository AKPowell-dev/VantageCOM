using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("000C171E-0000-0000-C000-000000000046")]
public interface IMsoLegendKey
{
	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap1_3();

	[DispId(129)]
	IMsoInterior Interior
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_3();

	[DispId(73)]
	int MarkerBackgroundColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap3_2();

	[DispId(75)]
	int MarkerForegroundColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap4_4();

	[DispId(72)]
	XlMarkerStyle MarkerStyle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap5_7();

	[DispId(127)]
	double Left
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(126)]
	double Top
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(122)]
	double Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(123)]
	double Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	void _VtblGap6_2();

	[DispId(1610743841)]
	IMsoChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
