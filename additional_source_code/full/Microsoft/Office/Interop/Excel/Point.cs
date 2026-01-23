using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("0002086A-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
public interface Point
{
	void _VtblGap1_2();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_1();

	[DispId(128)]
	Border Border
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_2();

	[DispId(158)]
	DataLabel DataLabel
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(158)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_3();

	[DispId(77)]
	bool HasDataLabel
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(77)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(77)]
		set;
	}

	void _VtblGap5_7();

	[DispId(75)]
	int MarkerForegroundColor
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(75)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(75)]
		set;
	}

	void _VtblGap6_23();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1922)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ApplyDataLabels([In] XlDataLabelsType Type = XlDataLabelsType.xlDataLabelsShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object LegendKey, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AutoText, [Optional][In][MarshalAs(UnmanagedType.Struct)] object HasLeaderLines, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowSeriesName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowCategoryName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowPercentage, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowBubbleSize, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Separator);

	void _VtblGap7_4();

	[DispId(116)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(116)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_1();

	[DispId(122)]
	double Width
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		get;
	}

	[DispId(126)]
	double Top
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		get;
	}

	[DispId(127)]
	double Left
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		get;
	}
}
