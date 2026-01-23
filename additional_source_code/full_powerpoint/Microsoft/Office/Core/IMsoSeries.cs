using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C170B-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface IMsoSeries
{
	void _VtblGap1_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object DataLabels([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap2_2();

	[DispId(159)]
	IMsoErrorBars ErrorBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_10();

	[DispId(78)]
	bool HasDataLabels
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	[DispId(160)]
	bool HasErrorBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap4_4();

	[DispId(73)]
	int MarkerBackgroundColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap5_2();

	[DispId(75)]
	int MarkerForegroundColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	void _VtblGap6_4();

	[DispId(72)]
	XlMarkerStyle MarkerStyle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap7_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Points([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select();

	void _VtblGap8_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Trendlines([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap9_23();

	[DispId(1394)]
	bool HasLeaderLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[param: In]
		set;
	}

	[DispId(1666)]
	IMsoLeaderLines LeaderLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ApplyDataLabels([In] XlDataLabelsType Type = XlDataLabelsType.xlDataLabelsShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IMsoLegendKey, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AutoText, [Optional][In][MarshalAs(UnmanagedType.Struct)] object HasLeaderLines, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowSeriesName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowCategoryName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowPercentage, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowBubbleSize, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Separator);

	[DispId(1610743890)]
	IMsoChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
