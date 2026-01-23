using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("0002086B-0000-0000-C000-000000000046")]
[InterfaceType(2)]
public interface Series
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

	[DispId(47)]
	XlAxisGroup AxisGroup
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(47)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(47)]
		set;
	}

	[DispId(128)]
	Border Border
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(112)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ClearFormats();

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(157)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object DataLabels([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Delete();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(152)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ErrorBar([In] XlErrorBarDirection Direction, [In] XlErrorBarInclude Include, [In] XlErrorBarType Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Amount, [Optional][In][MarshalAs(UnmanagedType.Struct)] object MinusValues);

	[DispId(159)]
	ErrorBars ErrorBars
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(159)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(182)]
	int Explosion
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(182)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(182)]
		set;
	}

	[DispId(261)]
	string Formula
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(261)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(261)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap3_6();

	[DispId(78)]
	bool HasDataLabels
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(78)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(78)]
		set;
	}

	[DispId(160)]
	bool HasErrorBars
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(160)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(160)]
		set;
	}

	void _VtblGap4_4();

	[DispId(73)]
	int MarkerBackgroundColor
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(73)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(73)]
		set;
	}

	[DispId(74)]
	XlColorIndex MarkerBackgroundColorIndex
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(74)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(74)]
		set;
	}

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

	[DispId(76)]
	XlColorIndex MarkerForegroundColorIndex
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(76)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(76)]
		set;
	}

	[DispId(231)]
	int MarkerSize
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(231)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(231)]
		set;
	}

	[DispId(72)]
	XlMarkerStyle MarkerStyle
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(72)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(72)]
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

	void _VtblGap5_5();

	[DispId(228)]
	int PlotOrder
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(228)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(228)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(70)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Points([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(235)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select();

	[DispId(163)]
	bool Smooth
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(163)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(163)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(154)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Trendlines([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap6_2();

	[DispId(1400)]
	XlChartType ChartType
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1400)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1400)]
		set;
	}

	void _VtblGap7_1();

	[DispId(164)]
	object Values
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(164)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(164)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(1111)]
	object XValues
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1111)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1111)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap8_2();

	[DispId(1403)]
	XlBarShape BarShape
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1403)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1403)]
		set;
	}

	void _VtblGap9_10();

	[DispId(1394)]
	bool HasLeaderLines
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1394)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1394)]
		set;
	}

	[DispId(1666)]
	LeaderLines LeaderLines
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1666)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1922)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ApplyDataLabels([In] XlDataLabelsType Type = XlDataLabelsType.xlDataLabelsShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object LegendKey, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AutoText, [Optional][In][MarshalAs(UnmanagedType.Struct)] object HasLeaderLines, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowSeriesName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowCategoryName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowValue, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowPercentage, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShowBubbleSize, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Separator);

	void _VtblGap10_2();

	[DispId(116)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(116)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_5();

	[DispId(3083)]
	bool IsFiltered
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3083)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3083)]
		set;
	}
}
