using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("92D41A55-F07E-4CA4-AF6F-BEF486AA4E6F")]
[CompilerGenerated]
public interface Chart
{
	void _VtblGap1_2();

	[DispId(1400)]
	XlChartType ChartType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1400)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1400)]
		[param: In]
		set;
	}

	[DispId(1396)]
	bool HasDataTable
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1396)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1396)]
		[param: In]
		set;
	}

	void _VtblGap2_7();

	[DispId(2003)]
	DataTable DataTable
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	void ApplyChartTemplate([In][MarshalAs(UnmanagedType.BStr)] string FileName);

	void _VtblGap4_12();

	[DispId(2011)]
	ChartData ChartData
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2011)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2012)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2012)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_19();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2016)]
	[LCIDConversion(2)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Axes([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type, [In] XlAxisGroup AxisGroup = XlAxisGroup.xlPrimary);

	[DispId(2017)]
	ChartArea ChartArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(2017)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2018)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object ChartGroups([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[DispId(2019)]
	ChartTitle ChartTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2019)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_15();

	[DispId(2031)]
	object HasAxis
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2031)]
		[LCIDConversion(2)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(2)]
		[DispId(2031)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(2032)]
	bool HasLegend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2032)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2032)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(2033)]
	bool HasTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(2033)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(2033)]
		[param: In]
		set;
	}

	void _VtblGap7_2();

	[DispId(2035)]
	Legend Legend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2035)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_4();

	[DispId(2038)]
	PlotArea PlotArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2038)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap9_6();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(2042)]
	void Select([Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(2043)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object SeriesCollection([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);
}
