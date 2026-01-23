using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("6FFA84BB-A350-4442-BB53-A43653459A84")]
[TypeIdentifier]
public interface Chart
{
	void _VtblGap1_1();

	[DispId(1610743809)]
	bool HasTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743809)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743809)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(1610743811)]
	ChartTitle ChartTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1610743811)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_16();

	[DispId(8)]
	object ChartGroups
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(1)]
		[DispId(8)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(68)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object SeriesCollection([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap3_8();

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

	void _VtblGap4_5();

	[DispId(53)]
	bool HasLegend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(53)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(53)]
		[param: In]
		set;
	}

	[DispId(84)]
	Legend Legend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(84)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(2)]
	[DispId(1610743861)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Axes([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type, [In] XlAxisGroup AxisGroup = XlAxisGroup.xlPrimary);

	[DispId(1610743862)]
	object HasAxis
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743862)]
		[LCIDConversion(2)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(2)]
		[DispId(1610743862)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap5_2();

	[DispId(1610743866)]
	PlotArea PlotArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743866)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap6_2();

	[DispId(1610743869)]
	ChartArea ChartArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1610743869)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_6();

	[DispId(1610743876)]
	DataTable DataTable
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743876)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610743884)]
	void ApplyChartTemplate([In][MarshalAs(UnmanagedType.BStr)] string FileName);

	void _VtblGap9_14();

	[DispId(1610743901)]
	ChartData ChartData
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743901)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_19();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(235)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select([Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);
}
