using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("000208D6-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _Chart
{
	[DispId(148)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_1();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(304)]
	void Activate();

	void _VtblGap2_5();

	[DispId(486)]
	int Index
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(486)]
		get;
	}

	void _VtblGap3_1();

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap4_19();

	[DispId(558)]
	XlSheetVisibility Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap5_6();

	[DispId(107)]
	bool AutoScaling
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(107)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(107)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(2)]
	[DispId(23)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Axes([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type, [In] XlAxisGroup AxisGroup = XlAxisGroup.xlPrimary);

	void _VtblGap6_4();

	[DispId(80)]
	ChartArea ChartArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(80)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(8)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object ChartGroups([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap7_1();

	[DispId(81)]
	ChartTitle ChartTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(81)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_8();

	[DispId(1395)]
	DataTable DataTable
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1395)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(48)]
	int DepthPercent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(48)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(48)]
		[param: In]
		set;
	}

	void _VtblGap9_1();

	[DispId(93)]
	XlDisplayBlanksAs DisplayBlanksAs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(93)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(93)]
		[param: In]
		set;
	}

	void _VtblGap10_4();

	[DispId(49)]
	int Elevation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(49)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(49)]
		[param: In]
		set;
	}

	void _VtblGap11_7();

	[DispId(52)]
	object HasAxis
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(2)]
		[DispId(52)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(52)]
		[LCIDConversion(2)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
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

	[DispId(53)]
	bool HasLegend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(53)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(53)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(54)]
	bool HasTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(54)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(54)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(55)]
	int HeightPercent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(55)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(55)]
		[param: In]
		set;
	}

	void _VtblGap12_2();

	[DispId(84)]
	Legend Legend
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(84)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap13_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(12)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object LineGroups([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap14_7();

	[DispId(57)]
	int Perspective
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(57)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(57)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap15_3();

	[DispId(85)]
	PlotArea PlotArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(85)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(92)]
	bool PlotVisibleOnly
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(92)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(92)]
		[param: In]
		set;
	}

	void _VtblGap16_2();

	[DispId(58)]
	object RightAngleAxes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(58)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(58)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(59)]
	object Rotation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(59)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(59)]
		[LCIDConversion(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap17_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(68)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object SeriesCollection([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap18_11();

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

	void _VtblGap19_7();

	[DispId(202)]
	XlRowCol PlotBy
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(202)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(202)]
		[param: In]
		set;
	}

	void _VtblGap20_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1413)]
	void SetSourceData([In][MarshalAs(UnmanagedType.Interface)] Range Source, [Optional][In][MarshalAs(UnmanagedType.Struct)] object PlotBy);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1414)]
	bool Export([In][MarshalAs(UnmanagedType.BStr)] string Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object FilterName, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Interactive);

	void _VtblGap21_1();

	[DispId(1814)]
	PivotLayout PivotLayout
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1814)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap22_11();

	[DispId(2504)]
	bool ShowDataLabelsOverMaximum
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2504)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2504)]
		[param: In]
		set;
	}

	void _VtblGap23_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2507)]
	void ApplyChartTemplate([In][MarshalAs(UnmanagedType.BStr)] string Filename);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2508)]
	void SaveChartTemplate([In][MarshalAs(UnmanagedType.BStr)] string Filename);

	void _VtblGap24_2();

	[DispId(2509)]
	object ChartStyle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2509)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2509)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap25_14();

	[DispId(2864)]
	bool ShowAllFieldButtons
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2864)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2864)]
		[param: In]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(3047)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object FullSeriesCollection([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[DispId(3048)]
	XlCategoryLabelLevel CategoryLabelLevel
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3048)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3048)]
		[param: In]
		set;
	}

	[DispId(3049)]
	XlSeriesNameLevel SeriesNameLevel
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3049)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3049)]
		[param: In]
		set;
	}
}
