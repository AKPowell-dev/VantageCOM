using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("000208D8-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _Worksheet
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

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(2)]
	[DispId(551)]
	void Copy([Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(117)]
	void Delete();

	void _VtblGap2_3();

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

	void _VtblGap4_7();

	[DispId(998)]
	PageSetup PageSetup
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(998)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_4();

	[DispId(292)]
	bool ProtectContents
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(292)]
		get;
	}

	void _VtblGap6_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(235)]
	void Select([Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(285)]
	void Unprotect([Optional][In][MarshalAs(UnmanagedType.Struct)] object Password);

	[DispId(558)]
	XlSheetVisibility Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(558)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(1377)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1377)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_3();

	[DispId(792)]
	bool AutoFilterMode
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(792)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(792)]
		[param: In]
		set;
	}

	void _VtblGap8_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(279)]
	void Calculate();

	[DispId(1424)]
	bool EnableCalculation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1424)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1424)]
		[param: In]
		set;
	}

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1060)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object ChartObjects([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(824)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object CheckBoxes([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap9_1();

	[DispId(1069)]
	Range CircularReference
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1069)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(970)]
	void ClearArrows();

	[DispId(241)]
	Range Columns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(241)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap10_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(836)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object DropDowns([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap11_8();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(1)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Evaluate([In][MarshalAs(UnmanagedType.Struct)] object Name);

	void _VtblGap12_1();

	[DispId(800)]
	bool FilterMode
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(800)]
		get;
	}

	void _VtblGap13_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(799)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object OLEObjects([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap14_7();

	[DispId(102)]
	Outline Outline
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(102)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap15_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(2)]
	[DispId(211)]
	void Paste([Optional][In][MarshalAs(UnmanagedType.Struct)] object Destination, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Link);

	void _VtblGap16_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(690)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object PivotTables([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap17_1();

	[DispId(197)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(197)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap18_1();

	[DispId(258)]
	Range Rows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(258)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap19_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(794)]
	void ShowAllData();

	void _VtblGap20_2();

	[DispId(407)]
	double StandardHeight
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(407)]
		get;
	}

	[DispId(408)]
	double StandardWidth
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(408)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(408)]
		[param: In]
		set;
	}

	void _VtblGap21_1();

	[DispId(402)]
	bool TransitionFormEntry
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(402)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(402)]
		[param: In]
		set;
	}

	void _VtblGap22_1();

	[DispId(412)]
	Range UsedRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(412)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap23_2();

	[DispId(1434)]
	QueryTables QueryTables
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1434)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1435)]
	bool DisplayPageBreaks
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1435)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1435)]
		[param: In]
		set;
	}

	[DispId(575)]
	Comments Comments
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(575)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1393)]
	Hyperlinks Hyperlinks
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1393)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap24_4();

	[DispId(793)]
	AutoFilter AutoFilter
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(793)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap25_13();

	[DispId(2259)]
	ListObjects ListObjects
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2259)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap26_6();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2493)]
	void ExportAsFixedFormat([In] XlFixedFormatType Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Quality, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IncludeDocProperties, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IgnorePrintAreas, [Optional][In][MarshalAs(UnmanagedType.Struct)] object From, [Optional][In][MarshalAs(UnmanagedType.Struct)] object To, [Optional][In][MarshalAs(UnmanagedType.Struct)] object OpenAfterPublish, [Optional][In][MarshalAs(UnmanagedType.Struct)] object FixedFormatExtClassPtr);
}
