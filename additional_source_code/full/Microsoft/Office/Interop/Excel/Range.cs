using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00020846-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
public interface Range : IEnumerable
{
	[DispId(148)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_1();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(304)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Activate();

	void _VtblGap2_2();

	[DispId(236)]
	string Address
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(236)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(441)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ApplyNames([Optional][In][MarshalAs(UnmanagedType.Struct)] object Names, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IgnoreRelativeAbsolute, [Optional][In][MarshalAs(UnmanagedType.Struct)] object UseRowColumnNames, [Optional][In][MarshalAs(UnmanagedType.Struct)] object OmitColumn, [Optional][In][MarshalAs(UnmanagedType.Struct)] object OmitRow, [In] XlApplyNamesOrder Order = XlApplyNamesOrder.xlRowThenColumn, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AppendLast);

	void _VtblGap4_1();

	[DispId(568)]
	Areas Areas
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(568)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(449)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object AutoFill([In][MarshalAs(UnmanagedType.Interface)] Range Destination, [In] XlAutoFillType Type = XlAutoFillType.xlFillDefault);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(793)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object AutoFilter([Optional][In][MarshalAs(UnmanagedType.Struct)] object Field, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Criteria1, [In] XlAutoFilterOperator Operator = XlAutoFilterOperator.xlAnd, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Criteria2, [Optional][In][MarshalAs(UnmanagedType.Struct)] object VisibleDropDown);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(237)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object AutoFit();

	void _VtblGap6_3();

	[DispId(435)]
	Borders Borders
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(435)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_1();

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(603)]
	Characters Characters
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(603)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(111)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Clear();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(113)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ClearContents();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(112)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ClearFormats();

	void _VtblGap9_2();

	[DispId(240)]
	int Column
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(240)]
		get;
	}

	void _VtblGap10_1();

	[DispId(241)]
	Range Columns
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(241)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(242)]
	object ColumnWidth
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(242)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(242)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap11_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Copy([Optional][In][MarshalAs(UnmanagedType.Struct)] object Destination);

	void _VtblGap12_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(213)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object CopyPicture([In] XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, [In] XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	void _VtblGap13_2();

	[DispId(501)]
	Range CurrentArray
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(501)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(243)]
	Range CurrentRegion
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(243)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(565)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Cut([Optional][In][MarshalAs(UnmanagedType.Struct)] object Destination);

	void _VtblGap14_1();

	[IndexerName("_Default")]
	[DispId(0)]
	object this[[Optional][In][MarshalAs(UnmanagedType.Struct)] object RowIndex, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ColumnIndex]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Delete([Optional][In][MarshalAs(UnmanagedType.Struct)] object Shift);

	void _VtblGap15_2();

	[DispId(545)]
	Range DirectDependents
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(545)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(546)]
	Range DirectPrecedents
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(546)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap16_1();

	[DispId(500)]
	Range End
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(500)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(246)]
	Range EntireColumn
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(246)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(247)]
	Range EntireRow
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(247)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(248)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object FillDown();

	void _VtblGap17_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(250)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object FillRight();

	void _VtblGap18_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(398)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Find([In][MarshalAs(UnmanagedType.Struct)] object What, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After, [Optional][In][MarshalAs(UnmanagedType.Struct)] object LookIn, [Optional][In][MarshalAs(UnmanagedType.Struct)] object LookAt, [Optional][In][MarshalAs(UnmanagedType.Struct)] object SearchOrder, [In] XlSearchDirection SearchDirection = XlSearchDirection.xlNext, [Optional][In][MarshalAs(UnmanagedType.Struct)] object MatchCase, [Optional][In][MarshalAs(UnmanagedType.Struct)] object MatchByte, [Optional][In][MarshalAs(UnmanagedType.Struct)] object SearchFormat);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(399)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range FindNext([Optional][In][MarshalAs(UnmanagedType.Struct)] object After);

	void _VtblGap19_1();

	[DispId(146)]
	Font Font
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(146)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(261)]
	object Formula
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(261)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(261)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(586)]
	object FormulaArray
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(586)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(586)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap20_4();

	[DispId(263)]
	object FormulaLocal
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(263)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(263)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(264)]
	object FormulaR1C1
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(264)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(264)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap21_4();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(46)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Group([Optional][In][MarshalAs(UnmanagedType.Struct)] object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] object End, [Optional][In][MarshalAs(UnmanagedType.Struct)] object By, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Periods);

	[DispId(266)]
	object HasArray
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(266)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[DispId(267)]
	object HasFormula
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(267)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[DispId(123)]
	object Height
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[DispId(268)]
	object Hidden
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(268)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(268)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(136)]
	object HorizontalAlignment
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(136)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(136)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(201)]
	object IndentLevel
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(201)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(201)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap22_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(252)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Insert([Optional][In][MarshalAs(UnmanagedType.Struct)] object Shift, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CopyOrigin);

	[DispId(129)]
	Interior Interior
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(129)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(170)]
	object Item
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap23_1();

	[DispId(127)]
	object Left
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap24_6();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1384)]
	void UnMerge();

	[DispId(1385)]
	Range MergeArea
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1385)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(208)]
	object MergeCells
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(208)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(208)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(110)]
	object Name
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1032)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object NavigateArrow([Optional][In][MarshalAs(UnmanagedType.Struct)] object TowardPrecedent, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ArrowNumber, [Optional][In][MarshalAs(UnmanagedType.Struct)] object LinkNumber);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap25_2();

	[DispId(193)]
	object NumberFormat
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(193)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(193)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(1097)]
	object NumberFormatLocal
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1097)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1097)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(254)]
	Range Offset
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(254)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(134)]
	object Orientation
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(134)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(134)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(271)]
	object OutlineLevel
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(271)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(271)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap26_6();

	[DispId(716)]
	PivotTable PivotTable
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(716)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap27_1();

	[DispId(504)]
	object PrefixCharacter
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(504)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap28_7();

	[DispId(256)]
	Range Resize
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(256)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(257)]
	int Row
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(257)]
		get;
	}

	void _VtblGap29_1();

	[DispId(272)]
	object RowHeight
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(272)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(272)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(258)]
	Range Rows
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(258)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap30_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(235)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select();

	void _VtblGap31_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(877)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ShowDependents([Optional][In][MarshalAs(UnmanagedType.Struct)] object Remove);

	[DispId(585)]
	object ShowDetail
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(585)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(585)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap32_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(879)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ShowPrecedents([Optional][In][MarshalAs(UnmanagedType.Struct)] object Remove);

	void _VtblGap33_5();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(410)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range SpecialCells([In] XlCellType Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Value);

	[DispId(260)]
	object Style
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(260)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(260)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap34_4();

	[DispId(138)]
	object Text
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(138)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap35_1();

	[DispId(126)]
	object Top
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(244)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Ungroup();

	void _VtblGap36_4();

	[DispId(1387)]
	Validation Validation
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1387)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(6)]
	object Value
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		[param: Optional]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(1388)]
	object Value2
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1388)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1388)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(137)]
	object VerticalAlignment
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(137)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(137)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[DispId(122)]
	object Width
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[DispId(348)]
	Worksheet Worksheet
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(348)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(276)]
	object WrapText
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(276)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(276)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1389)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Comment AddComment([Optional][In][MarshalAs(UnmanagedType.Struct)] object Text);

	[DispId(910)]
	Comment Comment
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(910)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1390)]
	void ClearComments();

	void _VtblGap37_1();

	[DispId(1392)]
	FormatConditions FormatConditions
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1392)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap38_2();

	[DispId(1393)]
	Hyperlinks Hyperlinks
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1393)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap39_7();

	[DispId(2015)]
	Errors Errors
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2015)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap40_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1928)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object PasteSpecial([In] XlPasteType Paste = XlPasteType.xlPasteAll, [In] XlPasteSpecialOperation Operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, [Optional][In][MarshalAs(UnmanagedType.Struct)] object SkipBlanks, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Transpose);

	void _VtblGap41_1();

	[DispId(2257)]
	ListObject ListObject
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2257)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap42_6();

	[DispId(2499)]
	object CountLarge
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2499)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap43_1();

	[DispId(2853)]
	SparklineGroups SparklineGroups
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2853)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2854)]
	void ClearHyperlinks();
}
