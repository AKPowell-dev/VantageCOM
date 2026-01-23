using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("00020846-0000-0000-C000-000000000046")]
[CompilerGenerated]
[InterfaceType(2)]
public interface Range : IEnumerable
{
	void _VtblGap1_21();

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(111)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Clear();

	void _VtblGap3_6();

	[DispId(241)]
	Range Columns
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(241)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Copy([Optional][In][MarshalAs(UnmanagedType.Struct)] object Destination);

	void _VtblGap5_2();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	void _VtblGap6_6();

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

	void _VtblGap7_36();

	[DispId(123)]
	object Height
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap8_59();

	[DispId(258)]
	Range Rows
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(258)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap9_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(235)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select();

	void _VtblGap10_18();

	[DispId(138)]
	object Text
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(138)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	void _VtblGap11_10();

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

	void _VtblGap12_2();

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
}
