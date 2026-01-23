using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("0002086C-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
[DefaultMember("_Default")]
public interface SeriesCollection : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Series Add([In][MarshalAs(UnmanagedType.Struct)] object Source, [In] XlRowCol Rowcol = (XlRowCol)(-4105), [Optional][In][MarshalAs(UnmanagedType.Struct)] object SeriesLabels, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CategoryLabels, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(170)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Series Item([In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1117)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Series NewSeries();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Series _Default([In][MarshalAs(UnmanagedType.Struct)] object Index);
}
