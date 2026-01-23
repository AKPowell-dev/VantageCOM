using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208DB-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface Workbooks : IEnumerable
{
	void _VtblGap1_5();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap3_2();

	[IndexerName("_Default")]
	[DispId(0)]
	Workbook this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(15)]
	[DispId(1923)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Workbook Open([In][MarshalAs(UnmanagedType.BStr)] string Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object UpdateLinks, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ReadOnly, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Format, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] object WriteResPassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IgnoreReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Origin, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Delimiter, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Editable, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Notify, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Converter, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AddToMru, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Local, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CorruptLoad);
}
