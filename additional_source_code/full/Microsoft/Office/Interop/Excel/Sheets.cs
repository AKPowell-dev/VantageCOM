using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208D7-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface Sheets : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[LCIDConversion(4)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Add([Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Count, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Type);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	[LCIDConversion(2)]
	void Copy([Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After);

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	[LCIDConversion(0)]
	void Delete();

	void _VtblGap2_1();

	[DispId(170)]
	object Item
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(637)]
	[LCIDConversion(2)]
	void Move([Optional][In][MarshalAs(UnmanagedType.Struct)] object Before, [Optional][In][MarshalAs(UnmanagedType.Struct)] object After);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(235)]
	void Select([Optional][In][MarshalAs(UnmanagedType.Struct)] object Replace);

	void _VtblGap4_4();

	[IndexerName("_Default")]
	[DispId(0)]
	object this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}
}
