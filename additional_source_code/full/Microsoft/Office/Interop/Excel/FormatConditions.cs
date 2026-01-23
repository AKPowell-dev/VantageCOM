using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00024424-0000-0000-C000-000000000046")]
[CompilerGenerated]
[InterfaceType(2)]
[TypeIdentifier]
public interface FormatConditions : IEnumerable
{
	void _VtblGap1_3();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(170)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Item([In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object Add([In] XlFormatConditionType Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Operator, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Formula1, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Formula2, [Optional][In][MarshalAs(UnmanagedType.Struct)] object String, [Optional][In][MarshalAs(UnmanagedType.Struct)] object TextOperator, [Optional][In][MarshalAs(UnmanagedType.Struct)] object DateOperator, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ScopeType);

	[IndexerName("_Default")]
	[DispId(0)]
	object this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	void Delete();
}
