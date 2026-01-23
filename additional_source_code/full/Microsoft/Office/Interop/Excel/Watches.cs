using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00024456-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
public interface Watches : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Watch Add([In][MarshalAs(UnmanagedType.Struct)] object Source);

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	[IndexerName("_Default")]
	[DispId(0)]
	Watch this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(170)]
	Watch Item
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
