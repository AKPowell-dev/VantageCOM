using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("0002445C-0000-0000-C000-000000000046")]
[InterfaceType(2)]
[CompilerGenerated]
public interface Errors
{
	void _VtblGap1_3();

	[DispId(170)]
	Error Item
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[IndexerName("_Default")]
	[DispId(0)]
	Error this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
