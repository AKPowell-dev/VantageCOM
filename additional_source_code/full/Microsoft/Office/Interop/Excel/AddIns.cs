using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[InterfaceType(2)]
[TypeIdentifier]
[Guid("00020858-0000-0000-C000-000000000046")]
public interface AddIns : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	AddIn Add([In][MarshalAs(UnmanagedType.BStr)] string Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CopyFile);

	void _VtblGap2_1();

	[DispId(170)]
	AddIn Item
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(170)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_1();

	[IndexerName("_Default")]
	[DispId(0)]
	AddIn this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
