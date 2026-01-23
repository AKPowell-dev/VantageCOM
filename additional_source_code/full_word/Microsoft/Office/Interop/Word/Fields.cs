using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("00020930-0000-0000-C000-000000000046")]
public interface Fields : IEnumerable
{
	void _VtblGap1_7();

	[DispId(0)]
	Field this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(105)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Field Add([In][MarshalAs(UnmanagedType.Interface)] Range Range, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Text, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object PreserveFormatting);
}
