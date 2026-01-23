using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("0002097F-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface AddIns : IEnumerable
{
	void _VtblGap1_5();

	[DispId(0)]
	AddIn this[[In][MarshalAs(UnmanagedType.Struct)] ref object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2)]
	[return: MarshalAs(UnmanagedType.Interface)]
	AddIn Add([In][MarshalAs(UnmanagedType.BStr)] string FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Install);
}
