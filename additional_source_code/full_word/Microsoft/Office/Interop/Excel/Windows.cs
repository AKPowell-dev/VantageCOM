using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("00020892-0000-0000-C000-000000000046")]
[InterfaceType(2)]
public interface Windows : IEnumerable
{
	void _VtblGap1_7();

	[IndexerName("_Default")]
	[DispId(0)]
	Window this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
