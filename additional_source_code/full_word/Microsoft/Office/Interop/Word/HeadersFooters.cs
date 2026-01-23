using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020984-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface HeadersFooters : IEnumerable
{
	void _VtblGap1_5();

	[DispId(0)]
	HeaderFooter this[[In] WdHeaderFooterIndex Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
