using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("00020967-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface Bookmarks : IEnumerable
{
	void _VtblGap1_9();

	[DispId(0)]
	Bookmark this[[In][MarshalAs(UnmanagedType.Struct)] ref object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
