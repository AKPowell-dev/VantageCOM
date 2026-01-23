using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("0002098E-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface ListLevels : IEnumerable
{
	void _VtblGap1_5();

	[DispId(0)]
	ListLevel this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
