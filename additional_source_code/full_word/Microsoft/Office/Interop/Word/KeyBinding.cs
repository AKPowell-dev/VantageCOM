using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020998-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface KeyBinding
{
	void _VtblGap1_11();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(101)]
	void Clear();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(102)]
	void Disable();
}
