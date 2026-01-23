using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("Range")]
[TypeIdentifier]
[CompilerGenerated]
[Guid("00020985-0000-0000-C000-000000000046")]
public interface HeaderFooter
{
	void _VtblGap1_3();

	[DispId(0)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
