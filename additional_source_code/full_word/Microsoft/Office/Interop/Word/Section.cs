using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[DefaultMember("Range")]
[Guid("00020959-0000-0000-C000-000000000046")]
public interface Section
{
	[DispId(0)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_5();

	[DispId(121)]
	HeadersFooters Headers
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(121)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
