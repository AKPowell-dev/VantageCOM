using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[DefaultMember("_Default")]
[CompilerGenerated]
[Guid("000208B8-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface Names : IEnumerable
{
	void _VtblGap1_4();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(170)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Name Item([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IndexLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersTo);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Name _Default([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IndexLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersTo);
}
