using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[DefaultMember("_Default")]
[InterfaceType(2)]
[TypeIdentifier]
[Guid("000208B8-0000-0000-C000-000000000046")]
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

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();
}
