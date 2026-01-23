using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[Guid("000208B8-0000-0000-C000-000000000046")]
[DefaultMember("_Default")]
[TypeIdentifier]
[CompilerGenerated]
public interface Names : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Name Add([Optional][In][MarshalAs(UnmanagedType.Struct)] object Name, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersTo, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Visible, [Optional][In][MarshalAs(UnmanagedType.Struct)] object MacroType, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ShortcutKey, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Category, [Optional][In][MarshalAs(UnmanagedType.Struct)] object NameLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersToLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CategoryLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersToR1C1, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersToR1C1Local);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(170)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Name Item([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IndexLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersTo);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Name _Default([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index, [Optional][In][MarshalAs(UnmanagedType.Struct)] object IndexLocal, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RefersTo);

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();
}
