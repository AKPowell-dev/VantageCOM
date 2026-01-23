using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020990-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface ListTemplates : IEnumerable
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap1_4();

	[DispId(0)]
	ListTemplate this[[In][MarshalAs(UnmanagedType.Struct)] ref object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(100)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ListTemplate Add([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OutlineNumbered, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Name);
}
