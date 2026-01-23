using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("914934D6-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
[TypeIdentifier]
public interface Designs : Collection
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap1_1();

	[DispId(11)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		get;
	}

	void _VtblGap2_2();

	[DispId(0)]
	Design this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2003)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Design Add([In][MarshalAs(UnmanagedType.BStr)] string designName, [In] int Index = -1);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2005)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Design Clone([In][MarshalAs(UnmanagedType.Interface)] Design pOriginal, [In] int Index = -1);
}
