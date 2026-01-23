using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("0002099F-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface Shapes : IEnumerable
{
	void _VtblGap1_3();

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	[DispId(0)]
	Shape this[[In][MarshalAs(UnmanagedType.Struct)] ref object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(15)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddPicture([In][MarshalAs(UnmanagedType.BStr)] string FileName, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object LinkToFile, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveWithDocument, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Left, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Top, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Width, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Height, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Anchor);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(17)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddShape([In] int Type, [In] float Left, [In] float Top, [In] float Width, [In] float Height, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Anchor);
}
