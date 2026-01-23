using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[Guid("0002443A-0000-0000-C000-000000000046")]
[DefaultMember("_Default")]
[CompilerGenerated]
[TypeIdentifier]
public interface Shapes : IEnumerable
{
	void _VtblGap1_3();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(170)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape Item([In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape _Default([In][MarshalAs(UnmanagedType.Struct)] object Index);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap2_19();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(3088)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddChart2([Optional][In][MarshalAs(UnmanagedType.Struct)] object Style, [Optional][In][MarshalAs(UnmanagedType.Struct)] object XlChartType, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Left, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Top, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Width, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Height, [Optional][In][MarshalAs(UnmanagedType.Struct)] object NewLayout);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(3159)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddPicture2([In][MarshalAs(UnmanagedType.BStr)] string Filename, [In] MsoTriState LinkToFile, [In] MsoTriState SaveWithDocument, [In] float Left, [In] float Top, [In] float Width, [In] float Height, [In] MsoPictureCompress Compress);
}
