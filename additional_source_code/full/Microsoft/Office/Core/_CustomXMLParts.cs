using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("000CDB09-0000-0000-C000-000000000046")]
public interface _CustomXMLParts : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_3();

	[DispId(1610809345)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809345)]
		get;
	}

	[DispId(0)]
	CustomXMLPart this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809347)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLPart Add([In][MarshalAs(UnmanagedType.BStr)] string XML = "", [Optional][In][MarshalAs(UnmanagedType.Struct)] object SchemaCollection);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809348)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLPart SelectByID([In][MarshalAs(UnmanagedType.BStr)] string Id);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();
}
