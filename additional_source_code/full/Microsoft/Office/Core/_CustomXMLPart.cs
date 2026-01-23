using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000CDB05-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface _CustomXMLPart : _IMsoDispObj
{
	void _VtblGap1_4();

	[DispId(1610809346)]
	string Id
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809346)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(1610809347)]
	string NamespaceURI
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809347)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap2_2();

	[DispId(1610809350)]
	CustomXMLPrefixMappings NamespaceManager
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809350)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1610809351)]
	string XML
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809351)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809353)]
	void Delete();

	void _VtblGap4_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809356)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLNodes SelectNodes([In][MarshalAs(UnmanagedType.BStr)] string XPath);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809357)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLNode SelectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string XPath);
}
