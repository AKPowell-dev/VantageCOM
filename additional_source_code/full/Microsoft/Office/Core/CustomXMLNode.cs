using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000CDB04-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface CustomXMLNode : _IMsoDispObj
{
	void _VtblGap1_3();

	[DispId(1610809345)]
	CustomXMLNodes Attributes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809345)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_2();

	[DispId(1610809348)]
	CustomXMLNode FirstChild
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809348)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1610809349)]
	CustomXMLNode LastChild
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809349)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_9();

	[DispId(1610809359)]
	string Text
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809359)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809359)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap4_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809363)]
	void AppendChildNode([In][MarshalAs(UnmanagedType.BStr)] string Name = "", [In][MarshalAs(UnmanagedType.BStr)] string NamespaceURI = "", [In] MsoCustomXMLNodeType NodeType = MsoCustomXMLNodeType.msoCustomXMLNodeElement, [In][MarshalAs(UnmanagedType.BStr)] string NodeValue = "");

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809365)]
	void Delete();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809366)]
	bool HasChildNodes();

	void _VtblGap6_6();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809373)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLNode SelectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string XPath);
}
