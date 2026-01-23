using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000CDB05-0000-0000-C000-000000000046")]
[CompilerGenerated]
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

	void _VtblGap2_4();

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

	void _VtblGap4_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809357)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CustomXMLNode SelectSingleNode([In][MarshalAs(UnmanagedType.BStr)] string XPath);
}
