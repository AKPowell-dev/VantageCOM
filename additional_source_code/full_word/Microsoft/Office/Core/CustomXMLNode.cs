using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("000CDB04-0000-0000-C000-000000000046")]
public interface CustomXMLNode : _IMsoDispObj
{
	void _VtblGap1_6();

	[DispId(1610809348)]
	CustomXMLNode FirstChild
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809348)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_10();

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
}
