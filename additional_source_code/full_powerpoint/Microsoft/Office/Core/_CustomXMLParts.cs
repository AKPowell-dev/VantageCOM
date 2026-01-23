using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000CDB09-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
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
}
