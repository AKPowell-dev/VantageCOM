using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000CDB09-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _CustomXMLParts : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_4();

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
}
