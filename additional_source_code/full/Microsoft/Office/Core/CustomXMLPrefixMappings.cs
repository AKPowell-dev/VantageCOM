using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000CDB00-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface CustomXMLPrefixMappings : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_4();

	[DispId(0)]
	CustomXMLPrefixMapping this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809349)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string LookupPrefix([In][MarshalAs(UnmanagedType.BStr)] string NamespaceURI);
}
