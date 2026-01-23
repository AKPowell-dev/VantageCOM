using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208BB-0000-0000-C000-000000000046")]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
[DefaultMember("_Default")]
public interface LegendEntries : IEnumerable
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
	LegendEntry Item([In][MarshalAs(UnmanagedType.Struct)] object Index);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	LegendEntry _Default([In][MarshalAs(UnmanagedType.Struct)] object Index);
}
