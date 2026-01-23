using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("0002085A-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
public interface ChartGroups : IEnumerable
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
	ChartGroup Item([In][MarshalAs(UnmanagedType.Struct)] object Index);
}
