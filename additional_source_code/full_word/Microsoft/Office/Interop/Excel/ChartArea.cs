using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[InterfaceType(2)]
[Guid("000208CC-0000-0000-C000-000000000046")]
public interface ChartArea
{
	void _VtblGap1_4();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(235)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Select();

	void _VtblGap2_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(551)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Copy();
}
