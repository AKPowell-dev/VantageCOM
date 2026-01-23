using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[TypeIdentifier]
[Guid("0002440F-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface ChartEvents
{
	void _VtblGap1_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1530)]
	void Deactivate();

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1532)]
	void MouseUp([In] int Button, [In] int Shift, [In] int x, [In] int y);
}
