using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("000208AB-0000-0000-C000-000000000046")]
[InterfaceType(2)]
public interface Outline
{
	void _VtblGap1_5();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(960)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ShowLevels([Optional][In][MarshalAs(UnmanagedType.Struct)] object RowLevels, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ColumnLevels);
}
