using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
[Guid("00024413-0000-0000-C000-000000000046")]
public interface AppEvents
{
	void _VtblGap1_8();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1567)]
	void WorkbookOpen([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb);
}
