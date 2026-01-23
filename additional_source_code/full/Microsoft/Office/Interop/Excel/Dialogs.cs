using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00020879-0000-0000-C000-000000000046")]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
public interface Dialogs : IEnumerable
{
	void _VtblGap1_5();

	[IndexerName("_Default")]
	[DispId(0)]
	Dialog this[[In] XlBuiltInDialog Index]
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
