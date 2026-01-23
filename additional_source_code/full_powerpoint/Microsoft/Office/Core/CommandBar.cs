using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C0304-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface CommandBar : _IMsoOleAccDispObj
{
	void _VtblGap1_30();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610874887)]
	[return: MarshalAs(UnmanagedType.Interface)]
	CommandBarControl FindControl([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Id, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Tag, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Visible, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Recursive);
}
