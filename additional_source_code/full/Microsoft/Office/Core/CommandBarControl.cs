using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("000C0308-0000-0000-C000-000000000046")]
public interface CommandBarControl : _IMsoOleAccDispObj
{
	void _VtblGap1_28();

	[DispId(1610874885)]
	object Control
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610874885)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}
}
