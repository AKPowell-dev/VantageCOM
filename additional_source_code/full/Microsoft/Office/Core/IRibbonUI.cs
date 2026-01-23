using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C03A7-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface IRibbonUI
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1)]
	void Invalidate();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2)]
	void InvalidateControl([In][MarshalAs(UnmanagedType.BStr)] string ControlID);
}
