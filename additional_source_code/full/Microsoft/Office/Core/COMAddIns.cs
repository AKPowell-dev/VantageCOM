using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C0339-0000-0000-C000-000000000046")]
[CompilerGenerated]
[DefaultMember("Item")]
public interface COMAddIns : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	COMAddIn Item([In][MarshalAs(UnmanagedType.Struct)] ref object Index);
}
