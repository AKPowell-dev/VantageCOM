using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[DefaultMember("Colors")]
[Guid("000C03A2-0000-0000-C000-000000000046")]
public interface ThemeColorScheme : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ThemeColor Colors([In] MsoThemeColorSchemeIndex Index);
}
