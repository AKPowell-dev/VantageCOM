using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C03A0-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface OfficeTheme : _IMsoDispObj
{
	void _VtblGap1_3();

	[DispId(2)]
	ThemeColorScheme ThemeColorScheme
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
