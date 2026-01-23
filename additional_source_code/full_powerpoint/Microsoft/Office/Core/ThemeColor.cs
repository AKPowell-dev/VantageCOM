using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[DefaultMember("RGB")]
[CompilerGenerated]
[Guid("000C03A1-0000-0000-C000-000000000046")]
public interface ThemeColor : _IMsoDispObj
{
	void _VtblGap1_2();

	[DispId(0)]
	int RGB
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		set;
	}
}
