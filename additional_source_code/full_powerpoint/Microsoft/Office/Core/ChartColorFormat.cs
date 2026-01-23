using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("000C171D-0000-0000-C000-000000000046")]
[DefaultMember("_Default")]
public interface ChartColorFormat
{
	void _VtblGap1_5();

	[DispId(0)]
	int _Default
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}
}
