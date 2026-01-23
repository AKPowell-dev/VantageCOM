using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C1721-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface IMsoErrorBars
{
	void _VtblGap1_8();

	[DispId(1610743816)]
	IMsoChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
