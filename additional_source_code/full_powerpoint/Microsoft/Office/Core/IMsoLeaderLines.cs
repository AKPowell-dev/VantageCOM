using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("000C1723-0000-0000-C000-000000000046")]
public interface IMsoLeaderLines
{
	void _VtblGap1_3();

	[DispId(1610743811)]
	IMsoChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743811)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
