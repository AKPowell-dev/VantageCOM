using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C170E-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface IMsoTrendline
{
	void _VtblGap1_29();

	[DispId(1610743837)]
	IMsoChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
