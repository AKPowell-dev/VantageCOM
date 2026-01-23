using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C1730-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface IMsoChartFormat
{
	[DispId(1610743808)]
	FillFormat Fill
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743808)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_1();

	[DispId(1610743810)]
	LineFormat Line
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743810)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
