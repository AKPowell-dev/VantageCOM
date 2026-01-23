using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C1718-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface ChartFont
{
	void _VtblGap1_4();

	[DispId(1610743812)]
	object Color
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743812)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743812)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}
}
