using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("CDB0FF41-E862-47BB-AE77-3FA7B1AE3189")]
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
