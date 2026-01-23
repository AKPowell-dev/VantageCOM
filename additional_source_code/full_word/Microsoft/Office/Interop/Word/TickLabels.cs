using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("935D59F5-E365-4F92-B7F5-1C499A63ECA8")]
[TypeIdentifier]
[CompilerGenerated]
public interface TickLabels
{
	void _VtblGap1_2();

	[DispId(1610743810)]
	ChartFont Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743810)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
