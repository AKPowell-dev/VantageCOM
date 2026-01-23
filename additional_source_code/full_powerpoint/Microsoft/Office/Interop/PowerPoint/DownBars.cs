using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("92D41A65-F07E-4CA4-AF6F-BEF486AA4E6F")]
public interface DownBars
{
	void _VtblGap1_7();

	[DispId(2001)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2001)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
