using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("92D41A6A-F07E-4CA4-AF6F-BEF486AA4E6F")]
[CompilerGenerated]
[TypeIdentifier]
public interface Gridlines
{
	void _VtblGap1_5();

	[DispId(2001)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2001)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
