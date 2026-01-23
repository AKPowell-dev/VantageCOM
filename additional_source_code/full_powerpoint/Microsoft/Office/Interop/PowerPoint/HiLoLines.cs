using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("92D41A6B-F07E-4CA4-AF6F-BEF486AA4E6F")]
[CompilerGenerated]
public interface HiLoLines
{
	void _VtblGap1_3();

	[DispId(2004)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
