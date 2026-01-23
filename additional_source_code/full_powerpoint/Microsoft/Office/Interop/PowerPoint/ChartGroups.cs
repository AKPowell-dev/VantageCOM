using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("92D41A5E-F07E-4CA4-AF6F-BEF486AA4E6F")]
[TypeIdentifier]
public interface ChartGroups : IEnumerable
{
	void _VtblGap1_1();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}
}
