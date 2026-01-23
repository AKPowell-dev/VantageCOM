using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("92D41A6F-F07E-4CA4-AF6F-BEF486AA4E6F")]
[DefaultMember("_Default")]
[TypeIdentifier]
public interface LegendEntries : IEnumerable
{
	void _VtblGap1_1();

	[DispId(118)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		get;
	}

	void _VtblGap2_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	LegendEntry _Default([In][MarshalAs(UnmanagedType.Struct)] object Index);
}
