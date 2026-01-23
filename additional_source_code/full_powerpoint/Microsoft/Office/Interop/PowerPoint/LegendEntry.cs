using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("92D41A70-F07E-4CA4-AF6F-BEF486AA4E6F")]
[TypeIdentifier]
[CompilerGenerated]
public interface LegendEntry
{
	void _VtblGap1_2();

	[DispId(146)]
	ChartFont Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(146)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_7();

	[DispId(122)]
	double Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		get;
	}

	[DispId(123)]
	double Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		get;
	}
}
