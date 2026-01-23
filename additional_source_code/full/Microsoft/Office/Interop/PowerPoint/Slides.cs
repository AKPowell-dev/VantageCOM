using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493469-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
[TypeIdentifier]
public interface Slides : Collection
{
	void _VtblGap1_2();

	[DispId(11)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		get;
	}

	void _VtblGap2_2();

	[DispId(0)]
	Slide this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
