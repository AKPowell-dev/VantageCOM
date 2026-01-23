using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("9149346A-5A91-11CF-8700-00AA0060263B")]
public interface _Slide
{
	void _VtblGap1_2();

	[DispId(2003)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
