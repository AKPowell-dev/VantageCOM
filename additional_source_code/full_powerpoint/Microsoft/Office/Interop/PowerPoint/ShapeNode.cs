using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("91493487-5A91-11CF-8700-00AA0060263B")]
public interface ShapeNode
{
	void _VtblGap1_4();

	[DispId(101)]
	object Points
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}
}
