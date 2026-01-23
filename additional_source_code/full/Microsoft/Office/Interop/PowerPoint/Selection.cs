using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("91493454-5A91-11CF-8700-00AA0060263B")]
public interface Selection
{
	void _VtblGap1_8();

	[DispId(2009)]
	ShapeRange ShapeRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2009)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
