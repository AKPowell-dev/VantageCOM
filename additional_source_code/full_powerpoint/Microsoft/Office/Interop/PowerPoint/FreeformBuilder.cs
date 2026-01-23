using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("91493478-5A91-11CF-8700-00AA0060263B")]
public interface FreeformBuilder
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(10)]
	void AddNodes([In] MsoSegmentType SegmentType, [In] MsoEditingType EditingType, [In] float X1, [In] float Y1, [In] float X2 = 0f, [In] float Y2 = 0f, [In] float X3 = 0f, [In] float Y3 = 0f);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(11)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape ConvertToShape();
}
