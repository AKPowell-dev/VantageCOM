using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("91493486-5A91-11CF-8700-00AA0060263B")]
public interface ShapeNodes : IEnumerable
{
	void _VtblGap1_3();

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	[DispId(0)]
	ShapeNode this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(14)]
	void SetPosition([In] int Index, [In] float X1, [In] float Y1);
}
