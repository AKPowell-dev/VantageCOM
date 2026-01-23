using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("000C03C8-0000-0000-C000-000000000046")]
public interface SmartArtNode : _IMsoDispObj
{
	void _VtblGap1_9();

	[DispId(7)]
	ShapeRange Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(7)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(8)]
	TextFrame2 TextFrame2
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(8)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
