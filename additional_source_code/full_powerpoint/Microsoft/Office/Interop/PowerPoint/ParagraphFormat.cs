using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493496-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
[TypeIdentifier]
public interface ParagraphFormat
{
	void _VtblGap1_1();

	[DispId(2002)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2002)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[DispId(2003)]
	PpParagraphAlignment Alignment
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[param: In]
		set;
	}

	[DispId(2004)]
	BulletFormat Bullet
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
