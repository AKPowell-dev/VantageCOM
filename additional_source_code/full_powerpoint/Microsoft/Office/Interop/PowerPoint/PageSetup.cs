using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493466-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
[CompilerGenerated]
public interface PageSetup
{
	void _VtblGap1_4();

	[DispId(2004)]
	float SlideHeight
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[param: In]
		set;
	}

	[DispId(2005)]
	float SlideWidth
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		[param: In]
		set;
	}

	[DispId(2006)]
	PpSlideSizeType SlideSize
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[param: In]
		set;
	}

	void _VtblGap2_2();

	[DispId(2008)]
	MsoOrientation SlideOrientation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2008)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2008)]
		[param: In]
		set;
	}
}
