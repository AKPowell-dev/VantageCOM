using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493471-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
[CompilerGenerated]
public interface SlideShowTransition
{
	void _VtblGap1_8();

	[DispId(2006)]
	PpEntryEffect EntryEffect
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[param: In]
		set;
	}

	[DispId(2007)]
	MsoTriState Hidden
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2007)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2007)]
		[param: In]
		set;
	}
}
