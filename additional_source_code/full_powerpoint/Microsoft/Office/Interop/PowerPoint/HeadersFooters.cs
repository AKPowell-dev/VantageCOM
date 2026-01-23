using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("91493474-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface HeadersFooters
{
	void _VtblGap1_2();

	[DispId(2003)]
	HeaderFooter DateAndTime
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2004)]
	HeaderFooter SlideNumber
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_1();

	[DispId(2006)]
	HeaderFooter Footer
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
