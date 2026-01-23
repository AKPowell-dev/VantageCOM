using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("BA72E551-4FF5-48F4-8215-5505F990966F")]
[TypeIdentifier]
public interface SectionProperties
{
	void _VtblGap1_2();

	[DispId(2003)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2004)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Name([In] int sectionIndex);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2005)]
	void Rename([In] int sectionIndex, [In][MarshalAs(UnmanagedType.BStr)] string sectionName);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2006)]
	int SlidesCount([In] int sectionIndex);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	int FirstSlide([In] int sectionIndex);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2008)]
	int AddBeforeSlide([In] int SlideIndex, [In][MarshalAs(UnmanagedType.BStr)] string sectionName);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2010)]
	void Move([In] int sectionIndex, [In] int toPos);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2011)]
	void Delete([In] int sectionIndex, [In] bool deleteSlides);
}
