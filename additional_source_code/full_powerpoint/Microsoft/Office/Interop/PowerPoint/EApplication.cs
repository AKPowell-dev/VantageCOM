using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("914934C2-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface EApplication
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2001)]
	void WindowSelectionChange([In][MarshalAs(UnmanagedType.Interface)] Selection Sel);

	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2005)]
	void PresentationSave([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2006)]
	void PresentationOpen([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	void NewPresentation([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2008)]
	void PresentationNewSlide([In][MarshalAs(UnmanagedType.Interface)] Slide Sld);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2009)]
	void WindowActivate([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres, [In][MarshalAs(UnmanagedType.Interface)] DocumentWindow Wn);

	void _VtblGap2_6();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2016)]
	void SlideSelectionChanged([In][MarshalAs(UnmanagedType.Interface)] SlideRange SldRange);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2018)]
	void PresentationBeforeSave([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres, [In][Out] ref bool Cancel);

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2020)]
	void AfterNewPresentation([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2021)]
	void AfterPresentationOpen([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	void _VtblGap5_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2025)]
	void PresentationBeforeClose([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres, [In][Out] ref bool Cancel);

	void _VtblGap6_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2031)]
	void PresentationCloseFinal([In][MarshalAs(UnmanagedType.Interface)] Presentation Pres);

	void _VtblGap7_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2033)]
	void AfterShapeSizeChange([In][MarshalAs(UnmanagedType.Interface)] Shape shp);
}
