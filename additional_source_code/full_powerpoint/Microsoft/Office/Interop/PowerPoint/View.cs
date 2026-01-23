using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493458-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
[TypeIdentifier]
public interface View
{
	void _VtblGap1_2();

	[DispId(2003)]
	PpViewType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		get;
	}

	[DispId(2004)]
	int Zoom
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[param: In]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2005)]
	void Paste();

	[DispId(2006)]
	object Slide
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[param: In]
		[param: MarshalAs(UnmanagedType.IDispatch)]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	void GotoSlide([In] int Index);

	void _VtblGap2_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2010)]
	void PasteSpecial([In] PpPasteDataType DataType = PpPasteDataType.ppPasteDefault, [In] MsoTriState DisplayAsIcon = MsoTriState.msoFalse, [In][MarshalAs(UnmanagedType.BStr)] string IconFileName = "", [In] int IconIndex = 0, [In][MarshalAs(UnmanagedType.BStr)] string IconLabel = "", [In] MsoTriState Link = MsoTriState.msoFalse);
}
