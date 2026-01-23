using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("91493475-5A91-11CF-8700-00AA0060263B")]
public interface Shapes : IEnumerable
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
	Shape this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap2_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(15)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddPicture([In][MarshalAs(UnmanagedType.BStr)] string FileName, [In] MsoTriState LinkToFile, [In] MsoTriState SaveWithDocument, [In] float Left, [In] float Top, [In] float Width = -1f, [In] float Height = -1f);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(17)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddShape([In] MsoAutoShapeType Type, [In] float Left, [In] float Top, [In] float Width, [In] float Height);

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(19)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddTextbox([In] MsoTextOrientation Orientation, [In] float Left, [In] float Top, [In] float Width, [In] float Height);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(20)]
	[return: MarshalAs(UnmanagedType.Interface)]
	FreeformBuilder BuildFreeform([In] MsoEditingType EditingType, [In] float X1, [In] float Y1);

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2003)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ShapeRange Range([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);

	[DispId(2004)]
	MsoTriState HasTitle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
	}

	void _VtblGap6_1();

	[DispId(2006)]
	Shape Title
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2007)]
	Placeholders Placeholders
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2007)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2012)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ShapeRange Paste();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2013)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddTable([In] int NumRows, [In] int NumColumns, [In] float Left = -1f, [In] float Top = -1f, [In] float Width = -1f, [In] float Height = -1f);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2014)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ShapeRange PasteSpecial([In] PpPasteDataType DataType = PpPasteDataType.ppPasteDefault, [In] MsoTriState DisplayAsIcon = MsoTriState.msoFalse, [In][MarshalAs(UnmanagedType.BStr)] string IconFileName = "", [In] int IconIndex = 0, [In][MarshalAs(UnmanagedType.BStr)] string IconLabel = "", [In] MsoTriState Link = MsoTriState.msoFalse);

	void _VtblGap8_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(30)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Shape AddPicture2([In][MarshalAs(UnmanagedType.BStr)] string FileName, [In] MsoTriState LinkToFile, [In] MsoTriState SaveWithDocument, [In] float Left, [In] float Top, [In] float Width = -1f, [In] float Height = -1f, [In] MsoPictureCompress compress = MsoPictureCompress.msoPictureCompressDocDefault);
}
