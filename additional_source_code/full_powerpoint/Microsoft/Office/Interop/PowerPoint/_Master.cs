using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("9149346C-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
[CompilerGenerated]
public interface _Master
{
	void _VtblGap1_2();

	[DispId(2003)]
	Shapes Shapes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2004)]
	HeadersFooters HeadersFooters
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_6();

	[DispId(2009)]
	float Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2009)]
		get;
	}

	[DispId(2010)]
	float Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2010)]
		get;
	}

	void _VtblGap3_3();

	[DispId(2014)]
	Design Design
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_2();

	[DispId(2017)]
	CustomLayouts CustomLayouts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2017)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2018)]
	OfficeTheme Theme
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2018)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
