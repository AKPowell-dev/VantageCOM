using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("914934DF-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface Effect
{
	void _VtblGap1_2();

	[DispId(2003)]
	Shape Shape
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2004)]
	void MoveTo([In] int toPos);

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	void Delete();

	void _VtblGap3_1();

	[DispId(2009)]
	Timing Timing
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2009)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2010)]
	MsoAnimEffect EffectType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2010)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2010)]
		[param: In]
		set;
	}

	void _VtblGap4_3();

	[DispId(2014)]
	int Paragraph
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		[param: In]
		set;
	}

	void _VtblGap5_1();

	[DispId(2016)]
	MsoTriState Exit
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2016)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2016)]
		[param: In]
		set;
	}
}
