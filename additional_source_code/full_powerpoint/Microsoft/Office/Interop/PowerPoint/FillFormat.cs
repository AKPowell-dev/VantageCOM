using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("9149347E-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface FillFormat
{
	void _VtblGap1_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(11)]
	void OneColorGradient([In] MsoGradientStyle Style, [In] int Variant, [In] float Degree);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(12)]
	void Patterned([In] MsoPatternType Pattern);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(13)]
	void PresetGradient([In] MsoGradientStyle Style, [In] int Variant, [In] MsoPresetGradientType PresetGradientType);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(15)]
	void Solid();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(16)]
	void TwoColorGradient([In] MsoGradientStyle Style, [In] int Variant);

	void _VtblGap3_2();

	[DispId(100)]
	ColorFormat BackColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	[DispId(101)]
	ColorFormat ForeColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	[DispId(102)]
	MsoGradientColorType GradientColorType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(102)]
		get;
	}

	[DispId(103)]
	float GradientDegree
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(103)]
		get;
	}

	[DispId(104)]
	MsoGradientStyle GradientStyle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(104)]
		get;
	}

	[DispId(105)]
	int GradientVariant
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(105)]
		get;
	}

	[DispId(106)]
	MsoPatternType Pattern
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(106)]
		get;
	}

	[DispId(107)]
	MsoPresetGradientType PresetGradientType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(107)]
		get;
	}

	void _VtblGap4_3();

	[DispId(111)]
	float Transparency
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		[param: In]
		set;
	}

	[DispId(112)]
	MsoFillType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(112)]
		get;
	}

	[DispId(113)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		[param: In]
		set;
	}

	[DispId(114)]
	GradientStops GradientStops
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(114)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap5_14();

	[DispId(122)]
	PictureEffects PictureEffects
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(123)]
	float GradientAngle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		[param: In]
		set;
	}
}
