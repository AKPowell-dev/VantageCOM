using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("91493479-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface Shape
{
	[DispId(2001)]
	object Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2001)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap1_1();

	[DispId(1)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(10)]
	void Apply();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(11)]
	void Delete();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(13)]
	void Flip([In] MsoFlipCmd FlipCmd);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(15)]
	void IncrementRotation([In] float Increment);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(17)]
	void PickUp();

	void _VtblGap4_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(19)]
	void ScaleHeight([In] float Factor, [In] MsoTriState RelativeToOriginalSize, [In] MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(20)]
	void ScaleWidth([In] float Factor, [In] MsoTriState RelativeToOriginalSize, [In] MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft);

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(23)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ShapeRange Ungroup();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(24)]
	void ZOrder([In] MsoZOrderCmd ZOrderCmd);

	[DispId(100)]
	Adjustments Adjustments
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(101)]
	MsoAutoShapeType AutoShapeType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[param: In]
		set;
	}

	[DispId(102)]
	MsoBlackWhiteMode BlackWhiteMode
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(102)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(102)]
		[param: In]
		set;
	}

	void _VtblGap6_4();

	[DispId(107)]
	FillFormat Fill
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(107)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(108)]
	GroupShapes GroupItems
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(108)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(109)]
	float Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(109)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(109)]
		[param: In]
		set;
	}

	[DispId(110)]
	MsoTriState HorizontalFlip
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		get;
	}

	[DispId(111)]
	float Left
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
	LineFormat Line
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(112)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(113)]
	MsoTriState LockAspectRatio
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		[param: In]
		set;
	}

	[DispId(115)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(115)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(115)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	[DispId(116)]
	ShapeNodes Nodes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(116)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(117)]
	float Rotation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		[param: In]
		set;
	}

	[DispId(118)]
	PictureFormat PictureFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(118)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(119)]
	ShadowFormat Shadow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(119)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(120)]
	TextEffectFormat TextEffect
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(120)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(121)]
	TextFrame TextFrame
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(121)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(122)]
	ThreeDFormat ThreeD
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(122)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(123)]
	float Top
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(123)]
		[param: In]
		set;
	}

	[DispId(124)]
	MsoShapeType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(124)]
		get;
	}

	[DispId(125)]
	MsoTriState VerticalFlip
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(125)]
		get;
	}

	void _VtblGap7_1();

	[DispId(127)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		[param: In]
		set;
	}

	[DispId(128)]
	float Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(128)]
		[param: In]
		set;
	}

	[DispId(129)]
	int ZOrderPosition
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(129)]
		get;
	}

	[DispId(2003)]
	OLEFormat OLEFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2004)]
	LinkFormat LinkFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2005)]
	PlaceholderFormat PlaceholderFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2006)]
	AnimationSettings AnimationSettings
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2006)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_1();

	[DispId(2008)]
	Tags Tags
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2008)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2009)]
	void Cut();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2010)]
	void Copy();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2011)]
	void Select([In] MsoTriState Replace = MsoTriState.msoTrue);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2012)]
	[return: MarshalAs(UnmanagedType.Interface)]
	ShapeRange Duplicate();

	[DispId(2013)]
	PpMediaType MediaType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2013)]
		get;
	}

	[DispId(2014)]
	MsoTriState HasTextFrame
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		get;
	}

	void _VtblGap9_2();

	[DispId(131)]
	string AlternativeText
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(131)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(131)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	[DispId(2016)]
	MsoTriState HasTable
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2016)]
		get;
	}

	[DispId(2017)]
	Table Table
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2017)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2018)]
	void Export([In][MarshalAs(UnmanagedType.BStr)] string PathName, [In] PpShapeFormat Filter, [In] int ScaleWidth = 0, [In] int ScaleHeight = 0, [In] PpExportMode ExportMode = PpExportMode.ppRelativeToSlide);

	void _VtblGap10_5();

	[DispId(137)]
	Shape ParentGroup
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(137)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_1();

	[DispId(139)]
	int Id
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(139)]
		get;
	}

	void _VtblGap12_6();

	[DispId(145)]
	TextFrame2 TextFrame2
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(145)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(148)]
	MsoTriState HasChart
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		get;
	}

	void _VtblGap13_4();

	[DispId(152)]
	SoftEdgeFormat SoftEdge
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(152)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(153)]
	GlowFormat Glow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(153)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(154)]
	ReflectionFormat Reflection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(154)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2021)]
	Chart Chart
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2021)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(155)]
	MsoTriState HasSmartArt
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(155)]
		get;
	}

	[DispId(156)]
	SmartArt SmartArt
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(156)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap14_1();

	[DispId(158)]
	string Title
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(158)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(158)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap15_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2023)]
	void PickupAnimation();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2024)]
	void ApplyAnimation();
}
