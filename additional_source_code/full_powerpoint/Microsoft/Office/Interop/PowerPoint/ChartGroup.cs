using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("92D41A5D-F07E-4CA4-AF6F-BEF486AA4E6F")]
public interface ChartGroup
{
	[DispId(1610743808)]
	DownBars DownBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743808)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1610743809)]
	DropLines DropLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743809)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1610743810)]
	bool HasDropLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743810)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743810)]
		[param: In]
		set;
	}

	[DispId(1610743812)]
	bool HasHiLoLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743812)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743812)]
		[param: In]
		set;
	}

	[DispId(1610743814)]
	bool HasRadarAxisLabels
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743814)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743814)]
		[param: In]
		set;
	}

	void _VtblGap1_2();

	[DispId(1610743818)]
	bool HasUpDownBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743818)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743818)]
		[param: In]
		set;
	}

	[DispId(1610743820)]
	HiLoLines HiLoLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743820)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_2();

	[DispId(1610743823)]
	UpBars UpBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743823)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_28();

	[DispId(2014)]
	TickLabels RadarAxisLabels
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2014)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
