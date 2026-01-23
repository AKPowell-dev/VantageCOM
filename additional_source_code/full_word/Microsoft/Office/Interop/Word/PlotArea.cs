using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("56AFD330-440C-4F4C-A39C-ED306D084D5F")]
[TypeIdentifier]
public interface PlotArea
{
	void _VtblGap1_15();

	[DispId(1667)]
	double InsideLeft
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1667)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1667)]
		[param: In]
		set;
	}

	[DispId(1668)]
	double InsideTop
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1668)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1668)]
		[param: In]
		set;
	}

	[DispId(1669)]
	double InsideWidth
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1669)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1669)]
		[param: In]
		set;
	}

	[DispId(1670)]
	double InsideHeight
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1670)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1670)]
		[param: In]
		set;
	}

	void _VtblGap2_2();

	[DispId(1610743833)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743833)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
