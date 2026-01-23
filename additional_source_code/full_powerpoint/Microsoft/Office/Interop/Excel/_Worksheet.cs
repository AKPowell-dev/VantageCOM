using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("000208D8-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _Worksheet
{
	void _VtblGap1_2();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(304)]
	[LCIDConversion(0)]
	void Activate();

	void _VtblGap2_21();

	[DispId(292)]
	bool ProtectContents
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(292)]
		[LCIDConversion(0)]
		get;
	}

	void _VtblGap3_6();

	[DispId(558)]
	XlSheetVisibility Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(558)]
		[param: In]
		set;
	}

	void _VtblGap4_11();

	[DispId(238)]
	Range Cells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(238)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(1060)]
	[return: MarshalAs(UnmanagedType.IDispatch)]
	object ChartObjects([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);
}
