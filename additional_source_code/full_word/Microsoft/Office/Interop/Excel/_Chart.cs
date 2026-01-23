using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[Guid("000208D6-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _Chart
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

	void _VtblGap2_44();

	[DispId(80)]
	ChartArea ChartArea
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(80)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_121();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2508)]
	void SaveChartTemplate([In][MarshalAs(UnmanagedType.BStr)] string Filename);
}
