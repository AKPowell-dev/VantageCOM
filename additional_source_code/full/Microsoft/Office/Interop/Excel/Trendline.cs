using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208BE-0000-0000-C000-000000000046")]
[CompilerGenerated]
[InterfaceType(2)]
[TypeIdentifier]
public interface Trendline
{
	void _VtblGap1_7();

	[DispId(158)]
	DataLabel DataLabel
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(158)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_3();

	[DispId(189)]
	bool DisplayRSquared
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(189)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(189)]
		set;
	}

	void _VtblGap3_22();

	[DispId(116)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(116)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
