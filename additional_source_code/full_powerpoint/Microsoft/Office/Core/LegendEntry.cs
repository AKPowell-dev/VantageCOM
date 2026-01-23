using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C171A-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface LegendEntry
{
	void _VtblGap1_2();

	[DispId(146)]
	ChartFont Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_5();

	[DispId(127)]
	double Left
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(126)]
	double Top
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(122)]
	double Width
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}

	[DispId(123)]
	double Height
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		get;
	}
}
