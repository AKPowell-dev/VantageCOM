using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[TypeIdentifier]
[Guid("000208BD-0000-0000-C000-000000000046")]
[DefaultMember("_Default")]
[CompilerGenerated]
public interface Trendlines : IEnumerable
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(181)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Trendline Add([In] XlTrendlineType Type = XlTrendlineType.xlLinear, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Order, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Period, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Forward, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Backward, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Intercept, [Optional][In][MarshalAs(UnmanagedType.Struct)] object DisplayEquation, [Optional][In][MarshalAs(UnmanagedType.Struct)] object DisplayRSquared, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Name);

	void _VtblGap2_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Trendline _Default([Optional][In][MarshalAs(UnmanagedType.Struct)] object Index);
}
