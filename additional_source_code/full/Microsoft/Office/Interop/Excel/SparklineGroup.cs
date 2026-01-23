using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
[Guid("000244B7-0000-0000-C000-000000000046")]
public interface SparklineGroup : IEnumerable
{
	void _VtblGap1_6();

	[DispId(1397)]
	Range Location
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1397)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1397)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Interface)]
		set;
	}

	void _VtblGap2_8();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	void Delete();
}
