using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00020895-0000-0000-C000-000000000046")]
[InterfaceType(2)]
[TypeIdentifier]
[CompilerGenerated]
public interface Pane
{
	void _VtblGap1_6();

	[DispId(654)]
	int ScrollColumn
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(654)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(654)]
		set;
	}

	[DispId(655)]
	int ScrollRow
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(655)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(655)]
		set;
	}

	void _VtblGap2_1();

	[DispId(1118)]
	Range VisibleRange
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1118)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
