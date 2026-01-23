using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00024427-0000-0000-C000-000000000046")]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
public interface Comment
{
	void _VtblGap1_2();

	[DispId(150)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(150)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_1();

	[DispId(1582)]
	Shape Shape
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1582)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(138)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Text([Optional][In][MarshalAs(UnmanagedType.Struct)] object Text, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Overwrite);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	void Delete();
}
