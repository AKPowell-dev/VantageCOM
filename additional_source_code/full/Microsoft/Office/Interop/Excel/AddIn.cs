using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[Guid("00020857-0000-0000-C000-000000000046")]
[TypeIdentifier]
[InterfaceType(2)]
public interface AddIn
{
	void _VtblGap1_5();

	[DispId(289)]
	string FullName
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(289)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(550)]
	bool Installed
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(550)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(550)]
		set;
	}

	void _VtblGap2_1();

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
