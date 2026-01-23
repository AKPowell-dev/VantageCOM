using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[DefaultMember("_Default")]
[InterfaceType(2)]
[CompilerGenerated]
[Guid("00024471-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface ListObject
{
	void _VtblGap1_7();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2309)]
	void Unlist();

	void _VtblGap2_2();

	[DispId(0)]
	string _Default
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap3_9();

	[DispId(1386)]
	QueryTable QueryTable
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1386)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(197)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(197)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
