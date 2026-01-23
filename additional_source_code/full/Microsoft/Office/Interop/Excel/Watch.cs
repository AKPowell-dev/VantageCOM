using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("00024457-0000-0000-C000-000000000046")]
[InterfaceType(2)]
public interface Watch
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(117)]
	void Delete();

	[DispId(222)]
	object Source
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(222)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}
}
