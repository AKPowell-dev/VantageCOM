using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[InterfaceType(2)]
[Guid("00020A01-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface ApplicationEvents4
{
	void _VtblGap1_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(4)]
	void DocumentOpen([In][MarshalAs(UnmanagedType.Interface)] Document Doc);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(6)]
	void DocumentBeforeClose([In][MarshalAs(UnmanagedType.Interface)] Document Doc, [In][Out] ref bool Cancel);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(8)]
	void DocumentBeforeSave([In][MarshalAs(UnmanagedType.Interface)] Document Doc, [In] ref bool SaveAsUI, [In][Out] ref bool Cancel);

	void _VtblGap3_3();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(12)]
	void WindowSelectionChange([In][MarshalAs(UnmanagedType.Interface)] Selection Sel);
}
