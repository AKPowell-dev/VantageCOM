using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C0365-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[DefaultMember("Item")]
public interface FileDialogFilters : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.Interface)]
	FileDialogFilter Item([In] int Index);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809349)]
	void Clear();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809350)]
	[return: MarshalAs(UnmanagedType.Interface)]
	FileDialogFilter Add([In][MarshalAs(UnmanagedType.BStr)] string Description, [In][MarshalAs(UnmanagedType.BStr)] string Extensions, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Position);
}
