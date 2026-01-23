using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020996-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface KeyBindings : IEnumerable
{
	void _VtblGap1_6();

	[DispId(0)]
	KeyBinding this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(101)]
	[return: MarshalAs(UnmanagedType.Interface)]
	KeyBinding Add([In] WdKeyCategory KeyCategory, [In][MarshalAs(UnmanagedType.BStr)] string Command, [In] int KeyCode, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object KeyCode2, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object CommandParameter);

	void _VtblGap2_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(110)]
	[return: MarshalAs(UnmanagedType.Interface)]
	KeyBinding Key([In] int KeyCode, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object KeyCode2);
}
