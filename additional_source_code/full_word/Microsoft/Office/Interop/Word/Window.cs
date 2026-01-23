using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("00020962-0000-0000-C000-000000000046")]
[TypeIdentifier]
[DefaultMember("Caption")]
public interface Window
{
	void _VtblGap1_6();

	[DispId(4)]
	Selection Selection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(4)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_12();

	[DispId(0)]
	string Caption
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap3_6();

	[DispId(14)]
	View View
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(14)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_24();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(100)]
	void Activate();

	void _VtblGap5_8();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(111)]
	void ScrollIntoView([In][MarshalAs(UnmanagedType.IDispatch)] object obj, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Start);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(112)]
	void GetPoint(out int ScreenPixelsLeft, out int ScreenPixelsTop, out int ScreenPixelsWidth, out int ScreenPixelsHeight, [In][MarshalAs(UnmanagedType.IDispatch)] object obj);

	void _VtblGap6_9();

	[DispId(36)]
	bool Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(36)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(36)]
		[param: In]
		set;
	}

	void _VtblGap7_7();

	[DispId(39)]
	int Hwnd
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(39)]
		get;
	}
}
