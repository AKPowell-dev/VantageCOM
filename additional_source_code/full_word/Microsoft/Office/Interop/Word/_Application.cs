using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020970-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
[DefaultMember("Name")]
public interface _Application
{
	void _VtblGap1_3();

	[DispId(0)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(6)]
	Documents Documents
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2)]
	Windows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(3)]
	Document ActiveDocument
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(4)]
	Window ActiveWindow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(4)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(5)]
	Selection Selection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(5)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_2();

	[DispId(8)]
	Template NormalTemplate
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(8)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_13();

	[DispId(22)]
	AddIns AddIns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(22)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_2();

	[DispId(24)]
	string Version
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(24)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(26)]
	bool ScreenUpdating
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(26)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(26)]
		[param: In]
		set;
	}

	void _VtblGap5_11();

	[DispId(47)]
	string Build
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(47)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap6_2();

	[DispId(52)]
	string UserName
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(52)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(52)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap7_7();

	[DispId(57)]
	CommandBars CommandBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(57)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_7();

	[DispId(67)]
	Templates Templates
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(67)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(68)]
	object CustomizationContext
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		[param: In]
		[param: MarshalAs(UnmanagedType.IDispatch)]
		set;
	}

	[DispId(69)]
	KeyBindings KeyBindings
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(69)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap9_1();

	[DispId(71)]
	KeyBinding FindKey
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(71)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(80)]
	string Caption
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(80)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(80)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap10_20();

	[DispId(94)]
	WdAlertLevel DisplayAlerts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(94)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(94)]
		[param: In]
		set;
	}

	void _VtblGap11_6();

	[DispId(100)]
	WdEnableCancelKey EnableCancelKey
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		[param: In]
		set;
	}

	[DispId(101)]
	bool UserControl
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		get;
	}

	void _VtblGap12_13();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(301)]
	void ScreenRefresh();

	void _VtblGap13_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(316)]
	int BuildKeyCode([In] WdKey Arg1, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Arg2, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Arg3, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Arg4);

	void _VtblGap14_28();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(370)]
	float InchesToPoints([In] float Inches);

	void _VtblGap15_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(387)]
	float PointsToPixels([In] float Points, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object fVertical);

	void _VtblGap16_15();

	[DispId(403)]
	LanguageSettings LanguageSettings
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(403)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap17_9();

	[DispId(450)]
	FileDialog FileDialog
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(450)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap18_41();

	[DispId(486)]
	UndoRecord UndoRecord
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(486)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
