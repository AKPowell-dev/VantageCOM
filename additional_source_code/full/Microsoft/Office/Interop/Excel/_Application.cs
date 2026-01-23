using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208D5-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
[DefaultMember("_Default")]
public interface _Application
{
	[DispId(148)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_2();

	[DispId(305)]
	Range ActiveCell
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(305)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(183)]
	Chart ActiveChart
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(183)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_4();

	[DispId(307)]
	object ActiveSheet
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(307)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[DispId(759)]
	Window ActiveWindow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(759)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(308)]
	Workbook ActiveWorkbook
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(308)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(549)]
	AddIns AddIns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(549)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_4();

	[DispId(241)]
	Range Columns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(241)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1439)]
	CommandBars CommandBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1439)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_7();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object Evaluate([In][MarshalAs(UnmanagedType.Struct)] object Name);

	void _VtblGap5_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(350)]
	[LCIDConversion(1)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ExecuteExcel4Macro([In][MarshalAs(UnmanagedType.BStr)] string String);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(766)]
	[LCIDConversion(30)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Intersect([In][MarshalAs(UnmanagedType.Interface)] Range Arg1, [In][MarshalAs(UnmanagedType.Interface)] Range Arg2, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg3, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg4, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg5, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg6, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg7, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg8, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg9, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg10, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg11, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg12, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg13, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg14, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg15, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg16, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg17, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg18, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg19, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg20, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg21, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg22, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg23, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg24, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg25, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg26, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg27, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg28, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg29, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg30);

	void _VtblGap6_3();

	[DispId(197)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(197)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(258)]
	Range Rows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(258)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap7_2();

	[DispId(147)]
	object Selection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(147)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(383)]
	[LCIDConversion(2)]
	void SendKeys([In][MarshalAs(UnmanagedType.Struct)] object Keys, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Wait);

	[DispId(485)]
	Sheets Sheets
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(485)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap8_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(779)]
	[LCIDConversion(30)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Range Union([In][MarshalAs(UnmanagedType.Interface)] Range Arg1, [In][MarshalAs(UnmanagedType.Interface)] Range Arg2, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg3, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg4, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg5, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg6, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg7, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg8, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg9, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg10, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg11, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg12, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg13, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg14, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg15, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg16, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg17, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg18, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg19, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg20, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg21, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg22, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg23, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg24, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg25, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg26, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg27, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg28, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg29, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Arg30);

	[DispId(430)]
	Windows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(430)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(572)]
	Workbooks Workbooks
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(572)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1440)]
	WorksheetFunction WorksheetFunction
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1440)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(494)]
	Sheets Worksheets
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(494)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap9_13();

	[DispId(1145)]
	AutoCorrect AutoCorrect
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1145)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(314)]
	int Build
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(314)]
		get;
	}

	[DispId(315)]
	bool CalculateBeforeSave
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(315)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(315)]
		[param: In]
		set;
	}

	[DispId(316)]
	XlCalculation Calculation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(316)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(316)]
		[param: In]
		set;
	}

	void _VtblGap10_3();

	[DispId(139)]
	string Caption
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(139)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(139)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap11_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(1086)]
	double CentimetersToPoints([In] double Centimeters);

	void _VtblGap12_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(5)]
	[DispId(325)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object ConvertFormula([In][MarshalAs(UnmanagedType.Struct)] object Formula, [In] XlReferenceStyle FromReferenceStyle, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ToReferenceStyle, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ToAbsolute, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RelativeTo);

	[DispId(991)]
	bool CopyObjectsWithCells
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(991)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(991)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[DispId(1161)]
	XlMousePointer Cursor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1161)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1161)]
		[param: In]
		set;
	}

	void _VtblGap13_1();

	[DispId(330)]
	XlCutCopyMode CutCopyMode
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(330)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(330)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap14_13();

	[DispId(0)]
	string _Default
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap15_4();

	[DispId(761)]
	Dialogs Dialogs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(761)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(343)]
	bool DisplayAlerts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(343)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(343)]
		[param: In]
		set;
	}

	[DispId(344)]
	bool DisplayFormulaBar
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(344)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(344)]
		[param: In]
		set;
	}

	void _VtblGap16_12();

	[DispId(347)]
	bool DisplayStatusBar
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(347)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(347)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap17_1();

	[DispId(929)]
	bool EditDirectlyInCell
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(929)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(929)]
		[param: In]
		set;
	}

	void _VtblGap18_2();

	[DispId(1096)]
	XlEnableCancelKey EnableCancelKey
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1096)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1096)]
		[param: In]
		set;
	}

	void _VtblGap19_15();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1076)]
	[LCIDConversion(5)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object GetSaveAsFilename([Optional][In][MarshalAs(UnmanagedType.Struct)] object InitialFilename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object FileFilter, [Optional][In][MarshalAs(UnmanagedType.Struct)] object FilterIndex, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Title, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ButtonText);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(475)]
	[LCIDConversion(2)]
	void Goto([Optional][In][MarshalAs(UnmanagedType.Struct)] object Reference, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Scroll);

	void _VtblGap20_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1087)]
	[LCIDConversion(1)]
	double InchesToPoints([In] double Inches);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(357)]
	[LCIDConversion(8)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object InputBox([In][MarshalAs(UnmanagedType.BStr)] string Prompt, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Title, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Default, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Left, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Top, [Optional][In][MarshalAs(UnmanagedType.Struct)] object HelpFile, [Optional][In][MarshalAs(UnmanagedType.Struct)] object HelpContextID, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Type);

	void _VtblGap21_2();

	[DispId(362)]
	object International
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(362)]
		[LCIDConversion(1)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
	}

	[DispId(363)]
	bool Iteration
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(363)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(363)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap22_19();

	[DispId(374)]
	bool MoveAfterReturn
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(374)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(374)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap23_17();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(626)]
	[LCIDConversion(2)]
	void OnKey([In][MarshalAs(UnmanagedType.BStr)] string Key, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Procedure);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(2)]
	[DispId(769)]
	void OnRepeat([In][MarshalAs(UnmanagedType.BStr)] string Text, [In][MarshalAs(UnmanagedType.BStr)] string Procedure);

	void _VtblGap24_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(770)]
	[LCIDConversion(2)]
	void OnUndo([In][MarshalAs(UnmanagedType.BStr)] string Text, [In][MarshalAs(UnmanagedType.BStr)] string Procedure);

	void _VtblGap25_14();

	[DispId(380)]
	XlReferenceStyle ReferenceStyle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(380)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(380)]
		[param: In]
		set;
	}

	void _VtblGap26_8();

	[DispId(382)]
	bool ScreenUpdating
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(382)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(382)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap27_12();

	[DispId(386)]
	object StatusBar
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(386)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(386)]
		[LCIDConversion(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}

	void _VtblGap28_13();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(303)]
	void Undo();

	void _VtblGap29_2();

	[DispId(1210)]
	bool UserControl
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1210)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1210)]
		[param: In]
		set;
	}

	[DispId(391)]
	string UserName
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(391)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(391)]
		[LCIDConversion(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap30_2();

	[DispId(392)]
	string Version
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(392)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(558)]
	bool Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(558)]
		[param: In]
		set;
	}

	void _VtblGap31_5();

	[DispId(396)]
	XlWindowState WindowState
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(396)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(396)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap32_9();

	[DispId(1212)]
	bool EnableEvents
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1212)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1212)]
		[param: In]
		set;
	}

	void _VtblGap33_7();

	[DispId(1796)]
	COMAddIns COMAddIns
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1796)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap34_5();

	[DispId(1801)]
	LanguageSettings LanguageSettings
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1801)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap35_20();

	[DispId(1939)]
	Watches Watches
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1939)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap36_4();

	[DispId(1942)]
	FileDialog FileDialog
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1942)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap37_23();

	[DispId(1809)]
	string DecimalSeparator
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1809)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1809)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	[DispId(1810)]
	string ThousandsSeparator
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1810)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1810)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	[DispId(1961)]
	bool UseSystemSeparators
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1961)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1961)]
		[param: In]
		set;
	}

	void _VtblGap38_28();

	[DispId(2385)]
	XlGenerateTableRefs GenerateTableRefs
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2385)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2385)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	void _VtblGap39_18();

	[DispId(2776)]
	bool PrintCommunication
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2776)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2776)]
		[param: In]
		set;
	}

	void _VtblGap40_11();

	[DispId(2784)]
	ProtectedViewWindow ActiveProtectedViewWindow
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2784)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
