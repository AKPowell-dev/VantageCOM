using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("000208DA-0000-0000-C000-000000000046")]
public interface _Workbook
{
	[DispId(148)]
	Application Application
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(148)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap1_4();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(0)]
	[DispId(304)]
	void Activate();

	void _VtblGap2_1();

	[DispId(307)]
	object ActiveSheet
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(307)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap3_8();

	[DispId(1176)]
	object BuiltinDocumentProperties
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1176)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap4_2();

	[DispId(121)]
	Sheets Charts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(121)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(3)]
	[DispId(277)]
	void Close([Optional][In][MarshalAs(UnmanagedType.Struct)] object SaveChanges, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object RouteWorkbook);

	void _VtblGap5_12();

	[DispId(1177)]
	object CustomDocumentProperties
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1177)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap6_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(397)]
	[LCIDConversion(1)]
	void DeleteNumberFormat([In][MarshalAs(UnmanagedType.BStr)] string NumberFormat);

	void _VtblGap7_1();

	[DispId(404)]
	XlDisplayDrawingObjects DisplayDrawingObjects
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(404)]
		[LCIDConversion(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(404)]
		[LCIDConversion(0)]
		[param: In]
		set;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1168)]
	[LCIDConversion(0)]
	bool ExclusiveAccess();

	[DispId(288)]
	XlFileFormat FileFormat
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(288)]
		[LCIDConversion(0)]
		get;
	}

	void _VtblGap8_1();

	[DispId(289)]
	string FullName
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(289)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap9_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(808)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object LinkSources([Optional][In][MarshalAs(UnmanagedType.Struct)] object Type);

	void _VtblGap10_3();

	[DispId(1169)]
	bool MultiUserEditing
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(1169)]
		get;
	}

	[DispId(110)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[DispId(442)]
	Names Names
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(442)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap11_8();

	[DispId(291)]
	string Path
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(291)]
		[LCIDConversion(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap12_14();

	[DispId(296)]
	bool ReadOnly
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(296)]
		[LCIDConversion(0)]
		get;
	}

	void _VtblGap13_10();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(283)]
	[LCIDConversion(0)]
	void Save();

	void _VtblGap14_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(1)]
	[DispId(175)]
	void SaveCopyAs([Optional][In][MarshalAs(UnmanagedType.Struct)] object Filename);

	[DispId(298)]
	bool Saved
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(298)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[LCIDConversion(0)]
		[DispId(298)]
		[param: In]
		set;
	}

	void _VtblGap15_5();

	[DispId(485)]
	Sheets Sheets
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(485)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap16_2();

	[DispId(493)]
	Styles Styles
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(493)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap17_14();

	[DispId(430)]
	Windows Windows
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(430)]
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

	void _VtblGap18_40();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[LCIDConversion(12)]
	[DispId(1925)]
	void SaveAs([Optional][In][MarshalAs(UnmanagedType.Struct)] object Filename, [Optional][In][MarshalAs(UnmanagedType.Struct)] object FileFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Password, [Optional][In][MarshalAs(UnmanagedType.Struct)] object WriteResPassword, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ReadOnlyRecommended, [Optional][In][MarshalAs(UnmanagedType.Struct)] object CreateBackup, [In] XlSaveAsAccessMode AccessMode = XlSaveAsAccessMode.xlNoChange, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ConflictResolution, [Optional][In][MarshalAs(UnmanagedType.Struct)] object AddToMru, [Optional][In][MarshalAs(UnmanagedType.Struct)] object TextCodepage, [Optional][In][MarshalAs(UnmanagedType.Struct)] object TextVisualLayout, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Local);

	void _VtblGap19_2();

	[DispId(2050)]
	bool RemovePersonalInformation
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2050)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2050)]
		[param: In]
		set;
	}

	void _VtblGap20_38();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2514)]
	void RemoveDocumentInformation([In] XlRemoveDocInfoType RemoveDocInfoType);

	void _VtblGap21_17();

	[DispId(2530)]
	CustomXMLParts CustomXMLParts
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2530)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2531)]
	bool Final
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2531)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2531)]
		[param: In]
		set;
	}

	void _VtblGap22_1();

	[DispId(2533)]
	OfficeTheme Theme
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2533)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
