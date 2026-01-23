using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("00024413-0000-0000-C000-000000000046")]
[InterfaceType(2)]
[TypeIdentifier]
[CompilerGenerated]
public interface AppEvents
{
	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1565)]
	void NewWorkbook([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1558)]
	void SheetSelectionChange([In][MarshalAs(UnmanagedType.IDispatch)] object Sh, [In][MarshalAs(UnmanagedType.Interface)] Range Target);

	void _VtblGap1_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1561)]
	void SheetActivate([In][MarshalAs(UnmanagedType.IDispatch)] object Sh);

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1564)]
	void SheetChange([In][MarshalAs(UnmanagedType.IDispatch)] object Sh, [In][MarshalAs(UnmanagedType.Interface)] Range Target);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1567)]
	void WorkbookOpen([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1568)]
	void WorkbookActivate([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1570)]
	void WorkbookBeforeClose([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb, [In][Out] ref bool Cancel);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1571)]
	void WorkbookBeforeSave([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb, [In] bool SaveAsUI, [In][Out] ref bool Cancel);

	void _VtblGap4_5();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1556)]
	void WindowActivate([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb, [In][MarshalAs(UnmanagedType.Interface)] Window Wn);

	void _VtblGap5_11();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2612)]
	void AfterCalculate();

	void _VtblGap6_4();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2903)]
	void ProtectedViewWindowOpen([In][MarshalAs(UnmanagedType.Interface)] ProtectedViewWindow Pvw);

	void _VtblGap7_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2906)]
	void ProtectedViewWindowBeforeClose([In][MarshalAs(UnmanagedType.Interface)] ProtectedViewWindow Pvw, [In] XlProtectedViewCloseReason Reason, [In][Out] ref bool Cancel);

	void _VtblGap8_1();

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2909)]
	void ProtectedViewWindowActivate([In][MarshalAs(UnmanagedType.Interface)] ProtectedViewWindow Pvw);

	[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2910)]
	void ProtectedViewWindowDeactivate([In][MarshalAs(UnmanagedType.Interface)] ProtectedViewWindow Pvw);
}
