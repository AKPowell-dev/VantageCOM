using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[TypeIdentifier("00020813-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Excel.AppEvents_Event")]
[ComEventInterface(typeof(AppEvents), typeof(AppEvents))]
public interface AppEvents_Event
{
	event AppEvents_NewWorkbookEventHandler NewWorkbook;

	event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange;

	void _VtblGap1_4();

	event AppEvents_SheetActivateEventHandler SheetActivate;

	void _VtblGap2_4();

	event AppEvents_SheetChangeEventHandler SheetChange;

	event AppEvents_WorkbookOpenEventHandler WorkbookOpen;

	event AppEvents_WorkbookActivateEventHandler WorkbookActivate;

	void _VtblGap3_2();

	event AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeClose;

	event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSave;

	void _VtblGap4_10();

	event AppEvents_WindowActivateEventHandler WindowActivate;

	void _VtblGap5_22();

	event AppEvents_AfterCalculateEventHandler AfterCalculate;

	void _VtblGap6_8();

	event AppEvents_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpen;

	void _VtblGap7_2();

	event AppEvents_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeClose;

	void _VtblGap8_2();

	event AppEvents_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivate;

	event AppEvents_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivate;
}
