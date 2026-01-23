using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[TypeIdentifier("00020813-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler")]
[CompilerGenerated]
public delegate void AppEvents_WorkbookBeforeCloseEventHandler([In][MarshalAs(UnmanagedType.Interface)] Workbook Wb, [In][Out] ref bool Cancel);
