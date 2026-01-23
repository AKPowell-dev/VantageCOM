using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[CompilerGenerated]
[TypeIdentifier("00020905-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler")]
public delegate void ApplicationEvents4_DocumentBeforeSaveEventHandler([In][MarshalAs(UnmanagedType.Interface)] Document Doc, [In] ref bool SaveAsUI, [In][Out] ref bool Cancel);
