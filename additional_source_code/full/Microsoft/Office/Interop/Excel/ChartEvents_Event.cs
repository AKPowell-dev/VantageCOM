using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[ComEventInterface(typeof(ChartEvents), typeof(ChartEvents))]
[CompilerGenerated]
[TypeIdentifier("00020813-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Excel.ChartEvents_Event")]
public interface ChartEvents_Event
{
	void _VtblGap1_2();

	event ChartEvents_DeactivateEventHandler Deactivate;

	void _VtblGap2_4();

	event ChartEvents_MouseUpEventHandler MouseUp;
}
