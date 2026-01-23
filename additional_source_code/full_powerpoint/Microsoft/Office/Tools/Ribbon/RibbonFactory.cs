using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[TypeIdentifier]
[CompilerGenerated]
[Guid("1012BDD2-303F-4464-A64B-3026BD91C31E")]
public interface RibbonFactory
{
	void _VtblGap1_4();

	RibbonButton CreateRibbonButton();

	void _VtblGap2_7();

	RibbonGroup CreateRibbonGroup();

	void _VtblGap3_1();

	RibbonMenu CreateRibbonMenu();

	RibbonSeparator CreateRibbonSeparator();

	void _VtblGap4_1();

	RibbonTab CreateRibbonTab();
}
