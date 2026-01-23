using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;

namespace Microsoft.Office.Tools;

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("857D6117-1E7B-422B-A9E3-F907304C709A")]
[TypeIdentifier]
[CompilerGenerated]
public interface Factory
{
	RibbonFactory GetRibbonFactory();

	void _VtblGap1_1();

	CustomTaskPaneCollection CreateCustomTaskPaneCollection(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent);

	SmartTagCollection CreateSmartTagCollection(IServiceProvider serviceProvider, IHostItemProvider hostItemProvider, string primaryCookie, string identifier, object containerComponent);
}
