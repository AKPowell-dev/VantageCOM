using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("5d8ffee9-0105-497d-af15-6dcc5cc78310")]
public interface RibbonGroup : RibbonComponent, IComponent, IDisposable
{
	new void _VtblGap1_5();

	IList<RibbonControl> Items { get; }

	void _VtblGap2_2();

	string Label { get; set; }
}
