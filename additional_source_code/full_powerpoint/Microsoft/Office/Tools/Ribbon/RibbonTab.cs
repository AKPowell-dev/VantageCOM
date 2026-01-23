using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[Guid("b2c81178-b119-41af-b358-46460ae005d9")]
[TypeIdentifier]
[CompilerGenerated]
public interface RibbonTab : RibbonComponent, IComponent, IDisposable
{
	IList<RibbonGroup> Groups { get; }

	void _VtblGap1_2();

	string KeyTip { get; set; }

	string Label { get; set; }
}
