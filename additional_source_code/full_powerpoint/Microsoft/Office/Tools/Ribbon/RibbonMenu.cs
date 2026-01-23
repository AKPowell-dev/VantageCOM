using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("59dc7f42-aca2-484a-9622-1ee34a6cfd7d")]
public interface RibbonMenu : RibbonControl, RibbonComponent, IComponent, IDisposable
{
	void _VtblGap1_2();

	IList<RibbonControl> Items { get; }

	RibbonControlSize ControlSize { get; set; }

	void _VtblGap2_10();

	string KeyTip { get; set; }

	string Label { get; set; }

	string OfficeImageId { get; set; }

	void _VtblGap3_4();

	bool ShowImage { get; set; }
}
