using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[Guid("b57e6217-33f2-46bf-9625-c313526de60c")]
[TypeIdentifier]
[CompilerGenerated]
public interface RibbonButton : RibbonControl, RibbonComponent, IComponent, IDisposable
{
	event RibbonControlEventHandler Click;

	RibbonControlSize ControlSize { get; set; }

	void _VtblGap1_4();

	Image Image { get; set; }

	string ImageName { get; set; }

	string KeyTip { get; set; }

	string Label { get; set; }

	string OfficeImageId { get; set; }

	void _VtblGap2_2();

	bool ShowImage { get; set; }
}
