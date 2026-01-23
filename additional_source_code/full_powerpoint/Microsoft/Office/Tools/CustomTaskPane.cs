using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace Microsoft.Office.Tools;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("881b42fd-484d-4494-8500-779de4e4aac1")]
public interface CustomTaskPane : IDisposable
{
	event EventHandler VisibleChanged;

	void _VtblGap1_2();

	UserControl Control { get; }

	string Title { get; }

	void _VtblGap2_1();

	MsoCTPDockPosition DockPosition { get; set; }

	MsoCTPDockPositionRestrict DockPositionRestrict { get; set; }

	int Width { get; set; }

	void _VtblGap3_2();

	bool Visible { get; set; }
}
