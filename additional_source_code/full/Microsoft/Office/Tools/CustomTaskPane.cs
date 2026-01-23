using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace Microsoft.Office.Tools;

[ComImport]
[TypeIdentifier]
[Guid("881b42fd-484d-4494-8500-779de4e4aac1")]
[CompilerGenerated]
public interface CustomTaskPane : IDisposable
{
	event EventHandler VisibleChanged;

	event EventHandler DockPositionChanged;

	UserControl Control { get; }

	string Title { get; }

	object Window { get; }

	MsoCTPDockPosition DockPosition { get; set; }

	MsoCTPDockPositionRestrict DockPositionRestrict { get; set; }

	int Width { get; set; }

	int Height { get; set; }

	bool Visible { get; set; }
}
