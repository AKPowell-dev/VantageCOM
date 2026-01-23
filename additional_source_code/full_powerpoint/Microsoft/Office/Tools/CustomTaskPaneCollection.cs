using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Microsoft.Office.Tools;

[ComImport]
[Guid("c3a84bf1-e95b-4d23-952d-59e35673958e")]
[TypeIdentifier]
[CompilerGenerated]
public interface CustomTaskPaneCollection : IEnumerable<CustomTaskPane>, IEnumerable, IDisposable
{
	void BeginInit();

	void EndInit();

	int Count { get; }

	CustomTaskPane this[int index] { get; }

	void _VtblGap1_1();

	CustomTaskPane Add(UserControl control, string title, object window);

	bool Remove(CustomTaskPane customTaskPane);

	void RemoveAt(int index);
}
