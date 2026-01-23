using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools;

[ComImport]
[TypeIdentifier]
[Guid("30a90086-8c89-4e19-8299-47765d808408")]
[CompilerGenerated]
public interface SmartTagCollection : IEnumerable, IDisposable
{
	void BeginInit();

	void EndInit();

	void _VtblGap1_7();

	SmartTagBase this[int index] { get; }
}
