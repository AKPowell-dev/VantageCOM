using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("09b06894-74de-44ff-9d48-9661ae639f41")]
public interface RibbonComponent : IComponent, IDisposable
{
	void _VtblGap1_5();

	string Name { get; set; }

	void ResumeLayout(bool performLayout);

	void _VtblGap2_1();

	void PerformLayout();

	void _VtblGap3_1();

	void SuspendLayout();
}
