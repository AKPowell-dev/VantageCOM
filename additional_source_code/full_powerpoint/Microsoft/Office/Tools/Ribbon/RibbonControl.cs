using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Tools.Ribbon;

[ComImport]
[Guid("19509b89-9091-4063-a495-f688299c8c3a")]
[TypeIdentifier]
[CompilerGenerated]
public interface RibbonControl : RibbonComponent, IComponent, IDisposable
{
	bool Enabled { get; set; }
}
