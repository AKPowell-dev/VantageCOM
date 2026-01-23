using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("914934CC-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
[CompilerGenerated]
public interface Pane
{
	void _VtblGap1_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2001)]
	void Activate();

	[DispId(2002)]
	MsoTriState Active
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2002)]
		get;
	}

	void _VtblGap2_1();

	[DispId(2004)]
	PpViewType ViewType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
	}
}
