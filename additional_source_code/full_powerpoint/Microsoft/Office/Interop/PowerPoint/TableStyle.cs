using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("914934F5-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface TableStyle
{
	void _VtblGap1_1();

	[DispId(2002)]
	string Id
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2002)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
