using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("91493488-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface OLEFormat
{
	void _VtblGap1_4();

	[DispId(2005)]
	string ProgID
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
