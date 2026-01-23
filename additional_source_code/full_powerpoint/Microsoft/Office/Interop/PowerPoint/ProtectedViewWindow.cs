using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("BA72E55A-4FF5-48F4-8215-5505F990966F")]
[CompilerGenerated]
[TypeIdentifier]
[DefaultMember("Caption")]
public interface ProtectedViewWindow
{
	void _VtblGap1_6();

	[DispId(0)]
	string Caption
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
