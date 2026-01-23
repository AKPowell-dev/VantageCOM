using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("9149348C-5A91-11CF-8700-00AA0060263B")]
public interface ActionSettings : Collection
{
	void _VtblGap1_5();

	[DispId(0)]
	ActionSetting this[[In] PpMouseActivation Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
