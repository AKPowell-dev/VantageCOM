using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("91493472-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface SoundEffect
{
	void _VtblGap1_4();

	[DispId(2004)]
	PpSoundEffectType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2004)]
		[param: In]
		set;
	}
}
