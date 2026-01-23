using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace stdole;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[InterfaceType(2)]
[Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
[DefaultMember("Handle")]
public interface IPictureDisp
{
	[DispId(0)]
	int Handle
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		get;
	}
}
