using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace stdole;

[ComImport]
[DefaultMember("Handle")]
[CompilerGenerated]
[TypeIdentifier]
[Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
[InterfaceType(2)]
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
