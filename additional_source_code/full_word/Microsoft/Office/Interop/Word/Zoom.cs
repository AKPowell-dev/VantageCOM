using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("Percentage")]
[Guid("000209A6-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface Zoom
{
	void _VtblGap1_3();

	[DispId(0)]
	int Percentage
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		set;
	}
}
