using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[InterfaceType(2)]
[CompilerGenerated]
[Guid("0002445D-0000-0000-C000-000000000046")]
public interface Error
{
	void _VtblGap1_3();

	[DispId(6)]
	bool Value
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(6)]
		get;
	}
}
