using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[CompilerGenerated]
[Guid("000208D4-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface AutoCorrect
{
	void _VtblGap1_19();

	[DispId(2294)]
	bool AutoExpandListRange
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2294)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2294)]
		set;
	}
}
