using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000C0353-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface LanguageSettings : _IMsoDispObj
{
	void _VtblGap1_2();

	[DispId(1)]
	int LanguageID
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1)]
		get;
	}
}
