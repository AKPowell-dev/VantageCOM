using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000CDB05-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface _CustomXMLPart : _IMsoDispObj
{
	void _VtblGap1_11();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809353)]
	void Delete();
}
