using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("Range")]
[TypeIdentifier]
[Guid("00020957-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface Paragraph
{
	[DispId(0)]
	Range Range
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
