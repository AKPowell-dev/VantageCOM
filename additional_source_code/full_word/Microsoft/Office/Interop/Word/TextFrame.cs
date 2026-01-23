using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("000209B2-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface TextFrame
{
	void _VtblGap1_13();

	[DispId(1001)]
	Range TextRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1001)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_6();

	[DispId(5008)]
	int HasText
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(5008)]
		get;
	}
}
