using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C0398-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface TextFrame2 : _IMsoDispObj
{
	void _VtblGap1_28();

	[DispId(113)]
	MsoTriState HasText
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		get;
	}

	[DispId(114)]
	TextRange2 TextRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(114)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
