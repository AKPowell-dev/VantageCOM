using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("9149348D-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface ActionSetting
{
	void _VtblGap1_1();

	[DispId(2002)]
	object Parent
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2002)]
		[return: MarshalAs(UnmanagedType.IDispatch)]
		get;
	}

	void _VtblGap2_10();

	[DispId(2008)]
	Hyperlink Hyperlink
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2008)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
