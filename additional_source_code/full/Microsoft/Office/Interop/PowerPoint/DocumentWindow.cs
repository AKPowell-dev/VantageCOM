using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[DefaultMember("Caption")]
[TypeIdentifier]
[CompilerGenerated]
[Guid("91493457-5A91-11CF-8700-00AA0060263B")]
public interface DocumentWindow
{
	void _VtblGap1_2();

	[DispId(2003)]
	Selection Selection
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_9();

	[DispId(0)]
	string Caption
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
