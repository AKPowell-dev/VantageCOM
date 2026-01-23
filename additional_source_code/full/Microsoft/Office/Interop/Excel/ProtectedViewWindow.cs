using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[DefaultMember("_Default")]
[CompilerGenerated]
[InterfaceType(2)]
[TypeIdentifier]
[Guid("000244CD-0000-0000-C000-000000000046")]
public interface ProtectedViewWindow
{
	[DispId(0)]
	string _Default
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}
}
