using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[DefaultMember("Item")]
[CompilerGenerated]
[Guid("000C0363-0000-0000-C000-000000000046")]
public interface FileDialogSelectedItems : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_4();

	[DispId(1610809346)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610809346)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(0)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Item([In] int Index);
}
