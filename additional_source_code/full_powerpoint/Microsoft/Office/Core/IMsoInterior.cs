using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C171B-0000-0000-C000-000000000046")]
[CompilerGenerated]
[TypeIdentifier]
public interface IMsoInterior
{
	[DispId(1610743808)]
	object Color
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743808)]
		[return: MarshalAs(UnmanagedType.Struct)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743808)]
		[param: In]
		[param: MarshalAs(UnmanagedType.Struct)]
		set;
	}
}
