using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C03C6-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface SmartArt : _IMsoDispObj
{
	void _VtblGap1_3();

	[DispId(2)]
	SmartArtNodes AllNodes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(3)]
	SmartArtNodes Nodes
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_8();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(8)]
	void Reset();
}
