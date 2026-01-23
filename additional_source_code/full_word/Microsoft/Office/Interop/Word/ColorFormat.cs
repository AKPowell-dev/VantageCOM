using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("RGB")]
[TypeIdentifier]
[Guid("000209C6-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface ColorFormat
{
	void _VtblGap1_3();

	[DispId(0)]
	int RGB
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		set;
	}

	void _VtblGap2_2();

	[DispId(101)]
	MsoColorType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		get;
	}
}
