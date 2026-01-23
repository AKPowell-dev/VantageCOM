using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C03C0-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface GradientStops : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_2();

	[DispId(0)]
	GradientStop this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	void _VtblGap2_3();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(12)]
	void Insert2([In] int RGB, [In] float Position, [In] float Transparency = 0f, [In] int Index = -1, [In] float Brightness = 0f);
}
