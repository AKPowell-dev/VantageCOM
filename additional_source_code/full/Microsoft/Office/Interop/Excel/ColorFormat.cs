using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[TypeIdentifier]
[DefaultMember("RGB")]
[CompilerGenerated]
[Guid("000C0312-0000-0000-C000-000000000046")]
public interface ColorFormat : _IMsoDispObj
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

	void _VtblGap2_5();

	[DispId(104)]
	MsoThemeColorIndex ObjectThemeColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(104)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(104)]
		[param: In]
		set;
	}
}
