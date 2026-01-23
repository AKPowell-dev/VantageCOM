using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[InterfaceType(2)]
[Guid("00024435-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface ChartFillFormat
{
	void _VtblGap1_23();

	[DispId(558)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		get;
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(558)]
		set;
	}
}
