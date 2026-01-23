using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("DCE9F2C4-4C02-43BA-840E-B4276550EF79")]
public interface DataTable
{
	void _VtblGap1_9();

	[DispId(1610743817)]
	ChartFont Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743817)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_5();

	[DispId(1610743823)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743823)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
