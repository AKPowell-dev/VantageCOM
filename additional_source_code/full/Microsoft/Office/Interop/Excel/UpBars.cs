using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[Guid("000208C5-0000-0000-C000-000000000046")]
[InterfaceType(2)]
[TypeIdentifier]
[CompilerGenerated]
public interface UpBars
{
	void _VtblGap1_9();

	[DispId(116)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(116)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
