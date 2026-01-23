using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("FC9090AF-0DDB-4EC1-86E8-8751F2199F2C")]
[CompilerGenerated]
public interface Gridlines
{
	void _VtblGap1_5();

	[DispId(1610743813)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743813)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
