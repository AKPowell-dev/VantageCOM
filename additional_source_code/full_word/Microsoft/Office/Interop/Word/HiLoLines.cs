using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[TypeIdentifier]
[Guid("7A1BCE11-5783-4C7D-BD02-F3D84AB40E7F")]
public interface HiLoLines
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
