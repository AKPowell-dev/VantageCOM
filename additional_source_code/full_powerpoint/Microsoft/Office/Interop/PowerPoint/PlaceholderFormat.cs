using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[Guid("91493477-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
public interface PlaceholderFormat
{
	void _VtblGap1_2();

	[DispId(2003)]
	PpPlaceholderType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2003)]
		get;
	}

	void _VtblGap2_2();

	[DispId(2005)]
	MsoShapeType ContainedType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2005)]
		get;
	}
}
