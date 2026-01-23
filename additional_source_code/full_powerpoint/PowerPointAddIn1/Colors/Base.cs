using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Colors;

public sealed class Base
{
	public static readonly int TRANSPARENT = 16777215;

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		switch (A.Type)
		{
		case MsoShapeType.msoEmbeddedOLEObject:
		case MsoShapeType.msoLinkedOLEObject:
		case MsoShapeType.msoLinkedPicture:
		case MsoShapeType.msoOLEControlObject:
		case MsoShapeType.msoPicture:
		case MsoShapeType.msoMedia:
		case MsoShapeType.msoScriptAnchor:
		case MsoShapeType.msoInk:
		case MsoShapeType.msoInkComment:
		case MsoShapeType.msoWebVideo:
			return true;
		default:
			return Images.IsGraphic(A);
		}
	}
}
