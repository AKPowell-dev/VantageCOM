using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Colors;

public sealed class Base
{
	public static readonly int TRANSPARENT = 16777215;

	public static bool IgnoreShapeType(Microsoft.Office.Interop.Word.Shape shp)
	{
		switch (shp.Type)
		{
		case MsoShapeType.msoWebVideo:
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			goto case MsoShapeType.msoEmbeddedOLEObject;
		case MsoShapeType.msoEmbeddedOLEObject:
		case MsoShapeType.msoLinkedOLEObject:
		case MsoShapeType.msoLinkedPicture:
		case MsoShapeType.msoOLEControlObject:
		case MsoShapeType.msoPicture:
		case MsoShapeType.msoMedia:
		case MsoShapeType.msoScriptAnchor:
		case MsoShapeType.msoInk:
		case MsoShapeType.msoInkComment:
			return true;
		default:
			if (shp.Type == (MsoShapeType)28)
			{
				return true;
			}
			return false;
		}
	}

	public static bool IgnoreShapeType(InlineShape shp)
	{
		WdInlineShapeType type = shp.Type;
		if ((uint)(type - 1) > 4u)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if ((uint)(type - 10) > 1u)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				if (type != WdInlineShapeType.wdInlineShapeWebVideo)
				{
					return false;
				}
			}
		}
		return true;
	}
}
