using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck;

public abstract class BaseCheck
{
	public abstract void Check(Slide sld, Shape shp);
}
