using System.Windows.Media;
using A;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Pitchly;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Library2.Versioning;

public sealed class ShapeItem : ContentItem
{
	private Microsoft.Office.Interop.PowerPoint.Shape m_A;

	internal Microsoft.Office.Interop.PowerPoint.Shape Shape
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A();
		}
	}

	public ShapeItem(Microsoft.Office.Interop.PowerPoint.Shape shp, ContentInfo ci, ManifestInfo mi)
		: base(ci, mi)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		Shape = shp;
	}

	private void A()
	{
		string source;
		if (this.A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			source = AH.A(59451);
		}
		else if (Images.HasPictureOrGraphic(Shape))
		{
			source = AH.A(61793);
		}
		else
		{
			if (Shape.Type != MsoShapeType.msoChart)
			{
				if (Shape.HasChart != MsoTriState.msoTrue)
				{
					if (Shape.Type != MsoShapeType.msoTable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						if (Shape.HasTable != MsoTriState.msoTrue)
						{
							source = AH.A(62576);
							goto IL_00dc;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					source = AH.A(62126);
					goto IL_00dc;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			source = AH.A(62025);
		}
		goto IL_00dc;
		IL_00dc:
		((ContentItem)this).IconData = Geometry.Parse(source);
	}

	internal bool A()
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		return LibraryAdapter.IsPitchlyLibrary(((ContentItem)this).ContentInfo.LibraryId);
	}
}
