using System.Windows.Media;
using A;
using ExcelAddIn1.Shapes;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Library2.Versioning;

public sealed class ShapeItem : ContentItem
{
	private Shape m_A;

	internal Shape Shape
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

	internal ShapeItem(Shape A, ContentInfo B, ManifestInfo C)
		: base(B, C)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		Shape = A;
	}

	private void A()
	{
		Shape shape = Shape;
		string source;
		if (Images.A(shape))
		{
			source = VH.A(83164);
		}
		else
		{
			if (shape.Type != MsoShapeType.msoChart)
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
				if (shape.HasChart != MsoTriState.msoTrue)
				{
					if (shape.Type == MsoShapeType.msoTable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
						source = VH.A(83396);
					}
					else
					{
						source = VH.A(83846);
					}
					goto IL_0095;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			source = VH.A(41656);
		}
		goto IL_0095;
		IL_0095:
		shape = null;
		((ContentItem)this).IconData = Geometry.Parse(source);
	}
}
