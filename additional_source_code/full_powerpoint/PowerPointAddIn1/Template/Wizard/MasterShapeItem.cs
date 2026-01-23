using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.MasterShapes;

namespace PowerPointAddIn1.Template.Wizard;

public sealed class MasterShapeItem
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	[CompilerGenerated]
	private string C;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Shape A;

	public string Name
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public string Placement
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public string Placeholders
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public Microsoft.Office.Interop.PowerPoint.Shape Shape
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public MasterShapeItem(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Shape = shp;
		Name = shp.Name;
		switch (Base.A(shp.Name))
		{
		case Behavior.AllLayouts:
			Placement = AH.A(120511);
			break;
		case Behavior.AllSlides:
			Placement = AH.A(120534);
			break;
		case Behavior.ContentLayouts:
			Placement = AH.A(120555);
			break;
		case Behavior.ContentSlides:
			Placement = AH.A(120586);
			break;
		case Behavior.DynamicLayouts:
			Placement = AH.A(120615);
			break;
		case Behavior.DynamicSlides:
			Placement = AH.A(120646);
			break;
		case Behavior.LayoutsShowingBackgroundGraphics:
			Placement = AH.A(120675);
			break;
		case Behavior.SelectedSlides:
			Placement = AH.A(120746);
			break;
		case Behavior.SlidesShowingBackgroundGraphics:
			Placeholders = AH.A(120777);
			break;
		case Behavior.SpecialLayouts:
			Placeholders = AH.A(120846);
			break;
		case Behavior.SpecialSlides:
			Placeholders = AH.A(120877);
			break;
		}
		List<string> list = new List<string>();
		Placeholders = "";
		if (shp.HasTextFrame == MsoTriState.msoTrue && shp.TextFrame2.HasText == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (2)
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
			TextRange2 textRange = shp.TextFrame2.TextRange;
			string[] array = new string[14]
			{
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_TITLE,
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SECTION,
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SUBSECTION,
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SEC_INDEX,
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SUBSEC_INDEX,
				PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_STAMP,
				AH.A(120906),
				AH.A(120927),
				AH.A(120940),
				AH.A(120953),
				AH.A(120970),
				AH.A(120983),
				AH.A(120996),
				AH.A(121013)
			};
			foreach (string text in array)
			{
				if (!textRange.Text.Contains(text))
				{
					continue;
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
				list.Add(text);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			textRange = null;
			if (list.Count > 0)
			{
				Placeholders = string.Join(AH.A(14258), list.ToArray());
			}
		}
		list = null;
	}
}
