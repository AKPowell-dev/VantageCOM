using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Pagination;

public sealed class BlankSlides
{
	public static readonly string BLANK_NAME = AH.A(101151);

	public static CustomLayout GetBlankLayout(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		return Helpers.GetLayout(pres, SlideType.Blank);
	}

	public static CustomLayout CreateBlankLayout(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		Master slideMaster = pres.Designs.Add(BLANK_NAME).SlideMaster;
		for (int i = slideMaster.CustomLayouts.Count; i >= 2; i = checked(i + -1))
		{
			slideMaster.CustomLayouts[i].Delete();
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			CustomLayout customLayout = slideMaster.CustomLayouts[1];
			customLayout.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
			customLayout.Name = BLANK_NAME;
			slideMaster.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
			slideMaster = null;
			return customLayout;
		}
	}

	public static List<int> InsertAsNeeded(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		int count = pres.Slides.Count;
		FlySheetStyle presentationFlysheetStyle = Behavior.GetPresentationFlysheetStyle(pres);
		List<int> list = new List<int>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = pres.Slides.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Slide sld = (Slide)enumerator.Current;
				SlideNumbers.Freeze(sld);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		CustomLayout customLayout = GetBlankLayout(pres);
		if (customLayout == null)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			customLayout = CreateBlankLayout(pres);
		}
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				Slide sld = pres.Slides[i];
				if (A(sld) && IsSectionFlysheet(sld, pres, presentationFlysheetStyle))
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
					pres.Slides.AddSlide(i, customLayout);
					list.Add(i);
					i++;
				}
				else
				{
					if (!B(sld))
					{
						continue;
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
					if (!FacingSlides.IsFacingSlide(sld))
					{
						continue;
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
					pres.Slides.AddSlide(i, customLayout);
					list.Add(i);
					i++;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				Slide sld = null;
				customLayout = null;
				return list;
			}
		}
	}

	private static bool A(Slide A)
	{
		return A.SlideIndex % 2 == 0;
	}

	private static bool B(Slide A)
	{
		return !BlankSlides.A(A);
	}

	public static bool IsSectionFlysheet(Slide sld, Microsoft.Office.Interop.PowerPoint.Presentation pres, FlySheetStyle flysheetStyle)
	{
		bool flag;
		if (flysheetStyle == FlySheetStyle.Topic)
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
			flag = Helpers.GetSlideType(sld) == SlideType.Flysheet;
		}
		else
		{
			flag = Helpers.GetSlideType(sld) == SlideType.Agenda;
		}
		if (flag)
		{
			SectionProperties sectionProperties = pres.SectionProperties;
			if (sectionProperties.Count > 0)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					return !Update.A(sectionProperties.Name(sld.sectionIndex));
				}
			}
			sectionProperties = null;
		}
		return false;
	}
}
