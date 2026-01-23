using System.Runtime.InteropServices;
using A;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Library2.Insert;

public sealed class Slides
{
	internal static void A(Application A)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(A, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
	}

	internal static void B(Application A)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(A, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
	}
}
