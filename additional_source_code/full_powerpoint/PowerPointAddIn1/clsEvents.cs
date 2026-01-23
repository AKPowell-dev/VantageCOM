using System.Runtime.InteropServices;
using A;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.Library2;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Template;
using PowerPointAddIn1.TurboShapes;

namespace PowerPointAddIn1;

public sealed class clsEvents
{
	public static void Add()
	{
		Application application = NG.A.Application;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(69711)).AddEventHandler(application, new EApplication_PresentationOpenEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).AddEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Access.AfterPresentationOpen));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).AddEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Check.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).AddEventHandler(application, new EApplication_NewPresentationEventHandler(Create.Application_NewPresentation));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).AddEventHandler(application, new EApplication_AfterNewPresentationEventHandler(Create.Application_AfterNewPresentation));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(application, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(clsRibbon.Application_PresentationCloseFinal));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105591)).AddEventHandler(application, new EApplication_WindowActivateEventHandler(clsRibbon.Application_WindowActivate));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(B));
		application = null;
		Base.AddSelectionChangedEvent();
	}

	public static void Remove()
	{
		Application application = NG.A.Application;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(69711)).RemoveEventHandler(application, new EApplication_PresentationOpenEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).RemoveEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Access.AfterPresentationOpen));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).RemoveEventHandler(application, new EApplication_AfterPresentationOpenEventHandler(Check.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105519)).RemoveEventHandler(application, new EApplication_NewPresentationEventHandler(Create.Application_NewPresentation));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).RemoveEventHandler(application, new EApplication_AfterNewPresentationEventHandler(Create.Application_AfterNewPresentation));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(application, new EApplication_PresentationNewSlideEventHandler(Create.Application_PresentationNewSlide));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(clsRibbon.Application_PresentationCloseFinal));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105591)).RemoveEventHandler(application, new EApplication_WindowActivateEventHandler(clsRibbon.Application_WindowActivate));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(B));
		application = null;
		Base.RemoveSelectionChangedEvent();
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		clsRibbon.InvalidateOpenPresentationRequiredControls();
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		if (IG.A(NG.A.Application.Presentations) != 1)
		{
			return;
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
			clsPanes.RemoveTaskPanesByTitle(PowerPointAddIn1.Explorer.Pane.PANE_TITLE);
			return;
		}
	}
}
