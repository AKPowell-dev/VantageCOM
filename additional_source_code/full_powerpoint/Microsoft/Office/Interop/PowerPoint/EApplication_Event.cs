using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[TypeIdentifier("91493440-5a91-11cf-8700-00aa0060263b", "Microsoft.Office.Interop.PowerPoint.EApplication_Event")]
[ComEventInterface(typeof(EApplication), typeof(EApplication))]
public interface EApplication_Event
{
	event EApplication_WindowSelectionChangeEventHandler WindowSelectionChange;

	void _VtblGap1_6();

	event EApplication_PresentationSaveEventHandler PresentationSave;

	event EApplication_PresentationOpenEventHandler PresentationOpen;

	event EApplication_NewPresentationEventHandler NewPresentation;

	event EApplication_PresentationNewSlideEventHandler PresentationNewSlide;

	event EApplication_WindowActivateEventHandler WindowActivate;

	void _VtblGap2_12();

	event EApplication_SlideSelectionChangedEventHandler SlideSelectionChanged;

	void _VtblGap3_2();

	event EApplication_PresentationBeforeSaveEventHandler PresentationBeforeSave;

	void _VtblGap4_2();

	event EApplication_AfterNewPresentationEventHandler AfterNewPresentation;

	event EApplication_AfterPresentationOpenEventHandler AfterPresentationOpen;

	void _VtblGap5_6();

	event EApplication_PresentationBeforeCloseEventHandler PresentationBeforeClose;

	void _VtblGap6_10();

	event EApplication_PresentationCloseFinalEventHandler PresentationCloseFinal;

	void _VtblGap7_2();

	event EApplication_AfterShapeSizeChangeEventHandler AfterShapeSizeChange;
}
