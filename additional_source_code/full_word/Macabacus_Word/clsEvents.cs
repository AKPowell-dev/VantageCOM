using System.Runtime.InteropServices;
using A;
using Macabacus_Word.Library2.Versioning;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word;

public sealed class clsEvents
{
	internal static void A()
	{
		Application application = PC.A.Application;
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).AddEventHandler(application, new ApplicationEvents4_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1882)).AddEventHandler(application, new ApplicationEvents4_DocumentOpenEventHandler(Check.A));
		application = null;
	}

	internal static void B()
	{
		Application application = PC.A.Application;
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(application, new ApplicationEvents4_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1882)).RemoveEventHandler(application, new ApplicationEvents4_DocumentOpenEventHandler(Check.A));
		application = null;
	}
}
