using System.Runtime.InteropServices;
using Macabacus_Word.Colors;
using Macabacus_Word.Links;

namespace Macabacus_Word;

[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public sealed class clsRouter : IAddInUtilities
{
	public const string ClassId = "88903157-450b-41ba-9738-42f8d23d77d2";

	public const string InterfaceId = "47592a98-0756-4a04-8d25-78d86359a2f1";

	public const string EventsId = "68fe24d2-bb2d-430a-87ca-5b0531bb462f";

	public void CycleFontColor()
	{
		Font.Cycle();
	}

	void IAddInUtilities.CycleFontColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleFontColor
		this.CycleFontColor();
	}

	public void CycleFillColor()
	{
		Fill.Cycle();
	}

	void IAddInUtilities.CycleFillColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleFillColor
		this.CycleFillColor();
	}

	public void CycleBorderColor()
	{
		Border.Cycle();
	}

	void IAddInUtilities.CycleBorderColor()
	{
		//ILSpy generated this explicit interface implementation from .override directive in CycleBorderColor
		this.CycleBorderColor();
	}

	public void StyleCycle1()
	{
		clsStyles.StyleCycle1();
	}

	void IAddInUtilities.StyleCycle1()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle1
		this.StyleCycle1();
	}

	public void StyleCycle2()
	{
		clsStyles.StyleCycle2();
	}

	void IAddInUtilities.StyleCycle2()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle2
		this.StyleCycle2();
	}

	public void StyleCycle3()
	{
		clsStyles.StyleCycle3();
	}

	void IAddInUtilities.StyleCycle3()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle3
		this.StyleCycle3();
	}

	public void StyleCycle4()
	{
		clsStyles.StyleCycle4();
	}

	void IAddInUtilities.StyleCycle4()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle4
		this.StyleCycle4();
	}

	public void StyleCycle5()
	{
		clsStyles.StyleCycle5();
	}

	void IAddInUtilities.StyleCycle5()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle5
		this.StyleCycle5();
	}

	public void StyleCycle6()
	{
		clsStyles.StyleCycle6();
	}

	void IAddInUtilities.StyleCycle6()
	{
		//ILSpy generated this explicit interface implementation from .override directive in StyleCycle6
		this.StyleCycle6();
	}

	public void ImportExcel()
	{
	}

	void IAddInUtilities.ImportExcel()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ImportExcel
		this.ImportExcel();
	}

	public void UpdateLink()
	{
		Refresh.SelectedLinks();
	}

	void IAddInUtilities.UpdateLink()
	{
		//ILSpy generated this explicit interface implementation from .override directive in UpdateLink
		this.UpdateLink();
	}

	public void ViewSource()
	{
		View.ViewSource();
	}

	void IAddInUtilities.ViewSource()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ViewSource
		this.ViewSource();
	}

	public void ZoomIn()
	{
		clsView.ZoomIn();
	}

	void IAddInUtilities.ZoomIn()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ZoomIn
		this.ZoomIn();
	}

	public void ZoomOut()
	{
		clsView.ZoomOut();
	}

	void IAddInUtilities.ZoomOut()
	{
		//ILSpy generated this explicit interface implementation from .override directive in ZoomOut
		this.ZoomOut();
	}

	public void SaveAll()
	{
		clsFile.SaveAll();
	}

	void IAddInUtilities.SaveAll()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SaveAll
		this.SaveAll();
	}

	public void SaveUp()
	{
		clsFile.SaveUp();
	}

	void IAddInUtilities.SaveUp()
	{
		//ILSpy generated this explicit interface implementation from .override directive in SaveUp
		this.SaveUp();
	}

	public void Reopen()
	{
		clsFile.Reopen();
	}

	void IAddInUtilities.Reopen()
	{
		//ILSpy generated this explicit interface implementation from .override directive in Reopen
		this.Reopen();
	}
}
