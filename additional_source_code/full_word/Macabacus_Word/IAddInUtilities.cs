using System.Runtime.InteropServices;

namespace Macabacus_Word;

[ComVisible(true)]
public interface IAddInUtilities
{
	void CycleFontColor();

	void CycleFillColor();

	void CycleBorderColor();

	void StyleCycle1();

	void StyleCycle2();

	void StyleCycle3();

	void StyleCycle4();

	void StyleCycle5();

	void StyleCycle6();

	void ImportExcel();

	void UpdateLink();

	void ViewSource();

	void ZoomIn();

	void ZoomOut();

	void SaveAll();

	void SaveUp();

	void Reopen();
}
