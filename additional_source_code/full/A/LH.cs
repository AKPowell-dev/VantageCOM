using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace A;

internal sealed class LH : AxHost
{
	public LH()
		: base(null)
	{
	}

	public static IPictureDisp A(Image A)
	{
		return (IPictureDisp)AxHost.GetIPictureDispFromPicture(A);
	}
}
