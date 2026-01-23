using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace A;

internal sealed class OC : AxHost
{
	public OC()
		: base(null)
	{
	}

	public static IPictureDisp A(Image A)
	{
		return (IPictureDisp)AxHost.GetIPictureDispFromPicture(A);
	}
}
