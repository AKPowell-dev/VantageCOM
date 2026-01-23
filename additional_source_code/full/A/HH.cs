using System.Collections;
using System.Windows.Forms;

namespace A;

internal sealed class HH : IComparer
{
	private int A;

	private SortOrder A;

	public HH()
	{
		this.A = 0;
		A = SortOrder.Ascending;
	}

	public HH(int A, SortOrder B)
	{
		this.A = A;
		this.A = B;
	}

	public int Compare(object x, object y)
	{
		int num = -1;
		num = string.Compare(((ListViewItem)x).SubItems[this.A].Text, ((ListViewItem)y).SubItems[this.A].Text);
		if (A == SortOrder.Descending)
		{
			num = checked(num * -1);
		}
		return num;
	}

	int IComparer.Compare(object x, object y)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Compare
		return this.Compare(x, y);
	}
}
