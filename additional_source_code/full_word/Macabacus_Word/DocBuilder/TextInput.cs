using A;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.DocBuilder;

public sealed class TextInput : BaseQuestion
{
	private string A;

	public string Text
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
			NotifyPropertyChanged(XC.A(14361));
		}
	}

	public TextInput(ContentControl cc, int intIndex)
		: base(cc, intIndex)
	{
	}

	public void Apply(string strText)
	{
		ContentControl contentControl = base.ContentControl;
		contentControl.LockContents = false;
		contentControl.LockContentControl = false;
		contentControl.Range.Text = strText;
		contentControl.Delete();
		_ = null;
	}
}
