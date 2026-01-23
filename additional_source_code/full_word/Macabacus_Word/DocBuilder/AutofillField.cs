using System.Windows;
using A;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.DocBuilder;

public sealed class AutofillField : BaseQuestion
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

	public AutofillField(ContentControl cc, int intIndex, string strText)
		: base(cc, intIndex)
	{
		if (Operators.CompareString(cc.Title, string.Empty, TextCompare: false) != 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (cc.Title.Length != 0)
			{
				goto IL_0096;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		base.Question = intIndex + XC.A(21362) + cc.Tag.Replace(XC.A(21392), "").ToUpper();
		goto IL_0096;
		IL_0096:
		Text = strText;
		base.ApplyButtonVisibility = Visibility.Visible;
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
