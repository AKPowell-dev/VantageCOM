using System.Collections.Generic;
using System.Windows;
using A;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.DocBuilder;

public sealed class YesNoQuestion : BaseQuestion
{
	private List<Choice> A;

	public List<Choice> Choices
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public YesNoQuestion(ContentControl cc, int intIndex)
		: base(cc, intIndex)
	{
		Choices = new List<Choice>();
		Choices.Add(new Choice(this, cc, XC.A(22261), 0));
		Choices.Add(new Choice(this, cc, XC.A(22268), 0));
		Choices[0].CornerRadius = new CornerRadius(0.0);
	}
}
