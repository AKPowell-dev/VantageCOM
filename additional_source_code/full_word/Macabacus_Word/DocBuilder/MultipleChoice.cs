using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.DocBuilder;

public sealed class MultipleChoice : BaseQuestion
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

	public MultipleChoice(ContentControl cc, int intIndex)
		: base(cc, intIndex)
	{
		Choices = new List<Choice>();
	}
}
