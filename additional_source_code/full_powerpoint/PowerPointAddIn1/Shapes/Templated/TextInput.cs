using System.Runtime.CompilerServices;
using System.Windows;
using A;
using Microsoft.Office.Core;

namespace PowerPointAddIn1.Shapes.Templated;

public sealed class TextInput : BaseInput
{
	private string A;

	[CompilerGenerated]
	private TextRange2 A;

	public string Text
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(AH.A(70464));
			if (value.Length > 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						Range.Text = value;
						base.IsPopulated = true;
						return;
					}
				}
			}
			Range.Text = AH.A(15135) + base.Label + AH.A(15138);
			base.IsPopulated = false;
		}
	}

	private TextRange2 Range
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public TextInput(string strLabel, DataTemplate template, TextRange2 rng)
		: base(strLabel, template)
	{
		Range = rng;
	}
}
