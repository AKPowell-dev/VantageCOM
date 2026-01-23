using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using ExcelAddIn1.Audit.Check.Observations;

namespace A;

internal sealed class FC : Observation
{
	[CompilerGenerated]
	internal sealed class EC
	{
		public Observation A;

		[SpecialName]
		internal bool A(Observation A)
		{
			return A.Sheet == this.A.Sheet;
		}

		[SpecialName]
		internal bool B(Observation A)
		{
			return A.Chart == this.A.Chart;
		}

		[SpecialName]
		internal bool C(Observation A)
		{
			return A.Worksheet == this.A.Worksheet;
		}
	}

	internal override List<Observation> Children
	{
		get
		{
			return base.Children;
		}
		set
		{
			base.Children = value;
			B();
		}
	}

	internal FC(IEnumerable<Observation> A)
	{
		base.AffectsGroupCount = false;
		Children = A.ToList();
		Observation A2 = Children.First();
		base.Category = A2.Category;
		base.Severity = A2.Severity;
		base.Title = A2.Title;
		base.SubtitleVisibility = Visibility.Visible;
		if (Children.All([SpecialName] (Observation observation) => observation.Sheet == A2.Sheet))
		{
			while (true)
			{
				switch (1)
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
			base.Sheet = RuntimeHelpers.GetObjectValue(A2.Sheet);
		}
		if (Children.All([SpecialName] (Observation observation) => observation.Chart == A2.Chart))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			base.Chart = A2.Chart;
		}
		if (Children.All([SpecialName] (Observation observation) => observation.Worksheet == A2.Worksheet))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			base.Worksheet = A2.Worksheet;
		}
		base.TypeColor = A2.TypeColor;
		base.Explanation = A2.Explanation;
		base.Icon = A2.Icon;
		base.IconPadding = A2.IconPadding;
		base.TooltipVisibility = A2.TooltipVisibility;
		base.ErrorsCount = A2.ErrorsCount;
		base.WarningsCount = A2.WarningsCount;
		base.MessagesCount = A2.MessagesCount;
		base.HasFix = false;
		base.FixMenuVisibility = Visibility.Hidden;
		base.FixVisibility = Visibility.Hidden;
	}

	internal override bool A(Observation A)
	{
		if (!base.A(A))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return false;
				}
			}
		}
		B();
		return true;
	}

	private new void B()
	{
		base.Subtitle = string.Format(VH.A(8741), Children.Count);
	}
}
