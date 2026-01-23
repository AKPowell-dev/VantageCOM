using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public abstract class MarginsItem : BaseItem
{
	private readonly string A;

	private readonly string B;

	[CompilerGenerated]
	private Margins A;

	public Margins Margins
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

	public MarginsItem(Margins margins, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex, string strHeader)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, strHeader)
	{
		this.A = AH.A(50090);
		B = AH.A(50099);
		bool isMetric = RegionInfo.CurrentRegion.IsMetric;
		string text;
		if (Operators.CompareString(clsPublish.SystemDecimalSeparator(), AH.A(14417), TextCompare: false) == 0)
		{
			while (true)
			{
				switch (7)
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
			text = this.A;
		}
		else
		{
			text = B;
		}
		Margins = margins;
		if (!isMetric)
		{
			((BaseItem)this).Label = clsPublish.PointsToInches(margins.Top).ToString(text) + AH.A(17773) + clsPublish.PointsToInches(margins.Right).ToString(text) + AH.A(17773) + clsPublish.PointsToInches(margins.Bottom).ToString(text) + AH.A(17773) + clsPublish.PointsToInches(margins.Left).ToString(text) + AH.A(50108);
		}
		else
		{
			((BaseItem)this).Label = clsPublish.PointsToCentimeters(margins.Top).ToString(text) + AH.A(17773) + clsPublish.PointsToCentimeters(margins.Right).ToString(text) + AH.A(17773) + clsPublish.PointsToCentimeters(margins.Bottom).ToString(text) + AH.A(17773) + clsPublish.PointsToCentimeters(margins.Left).ToString(text) + AH.A(50115);
		}
	}

	public abstract void Reformat(MarginsOption opt, ref List<string> listErrors);
}
