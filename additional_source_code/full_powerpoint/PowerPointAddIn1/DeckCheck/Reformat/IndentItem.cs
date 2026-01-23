using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class IndentItem : BaseItem
{
	private readonly string A;

	private readonly string B;

	[CompilerGenerated]
	private Indent A;

	public Indent Indent
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

	public IndentItem(Indent indent, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(50055))
	{
		this.A = AH.A(50090);
		B = AH.A(50099);
		bool isMetric = RegionInfo.CurrentRegion.IsMetric;
		string text;
		if (Operators.CompareString(clsPublish.SystemDecimalSeparator(), AH.A(14417), TextCompare: false) == 0)
		{
			while (true)
			{
				switch (5)
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
		Indent = indent;
		if (!isMetric)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((BaseItem)this).Label = clsPublish.PointsToInches(indent.LeftIndent).ToString(text) + AH.A(17773) + clsPublish.PointsToInches(indent.FirstLineIndent).ToString(text) + AH.A(50108);
					return;
				}
			}
		}
		((BaseItem)this).Label = clsPublish.PointsToCentimeters(indent.LeftIndent).ToString(text) + AH.A(17773) + clsPublish.PointsToCentimeters(indent.FirstLineIndent).ToString(text) + AH.A(50115);
	}

	public void Reformat(IndentOption opt, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		foreach (NavigationItem @object in base.Objects)
		{
			IndexedObject indexedObject = @object.IndexedObject;
			try
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)indexedObject.Child).ParagraphFormat;
				paragraphFormat.LeftIndent = opt.Indent.LeftIndent;
				paragraphFormat.FirstLineIndent = opt.Indent.FirstLineIndent;
				_ = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				listErrors.Add(ex2.Message);
				ProjectData.ClearProjectError();
			}
			indexedObject = null;
		}
	}
}
