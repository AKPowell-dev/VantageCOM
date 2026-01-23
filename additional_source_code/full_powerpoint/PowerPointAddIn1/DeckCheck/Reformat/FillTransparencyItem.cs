using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class FillTransparencyItem : BaseItem
{
	[CompilerGenerated]
	private float A;

	public float Transparency
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

	public FillTransparencyItem(float sngTrans, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(49943))
	{
		Transparency = sngTrans;
	}

	public void Reformat(float sngNewTrans, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		using List<NavigationItem>.Enumerator enumerator = base.Objects.GetEnumerator();
		while (enumerator.MoveNext())
		{
			IndexedObject indexedObject = enumerator.Current.IndexedObject;
			try
			{
				if (indexedObject.Child is Microsoft.Office.Interop.PowerPoint.Shape)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						((Microsoft.Office.Interop.PowerPoint.Shape)indexedObject.Child).Fill.Transparency = sngNewTrans;
						break;
					}
				}
				else if (indexedObject.Child is Cell)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						((Cell)indexedObject.Child).Shape.Fill.Transparency = sngNewTrans;
						break;
					}
				}
				else if (indexedObject.Shape.HasChart != MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						if (!(indexedObject.Child is Microsoft.Office.Core.Shape))
						{
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							((Microsoft.Office.Core.Shape)indexedObject.Child).Fill.Transparency = sngNewTrans;
							break;
						}
						break;
					}
				}
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
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}
}
