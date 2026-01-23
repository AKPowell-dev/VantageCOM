using System;
using System.Collections.Generic;
using System.Windows;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class TextBoxMarginsItem : MarginsItem
{
	public TextBoxMarginsItem(Margins margins, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex, string strHeader)
		: base(margins, listObjects, template, navItemTemplate, intTotal, intIndex, strHeader)
	{
	}

	public override void Reformat(MarginsOption opt, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		using List<NavigationItem>.Enumerator enumerator = base.Objects.GetEnumerator();
		while (enumerator.MoveNext())
		{
			IndexedObject indexedObject = enumerator.Current.IndexedObject;
			try
			{
				TextFrame2 textFrame = ((Shape)indexedObject.Child).TextFrame2;
				textFrame.MarginTop = opt.Margins.Top;
				textFrame.MarginRight = opt.Margins.Right;
				textFrame.MarginBottom = opt.Margins.Bottom;
				textFrame.MarginLeft = opt.Margins.Left;
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
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}
}
