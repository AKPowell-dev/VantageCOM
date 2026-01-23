using System;
using System.Collections.Generic;
using System.Windows.Forms;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.TextOps;

public sealed class Swap
{
	public static void SwapText()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
		Selection selection;
		try
		{
			selection = application.ActiveWindow.Selection;
			if (selection.Type == PpSelectionType.ppSelectionShapes)
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
					shapeRange = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection);
					if (shapeRange.Count == 2)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							A(application, shapeRange[1], shapeRange[2]);
							break;
						}
						break;
					}
					if (shapeRange.Count == 1)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						if (shapeRange[1].HasTable == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(application, shapeRange[1]);
								break;
							}
							break;
						}
					}
					A();
					break;
				}
			}
			else
			{
				A();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A();
			ProjectData.ClearProjectError();
		}
		shapeRange = null;
		selection = null;
		application = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Application A, Microsoft.Office.Interop.PowerPoint.Shape B, Microsoft.Office.Interop.PowerPoint.Shape C)
	{
		if (B.HasTextFrame == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (6)
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
			if (C.HasTextFrame == MsoTriState.msoTrue)
			{
				A.StartNewUndoEntry();
				Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
				TextRange2 textRange;
				TextRange2 textRange2;
				try
				{
					shape = B.Duplicate()[1];
					textRange = shape.TextFrame2.TextRange;
					textRange2 = B.TextFrame2.TextRange;
					TextRange2 textRange3 = C.TextFrame2.TextRange;
					textRange3.Copy();
					textRange2.PasteSpecial(MsoClipboardFormat.msoClipboardFormatPlainText);
					textRange.Copy();
					textRange3.PasteSpecial(MsoClipboardFormat.msoClipboardFormatPlainText);
					B.Select();
					C.Select(MsoTriState.msoFalse);
					Base.LogActivity(AH.A(155909));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Clipboard.Clear();
					try
					{
						shape.Delete();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				shape = null;
				textRange = null;
				textRange2 = null;
				return;
			}
		}
		Forms.WarningMessage(AH.A(155928));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Application A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		List<Cell> list = new List<Cell>();
		List<Tuple<int, int>> list2 = new List<Tuple<int, int>>();
		Table table = B.Table;
		int count = table.Rows.Count;
		checked
		{
			for (int i = 1; i <= count; i++)
			{
				int count2 = table.Columns.Count;
				for (int j = 1; j <= count2; j++)
				{
					if (!table.Cell(i, j).Selected)
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					list.Add(table.Cell(i, j));
					list2.Add(new Tuple<int, int>(i, j));
				}
			}
			table = null;
			if (list.Count == 2)
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
				A.StartNewUndoEntry();
				Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
				TextRange2 textRange;
				TextRange2 textRange2;
				try
				{
					shape = B.Duplicate()[1];
					textRange = shape.Table.Cell(list2[0].Item1, list2[0].Item2).Shape.TextFrame2.TextRange;
					textRange2 = list[0].Shape.TextFrame2.TextRange;
					TextRange2 textRange3 = list[1].Shape.TextFrame2.TextRange;
					textRange3.Copy();
					textRange2.PasteSpecial(MsoClipboardFormat.msoClipboardFormatPlainText);
					textRange.Copy();
					textRange3.PasteSpecial(MsoClipboardFormat.msoClipboardFormatPlainText);
					Base.LogActivity(AH.A(155909));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Clipboard.Clear();
					try
					{
						shape.Delete();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				shape = null;
				textRange = null;
				textRange2 = null;
			}
			else
			{
				Swap.A();
			}
			list = null;
			list2 = null;
		}
	}

	private static void A()
	{
		Forms.WarningMessage(AH.A(156029));
	}
}
