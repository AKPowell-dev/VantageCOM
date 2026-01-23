using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TextOps;

public sealed class Tables
{
	public static void UngroupTable()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		Application application = NG.A.Application;
		checked
		{
			Selection selection;
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			Table table;
			List<string> list;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					IEnumerator enumerator = default(IEnumerator);
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (selection.ShapeRange.Count == 1)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								shape = selection.ShapeRange[1];
								if (shape.Type != MsoShapeType.msoTable)
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
									if (shape.HasTable != MsoTriState.msoTrue)
									{
										B();
										break;
									}
								}
								list = new List<string>();
								table = shape.Table;
								application.StartNewUndoEntry();
								int count = table.Rows.Count;
								for (int i = 1; i <= count; i++)
								{
									int count2 = table.Columns.Count;
									for (int j = 1; j <= count2; j++)
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2 = table.Cell(i, j).Shape;
										Microsoft.Office.Interop.PowerPoint.Shape shape3 = selection.SlideRange[1].Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, shape2.Left, shape2.Top, shape2.Width, shape2.Height);
										list.Add(shape3.Name);
										shape3.TextFrame2.TextRange.Text = shape2.TextFrame2.TextRange.Text;
										{
											enumerator = shape2.TextFrame2.TextRange.get_Runs(-1, -1).GetEnumerator();
											try
											{
												while (enumerator.MoveNext())
												{
													TextRange2 textRange = (TextRange2)enumerator.Current;
													Font2 font = textRange.Font;
													Font2 font2 = shape3.TextFrame2.TextRange.get_Characters(textRange.Start, textRange.Length).Font;
													font2.Bold = font.Bold;
													font2.Fill.ForeColor.RGB = font.Fill.ForeColor.RGB;
													font2.Italic = font.Italic;
													font2.UnderlineStyle = font.UnderlineStyle;
													font2.Name = font.Name;
													font2.Size = font.Size;
													font2.Spacing = font.Spacing;
													_ = null;
												}
												while (true)
												{
													switch (4)
													{
													case 0:
														break;
													default:
														goto end_IL_027c;
													}
													continue;
													end_IL_027c:
													break;
												}
											}
											finally
											{
												IDisposable disposable = enumerator as IDisposable;
												if (disposable != null)
												{
													disposable.Dispose();
												}
											}
										}
										Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape2.Fill;
										Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shape3.Fill;
										fill2.Visible = fill.Visible;
										if (fill2.Visible == MsoTriState.msoTrue)
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
											fill2.ForeColor.RGB = fill.ForeColor.RGB;
											fill2.Transparency = fill.Transparency;
										}
										fill2 = null;
										Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape2.TextFrame2;
										shape3.TextFrame2.MarginTop = textFrame.MarginTop;
										shape3.TextFrame2.MarginRight = textFrame.MarginRight;
										shape3.TextFrame2.MarginBottom = textFrame.MarginBottom;
										shape3.TextFrame2.MarginLeft = textFrame.MarginLeft;
										shape3.TextFrame2.TextRange.ParagraphFormat.Alignment = textFrame.TextRange.ParagraphFormat.Alignment;
										textFrame = null;
										shape2 = null;
										shape3 = null;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_03ab;
										}
										continue;
										end_IL_03ab:
										break;
									}
								}
								Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = application.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray());
								if (shape.Type == MsoShapeType.msoPlaceholder)
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
									shape.Delete();
									shapeRange.ZOrder(MsoZOrderCmd.msoBringToFront);
								}
								else
								{
									shape.Delete();
								}
								shapeRange.Select();
								shapeRange = null;
								Base.LogActivity(AH.A(156194));
								break;
							}
						}
						else
						{
							A();
						}
						break;
					}
				}
				else
				{
					B();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			shape = null;
			table = null;
			list = null;
		}
	}

	public static void TextboxToTable()
	{
		Application application = NG.A.Application;
		checked
		{
			Selection selection;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (selection.ShapeRange.Count == 1)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Microsoft.Office.Interop.PowerPoint.Shape shape = selection.ShapeRange[1];
								if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
									Dictionary<int, List<string>> dictionary = new Dictionary<int, List<string>>();
									int count = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).Count;
									int num = 1;
									int count2 = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).Count;
									for (int i = 1; i <= count2; i++)
									{
										string[] array = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).Item(i).Text.Split('\t');
										num = Math.Max(num, array.Length);
										dictionary.Add(i, array.ToList());
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										break;
									}
									selection.SlideRange[1].Shapes.AddTable(count, num, shape.Left, shape.Top);
									int num2 = count;
									for (int j = 1; j <= num2; j++)
									{
										int num3 = num;
										for (int k = 1; k <= num3; k++)
										{
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_01b8;
											}
											continue;
											end_IL_01b8:
											break;
										}
									}
								}
								shape = null;
								break;
							}
						}
						else
						{
							Forms.WarningMessage(AH.A(156221));
						}
						break;
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(156221));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
		}
	}

	private static void A()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected I4, but got Unknown
		string text = AH.A(156306);
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		switch ((int)applicationLanguage)
		{
		case 2:
			text = AH.A(156403);
			break;
		case 3:
			text = AH.A(156526);
			break;
		}
		Forms.WarningMessage(text);
	}

	private static void B()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected I4, but got Unknown
		string text = AH.A(156637);
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		switch ((int)applicationLanguage)
		{
		case 2:
			text = AH.A(156674);
			break;
		case 3:
			text = AH.A(156727);
			break;
		}
		Forms.WarningMessage(text);
	}

	public static bool ContainsMergedOrSplitCell(Table tbl)
	{
		int count = tbl.Rows.Count;
		int count2 = tbl.Columns.Count;
		checked
		{
			try
			{
				int num = count;
				for (int i = 1; i <= num; i++)
				{
					float height = tbl.Rows[i].Height;
					int num2 = count2;
					int num3 = 1;
					while (num3 <= num2)
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = tbl.Cell(i, num3).Shape;
						if (shape.Height == height)
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
							if (shape.Width == tbl.Columns[num3].Width)
							{
								num3++;
								continue;
							}
						}
						return true;
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00c6;
					}
					continue;
					end_IL_00c6:
					break;
				}
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			}
			return false;
		}
	}

	public static bool IsCellMergedOrSplit(Cell cell, float rowHeight, float colWidth)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = cell.Shape;
		if (shape.Height == rowHeight)
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
					return shape.Width != colWidth;
				}
			}
		}
		return true;
	}
}
