using System;
using System.Collections;
using System.Collections.Generic;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.TextOps;

public sealed class Bullets
{
	private struct YF
	{
		public MsoNumberedBulletStyle A;

		public MsoBulletType A;

		public Font2 A;

		public float A;

		public float B;

		public float C;

		public int A;

		public int B;

		public MsoTriState A;

		public MsoTriState B;
	}

	public static void Fix()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		Application application = NG.A.Application;
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Dictionary<int, YF> dictionary = null;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = default(Microsoft.Office.Interop.PowerPoint.Presentation);
		Selection selection = default(Selection);
		try
		{
			activePresentation = application.ActivePresentation;
			selection = application.ActiveWindow.Selection;
			slide = selection.SlideRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(AH.A(153908));
			ProjectData.ClearProjectError();
		}
		if (slide != null)
		{
			while (true)
			{
				switch (3)
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
			try
			{
				shape = PowerPointAddIn1.MasterShapes.TextBox.Shape(activePresentation);
				if (shape != null)
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
					dictionary = A(shape);
				}
				if (dictionary == null)
				{
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = slide.CustomLayout.Shapes.Placeholders.GetEnumerator();
						while (true)
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape3;
							if (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								shape3 = shape2;
								PpPlaceholderType type = shape3.PlaceholderFormat.Type;
								if (type != PpPlaceholderType.ppPlaceholderMixed)
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
									if (type != PpPlaceholderType.ppPlaceholderBody)
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
										if (type != PpPlaceholderType.ppPlaceholderObject)
										{
											goto IL_017e;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											break;
										}
									}
								}
								if (shape3.HasTextFrame == MsoTriState.msoTrue)
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
									if (shape3.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != MsoBulletType.msoBulletNone)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											dictionary = A(shape2);
											break;
										}
										break;
									}
								}
								goto IL_017e;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_018f;
								}
								continue;
								end_IL_018f:
								break;
							}
							break;
							IL_017e:
							shape3 = null;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				if (dictionary == null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						shape = Helpers.GetBodyPlaceholder(activePresentation);
						if (shape == null)
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
							dictionary = A(shape);
							break;
						}
						break;
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(ex4.Message);
				ProjectData.ClearProjectError();
			}
			finally
			{
				slide = null;
				shape = null;
			}
		}
		if (dictionary != null)
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
			try
			{
				application.StartNewUndoEntry();
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, dictionary);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0255;
						}
						continue;
						end_IL_0255:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				Base.LogActivity(AH.A(153949));
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				Forms.ErrorMessage(ex6.Message);
				ProjectData.ClearProjectError();
			}
			finally
			{
				dictionary = null;
			}
		}
		else
		{
			Forms.WarningMessage(AH.A(153972));
		}
		application = null;
		activePresentation = null;
		selection = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Dictionary<int, YF> B)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Bullets.A(A.TextFrame2.TextRange, B);
					return;
				}
			}
		}
		checked
		{
			if (A.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
					{
						Table table = A.Table;
						int count = table.Rows.Count;
						int count2 = table.Columns.Count;
						int num = count;
						for (int i = 1; i <= num; i++)
						{
							int num2 = count2;
							for (int j = 1; j <= num2; j++)
							{
								Cell cell = table.Cell(i, j);
								if (cell.Selected)
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
									if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
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
										Bullets.A(cell.Shape.TextFrame2.TextRange, B);
									}
								}
								cell = null;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_00e7;
								}
								continue;
								end_IL_00e7:
								break;
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								table = null;
								return;
							}
						}
					}
					}
				}
			}
			if (A.HasSmartArt != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.SmartArt.AllNodes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Bullets.A(((SmartArtNode)enumerator.Current).TextFrame2.TextRange, B);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}

	private static Dictionary<int, YF> A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Dictionary<int, YF> dictionary = new Dictionary<int, YF>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 textRange = (TextRange2)enumerator.Current;
				ParagraphFormat2 paragraphFormat = textRange.ParagraphFormat;
				if (paragraphFormat.Bullet.Type != MsoBulletType.msoBulletNone)
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
					if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
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
						YF value = default(YF);
						BulletFormat2 bullet = paragraphFormat.Bullet;
						value.A = bullet.Style;
						value.A = bullet.Type;
						value.A = bullet.Font;
						value.A = bullet.RelativeSize;
						value.A = bullet.StartValue;
						value.B = bullet.Character;
						value.A = bullet.UseTextColor;
						value.B = bullet.UseTextFont;
						bullet = null;
						value.B = textRange.ParagraphFormat.LeftIndent;
						value.C = textRange.ParagraphFormat.FirstLineIndent;
						dictionary.Add(paragraphFormat.IndentLevel, value);
					}
				}
				paragraphFormat = null;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				return dictionary;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void A(TextRange2 A, Dictionary<int, YF> B)
	{
		IEnumerator enumerator = A.get_Paragraphs(-1, -1).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
				if (B.TryGetValue(paragraphFormat.IndentLevel, out var value))
				{
					paragraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
					Font2 a = value.A;
					BulletFormat2 bullet = paragraphFormat.Bullet;
					bullet.RelativeSize = value.A;
					bullet.Type = value.A;
					switch (value.A)
					{
					case MsoBulletType.msoBulletUnnumbered:
						bullet.Character = value.B;
						break;
					case MsoBulletType.msoBulletNumbered:
						bullet.StartValue = value.A;
						bullet.Style = value.A;
						break;
					case MsoBulletType.msoBulletMixed:
						bullet.Style = MsoNumberedBulletStyle.msoBulletStyleMixed;
						break;
					}
					Font2 font = bullet.Font;
					font.Name = a.Name;
					font.Bold = a.Bold;
					font.Fill.ForeColor = a.Fill.ForeColor;
					font.Fill.BackColor = a.Fill.BackColor;
					_ = null;
					bullet.UseTextColor = value.A;
					if (value.A == MsoTriState.msoFalse)
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
						bullet.Font.Fill.ForeColor.RGB = value.A.Fill.ForeColor.RGB;
					}
					bullet.UseTextFont = value.B;
					if (value.B == MsoTriState.msoFalse)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							break;
						}
						bullet.Font.Name = value.A.Name;
					}
					bullet = null;
					paragraphFormat.LeftIndent = value.B;
					paragraphFormat.FirstLineIndent = value.C;
					a = null;
				}
				paragraphFormat = null;
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
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static void ApplyNumberedBullets(MsoNumberedBulletStyle style)
	{
		Application application = NG.A.Application;
		try
		{
			Selection selection = application.ActiveWindow.Selection;
			PpSelectionType type = selection.Type;
			if (type != PpSelectionType.ppSelectionShapes)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (type == PpSelectionType.ppSelectionText)
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
					application.StartNewUndoEntry();
					BulletFormat2 bullet = selection.TextRange2.ParagraphFormat.Bullet;
					bullet.Type = MsoBulletType.msoBulletNumbered;
					bullet.Style = style;
					_ = null;
				}
				else
				{
					Forms.WarningMessage(AH.A(154168));
				}
			}
			else
			{
				application.StartNewUndoEntry();
				if (selection.HasChildShapeRange)
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
					{
						IEnumerator enumerator = selection.ChildShapeRange.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, style);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_00c1;
								}
								continue;
								end_IL_00c1:
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
				}
				else
				{
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = selection.ShapeRange.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, style);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0117;
							}
							continue;
							end_IL_0117:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			Base.LogActivity(AH.A(154237));
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
			Selection selection = null;
		}
		application = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, MsoNumberedBulletStyle B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (A.HasTextFrame == MsoTriState.msoTrue)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								A.TextFrame2.TextRange.ParagraphFormat.Bullet.Style = B;
								return;
							}
						}
					}
					return;
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Bullets.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
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
}
