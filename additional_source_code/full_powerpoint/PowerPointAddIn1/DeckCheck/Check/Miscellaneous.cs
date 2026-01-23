using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Miscellaneous
{
	public static void CheckSlideMasterCount(Microsoft.Office.Interop.PowerPoint.Presentation pres, SlideRange sldRange)
	{
		if (sldRange.Count != pres.Slides.Count)
		{
			return;
		}
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
			Designs designs = pres.Designs;
			int count = designs.Count;
			if (count > 1)
			{
				List<string> list = new List<string>();
				int num = count;
				for (int i = 1; i <= num; i = checked(i + 1))
				{
					list.Add(designs[i].Name);
				}
				Main.Analysis.Errors.Add(new MultipleSlideMasters(string.Join(AH.A(14258), list.ToArray())));
				list = null;
			}
			designs = null;
			return;
		}
	}

	public static void CheckSlideCount(Microsoft.Office.Interop.PowerPoint.Presentation pres, int intMax)
	{
		int count = pres.Slides.Count;
		if (count > intMax)
		{
			Main.Analysis.Errors.Add(new SlideCount(count, intMax));
		}
	}

	public static void CheckSlideWordCount(Slide sld, int intMax)
	{
		int B = 0;
		foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in sld.Shapes)
		{
			A(shape, ref B);
		}
		if (B <= intMax)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new SlideWordCount(sld, B, intMax));
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref int B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		checked
		{
			if (shape.Type != MsoShapeType.msoGroup)
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
				if (shape.HasTextFrame == MsoTriState.msoTrue)
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
					if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
					{
						B += shape.TextFrame2.TextRange.get_Words(-1, -1).Count;
					}
				}
				else if (shape.HasTable == MsoTriState.msoTrue)
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
					Table table = A.Table;
					int count = table.Rows.Count;
					int count2 = table.Columns.Count;
					int num = count;
					for (int i = 1; i <= num; i++)
					{
						int num2 = count2;
						for (int j = 1; j <= num2; j++)
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = table.Cell(i, j).Shape;
							if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
								if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
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
									B += shape2.TextFrame2.TextRange.get_Words(-1, -1).Count;
								}
							}
							shape2 = null;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0131;
							}
							continue;
							end_IL_0131:
							break;
						}
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
					table = null;
				}
				else if (shape.HasSmartArt == MsoTriState.msoTrue)
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
					{
						IEnumerator enumerator = shape.SmartArt.AllNodes.GetEnumerator();
						try
						{
							IEnumerator enumerator2 = default(IEnumerator);
							while (enumerator.MoveNext())
							{
								SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
								try
								{
									enumerator2 = smartArtNode.Shapes.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										Microsoft.Office.Core.Shape shape3 = (Microsoft.Office.Core.Shape)enumerator2.Current;
										if (shape3.TextFrame2.HasText == MsoTriState.msoTrue)
										{
											B += shape3.TextFrame2.TextRange.get_Words(-1, -1).Count;
										}
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_01fe;
										}
										continue;
										end_IL_01fe:
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
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0238;
								}
								continue;
								end_IL_0238:
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
			}
			else
			{
				int count3 = shape.GroupItems.Count;
				for (int k = 1; k <= count3; k++)
				{
					Miscellaneous.A(shape.GroupItems[k], ref B);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			shape = null;
		}
	}

	public static void CheckMacabacusLink(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		if (!PowerPointAddIn1.Links.Shapes.IsLinked(shp))
		{
			return;
		}
		string source = PowerPointAddIn1.Links.Shapes.LinkDetails(shp).Source;
		if (clsFile.NewerVersions(source).Count > 0)
		{
			Main.Analysis.Errors.Add(new LinkNewerVersionAvailable(sld, shp, Path.GetFileName(source)));
		}
		if (clsFile.IsPathUrl(source))
		{
			return;
		}
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
			if (File.Exists(source))
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				Main.Analysis.Errors.Add(new LinkBroken(sld, shp, Path.GetFileName(source)));
				return;
			}
		}
	}

	public static void CheckBulletWordCount(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, int intMax)
	{
		int count = para.get_Words(-1, -1).Count;
		if (count <= intMax)
		{
			return;
		}
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
			Main.Analysis.Errors.Add(new BulletWordCount(sld, shp, para, count, intMax));
			return;
		}
	}
}
