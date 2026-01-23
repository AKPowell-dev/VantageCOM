using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class OverlappingText
{
	[CompilerGenerated]
	private Dictionary<Slide, List<ShapeBounds>> A;

	public Dictionary<Slide, List<ShapeBounds>> TextRangeBounds
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

	public OverlappingText(Dictionary<Slide, List<ShapeBounds>> dict)
	{
		TextRangeBounds = dict;
	}

	public void Check(Slide sld, Shape shp)
	{
		List<ShapeBounds> value = null;
		if (!TextRangeBounds.TryGetValue(sld, out value))
		{
			return;
		}
		using (List<ShapeBounds>.Enumerator enumerator = value.GetEnumerator())
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					ShapeBounds current = enumerator.Current;
					bool flag = false;
					bool flag2 = false;
					if ((float)shp.ZOrderPosition > current.Zorder)
					{
						Shape shape = shp;
						float left = shape.Left;
						float num = shape.Left + shape.Width;
						float top = shape.Top;
						float num2 = shape.Top + shape.Height;
						shape = null;
						float left2 = current.Left;
						float right = current.Right;
						float top2 = current.Top;
						float bottom = current.Bottom;
						if (left > left2)
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
							if (left < right)
							{
								flag = true;
							}
						}
						if (left < left2)
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
							if (num > left2)
							{
								flag = true;
							}
						}
						if (top > top2)
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
							if (top < bottom)
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
								flag2 = true;
							}
						}
						if (top < top2)
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
							if (num2 > top2)
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
								flag2 = true;
							}
						}
						if (flag && flag2)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.OverlappingText(sld, shp));
								break;
							}
							break;
						}
					}
					_ = null;
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_016b;
					}
					continue;
					end_IL_016b:
					break;
				}
				break;
			}
		}
		value = null;
	}
}
