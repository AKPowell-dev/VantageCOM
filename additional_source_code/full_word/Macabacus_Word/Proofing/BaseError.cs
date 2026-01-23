using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing;

public class BaseError : BaseError
{
	private int A;

	private Paragraph A;

	private Range A;

	private List<Range> A;

	private Microsoft.Office.Interop.Word.Shape A;

	private InlineShape A;

	private Microsoft.Office.Interop.Word.Shapes A;

	private Table A;

	private Chart A;

	private PlotArea A;

	private Axis A;

	private AxisTitle A;

	private Legend A;

	private ChartTitle A;

	private DataTable A;

	private ChartGroup A;

	private ErrorType A;

	public int PageNumber
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Paragraph Paragraph
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Range Range
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public List<Range> Ranges
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Microsoft.Office.Interop.Word.Shape Shape
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public InlineShape InlineShape
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Microsoft.Office.Interop.Word.Shapes Shapes
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Table Table
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Chart Chart
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public PlotArea PlotArea
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Axis Axis
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public AxisTitle AxisTitle
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public Legend Legend
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public ChartTitle ChartTitle
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public DataTable DataTable
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public ChartGroup ChartGroup
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public ErrorType Type
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public BaseError(ErrorType errType, Severity sev, object obj, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(sev, blnHasFix, blnCanFixMultiple)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		this.A = 0;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		this.A = null;
		Type = errType;
		if (obj is Range)
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
			Range = (Range)obj;
		}
		else if (obj is Microsoft.Office.Interop.Word.Shape)
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
			Shape = (Microsoft.Office.Interop.Word.Shape)obj;
			Range = Shape.Anchor;
			if (Shape.HasChart == MsoTriState.msoTrue)
			{
				Chart = Shape.Chart;
			}
		}
		else if (obj is InlineShape)
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
			InlineShape = (InlineShape)obj;
			Range = InlineShape.Range;
			if (InlineShape.HasChart == MsoTriState.msoTrue)
			{
				Chart = InlineShape.Chart;
			}
		}
		else if (obj is Table)
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
			Table = (Table)obj;
			Range = Table.Range;
		}
		if (Range != null)
		{
			Range duplicate = Range.Duplicate;
			Range range = duplicate;
			object Direction = WdCollapseDirection.wdCollapseStart;
			range.Collapse(ref Direction);
			PageNumber = Conversions.ToInteger(duplicate.get_Information(WdInformation.wdActiveEndPageNumber));
			duplicate = null;
		}
	}

	public string GenerateSnippet(Range rng)
	{
		string text = "";
		Range duplicate = rng.Duplicate;
		int start;
		try
		{
			object Unit = WdUnits.wdWord;
			object Count = 1;
			start = rng.Previous(ref Unit, ref Count).Start;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			start = rng.Start;
			ProjectData.ClearProjectError();
		}
		int end;
		try
		{
			object Count = WdUnits.wdWord;
			object Unit = 1;
			end = rng.Next(ref Count, ref Unit).End;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			end = rng.End;
			ProjectData.ClearProjectError();
		}
		try
		{
			duplicate.SetRange(start, end);
			text = XC.A(37943) + duplicate.Text;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		duplicate = null;
		return text.Replace(XC.A(17685), XC.A(18458)).Replace(XC.A(18455), XC.A(18458)).Replace(XC.A(21985), XC.A(18458));
	}

	public string GenerateSnippet(TextRange2 rng)
	{
		string text = "";
		checked
		{
			try
			{
				TextRange2 textRange = rng;
				int length = textRange.Length;
				int num = textRange.Start - 1;
				int num2 = 1;
				int length2 = textRange.Length;
				_ = null;
				if (length < rng.Text.Length)
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
					MatchCollection matchCollection = new Regex(XC.A(37948)).Matches(rng.Text);
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = matchCollection.GetEnumerator();
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Match match = (Match)enumerator.Current;
								int index = match.Groups[1].Index;
								if (index < num)
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
									num2 = index + 1;
								}
								else
								{
									if (index + match.Groups[1].Length <= num + length)
									{
										continue;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										length2 = index + match.Groups[1].Length - num2 + 1;
										break;
									}
									break;
								}
								continue;
							}
							while (true)
							{
								switch (7)
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
							break;
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
					matchCollection = null;
				}
				text = XC.A(37943) + rng.get_Characters(num2, length2).Text;
				rng = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			return text.Replace(XC.A(17685), XC.A(18458)).Replace(XC.A(18455), XC.A(18458)).Replace(XC.A(21985), XC.A(18458));
		}
	}
}
