using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Publishing.Share;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class WorksheetItem : SheetItem
{
	[CompilerGenerated]
	private new Worksheet m_A;

	internal Worksheet Worksheet
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public WorksheetItem(WorkbookItem wbi, Worksheet ws, Microsoft.Office.Interop.Excel.Workbook wb, int intResults)
		: base(ws, wb, wbi, Props.Icons.GeoWorksheet)
	{
		Worksheet = ws;
		base.ResultsCount = intResults;
	}

	internal void A(Range A)
	{
		base.Children.Add(new TextItem(this, A));
	}

	internal void B(Range A)
	{
		base.Children.Add(new ValueItem(this, A));
	}

	internal void C(Range A)
	{
		base.Children.Add(new DateItem(this, A));
	}

	internal void D(Range A)
	{
		base.Children.Add(new FormatItem(this, A));
	}

	internal void E(Range A)
	{
		base.Children.Add(new FormulaItem(this, A));
	}

	internal void F(Range A)
	{
		base.Children.Add(new RangeItem(this, A));
	}

	internal new void A()
	{
		base.Children.Add(new UsedRangeItem(this));
	}

	internal void G(Range A)
	{
		base.Children.Add(new PrintAreaItem(this, A));
	}

	internal void A(AutoFilter A)
	{
		base.Children.Add(new AutoFilterItem(this, A));
	}

	internal void A(ListObject A)
	{
		base.Children.Add(new TableItem(this, A));
	}

	internal void A(QueryTable A)
	{
		base.Children.Add(new QueryTableItem(this, A));
	}

	internal void A(PivotTable A)
	{
		base.Children.Add(new PivotTableItem(this, A));
	}

	internal void H(Range A)
	{
		base.Children.Add(new SpillRangeItem(this, A));
	}

	internal void I(Range A)
	{
		base.Children.Add(new ArrayFormulaItem(this, A));
	}

	internal void J(Range A)
	{
		base.Children.Add(new NumericInputItem(this, A));
	}

	internal void A(string A)
	{
		base.Children.Add(new OtherInputsItem(this, A));
	}

	internal void A(Range A, XlErrorChecks B)
	{
		base.Children.Add(new ErrorItem(this, A, B));
	}

	internal void K(Range A)
	{
		base.Children.Add(new ErrorItem(this, A, XlErrorChecks.xlEvaluateToError));
	}

	internal void B(string A)
	{
	}

	internal void A(ChartObject A)
	{
		base.Children.Add(new ChartItem(this, A));
	}

	internal void A(Shape A)
	{
		base.Children.Add(new ShapeItem(this, A));
	}

	internal void L(Range A)
	{
		base.Children.Add(new NoteItem(this, A));
	}

	internal void M(Range A)
	{
		base.Children.Add(new CommentItem(this, A));
	}

	internal void N(Range A)
	{
		base.Children.Add(new ValidationItem(this, A));
	}

	internal void O(Range A)
	{
		base.Children.Add(new ConditionalFormatItem(this, A));
	}

	internal void A(SparklineGroup A)
	{
		base.Children.Add(new SparklineItem(this, A));
	}

	internal void A(Watch A)
	{
		base.Children.Add(new WatchItem(this, A));
	}

	internal void A(Hyperlink A)
	{
		base.Children.Add(new HyperlinkItem(this, A));
	}

	internal void A(Name A, Range B, bool C, string D = null)
	{
		base.Children.Add(new NameItem(this, A, B, C, D));
	}

	internal void C(string A)
	{
		base.Children.Add(new OtherNamesItem(this, A));
	}

	internal void P(Range A)
	{
		base.Children.Add(new MacabacusLinkItem(this, A));
	}

	internal void Q(Range A)
	{
		base.Children.Add(new MergedCellsItem(this, A));
	}

	internal new void B()
	{
		S();
		base.C();
	}

	internal new void C()
	{
		base.Workbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	internal new void D()
	{
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				Worksheet.Cells.Clear();
				Base.K(Worksheet);
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					try
					{
						if (base.Children[i] is ChartItem)
						{
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
								((ChartItem)base.Children[i]).ChartObject.Delete();
								break;
							}
						}
						else if (base.Children[i] is SparklineItem)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								((SparklineItem)base.Children[i]).SparklineGroup.Delete();
								break;
							}
						}
						else if (base.Children[i] is ShapeItem)
						{
							((ShapeItem)base.Children[i]).Shape.Delete();
						}
						else
						{
							if (!(base.Children[i] is WatchItem))
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
								((WatchItem)base.Children[i]).Watch.Delete();
								break;
							}
							continue;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
			B();
		}
	}

	internal new void E()
	{
		Worksheet.Cells.ClearContents();
		B();
	}

	internal new void F()
	{
		Worksheet.Cells.ClearFormats();
		checked
		{
			for (int i = base.Children.Count - 1; i >= 0; i += -1)
			{
				if (!(base.Children[i] is ConditionalFormatItem))
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A(base.Children[i]);
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

	internal new void G()
	{
		Worksheet.Cells.ClearComments();
		checked
		{
			for (int i = base.Children.Count - 1; i >= 0; i += -1)
			{
				if (base.Children[i] is CommentItem || base.Children[i] is NoteItem)
				{
					A(base.Children[i]);
				}
			}
			while (true)
			{
				switch (3)
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

	internal new void H()
	{
		Worksheet.Cells.ClearHyperlinks();
		checked
		{
			for (int i = base.Children.Count - 1; i >= 0; i += -1)
			{
				if (base.Children[i] is HyperlinkItem)
				{
					A(base.Children[i]);
				}
			}
		}
	}

	internal void I()
	{
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				Base.K(Worksheet);
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					if (!(base.Children[i] is PrintAreaItem))
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
					A(base.Children[i]);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_007c;
					}
					continue;
					end_IL_007c:
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
		}
	}

	internal void J()
	{
		bool flag = false;
		bool flag2 = false;
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					if (!(base.Children[i] is ShapeItem))
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Shape shape = ((ShapeItem)base.Children[i]).Shape;
					if (shape.Type != MsoShapeType.msoFormControl)
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
						try
						{
							shape.Delete();
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						A(base.Children[i]);
					}
					else
					{
						if (!flag2)
						{
							flag = MessageBox.Show(VH.A(124189), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes;
							flag2 = true;
						}
						if (flag2)
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
							if (flag)
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
									shape.Delete();
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
								}
								A(base.Children[i]);
							}
						}
					}
					shape = null;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
		}
	}

	internal void K()
	{
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					if (!(base.Children[i] is ChartItem))
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						((ChartItem)base.Children[i]).ChartObject.Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					A(base.Children[i]);
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
		}
	}

	internal void L()
	{
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					if (!(base.Children[i] is SparklineItem))
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						((SparklineItem)base.Children[i]).SparklineGroup.Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					A(base.Children[i]);
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
		}
	}

	internal void M()
	{
		checked
		{
			for (int i = base.Children.Count - 1; i >= 0; i += -1)
			{
				if (!(base.Children[i] is WatchItem))
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					((WatchItem)base.Children[i]).Watch.Delete();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				A(base.Children[i]);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}

	internal void N()
	{
		Worksheet.Application.ScreenUpdating = false;
		checked
		{
			try
			{
				for (int i = base.Children.Count - 1; i >= 0; i += -1)
				{
					if (!(base.Children[i] is ConditionalFormatItem))
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						((ConditionalFormatItem)base.Children[i]).Range.FormatConditions.Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					A(base.Children[i]);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00ad;
					}
					continue;
					end_IL_00ad:
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			Worksheet.Application.ScreenUpdating = true;
		}
	}

	internal void O()
	{
		throw new NotImplementedException();
	}

	internal void P()
	{
		throw new NotImplementedException();
	}

	internal void Q()
	{
		Worksheet.Application.ActiveWindow.View = XlWindowView.xlNormalView;
	}

	internal void R()
	{
		Worksheet.Application.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
	}

	private void S()
	{
		((BaseItem)this).Label = Worksheet.Name;
		if (!Worksheet.ProtectContents)
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
			((BaseItem)this).Icon = Props.Icons.GeoWorksheet;
		}
		else
		{
			((BaseItem)this).Icon = Props.Icons.GeoLock;
		}
		((BaseItem)this).Icon.Freeze();
		A(Worksheet.Visible);
	}
}
