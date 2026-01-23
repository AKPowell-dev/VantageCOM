using System;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.UndoRedo;

public sealed class CellProp
{
	[CompilerGenerated]
	private KG m_A;

	[CompilerGenerated]
	private LG m_A;

	[CompilerGenerated]
	private JG m_A;

	[CompilerGenerated]
	private JG m_B;

	[CompilerGenerated]
	private JG C;

	[CompilerGenerated]
	private JG D;

	[CompilerGenerated]
	private JG E;

	[CompilerGenerated]
	private JG F;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private string m_B;

	[CompilerGenerated]
	private string C;

	[CompilerGenerated]
	private XlVAlign m_A;

	[CompilerGenerated]
	private XlHAlign m_A;

	[CompilerGenerated]
	private XlOrientation m_A;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private float m_B;

	private KG Font
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

	private LG Interior
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

	private JG BorderTop
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

	private JG BorderBottom
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private JG BorderLeft
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	private JG BorderRight
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	private JG BorderDiagUp
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	private JG BorderDiagDown
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	private string Formula
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

	private string FormulaArray
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private string NumberFormat
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	private XlVAlign VerticalAlignment
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

	private XlHAlign HorizontalAlignment
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

	private XlOrientation Orientation
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

	private int IndentLevel
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

	private bool WrapText
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

	private float RowHeight
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

	private float ColumnWidth
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal CellProp()
	{
		Font = new KG();
		Interior = new LG();
		BorderTop = new JG();
		BorderBottom = new JG();
		BorderLeft = new JG();
		BorderRight = new JG();
		BorderDiagUp = new JG();
		BorderDiagDown = new JG();
		Worksheet worksheet = (Worksheet)MH.A.Application.ActiveSheet;
		RowHeight = (float)worksheet.StandardHeight;
		ColumnWidth = (float)worksheet.StandardWidth;
		worksheet = null;
	}

	internal void A(ref Range A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				LG interior;
				Interior A2;
				JG borderDiagDown;
				JG borderTop;
				JG borderBottom;
				JG borderLeft;
				JG borderRight;
				KG font;
				Font A4;
				JG borderDiagUp;
				Borders A3;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 808:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000c;
						case 4:
							goto IL_0039;
						case 5:
							goto IL_005c;
						case 6:
							goto IL_0073;
						case 7:
							goto IL_0090;
						case 8:
							goto IL_00a7;
						case 9:
							goto IL_00bc;
						case 10:
							goto IL_00d2;
						case 11:
							goto IL_00ea;
						case 12:
							goto IL_0103;
						case 13:
							goto IL_0117;
						case 14:
							goto IL_012d;
						case 15:
							goto IL_0141;
						case 16:
							goto IL_0155;
						case 17:
							goto IL_016b;
						case 18:
							goto IL_0186;
						case 19:
							goto IL_01a2;
						case 20:
							goto IL_01bd;
						case 21:
							goto IL_01d9;
						case 22:
							goto IL_01f7;
						case 23:
							goto IL_0214;
						case 24:
							goto IL_0232;
						case 25:
							goto IL_024f;
						case 26:
							goto IL_026b;
						case 27:
							goto IL_0284;
						case 28:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 29:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_01bd:
					num2 = 20;
					interior = Interior;
					A2 = range.Interior;
					interior.A(ref A2);
					goto IL_01d9;
					IL_0007:
					num2 = 2;
					range = A;
					goto IL_000c;
					IL_000c:
					num2 = 3;
					Formula = NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null).ToString();
					goto IL_0039;
					IL_0039:
					num2 = 4;
					if (Information.Err().Number != 0)
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
						goto IL_005c;
					}
					goto IL_0073;
					IL_01d9:
					num2 = 21;
					if (!KH.A.UndoBorders)
					{
						break;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_01f7;
					IL_0284:
					num2 = 27;
					borderDiagDown = BorderDiagDown;
					A3 = range.Borders;
					borderDiagDown.A(ref A3, XlBordersIndex.xlDiagonalDown);
					break;
					IL_01f7:
					num2 = 22;
					borderTop = BorderTop;
					A3 = range.Borders;
					borderTop.A(ref A3, XlBordersIndex.xlEdgeTop);
					goto IL_0214;
					IL_005c:
					num2 = 5;
					Formula = range.Formula.ToString();
					goto IL_0073;
					IL_0073:
					num2 = 6;
					if (Conversions.ToBoolean(range.HasArray))
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
						goto IL_0090;
					}
					goto IL_00a7;
					IL_0214:
					num2 = 23;
					borderBottom = BorderBottom;
					A3 = range.Borders;
					borderBottom.A(ref A3, XlBordersIndex.xlEdgeBottom);
					goto IL_0232;
					IL_0090:
					num2 = 7;
					FormulaArray = Conversions.ToString(range.FormulaArray);
					goto IL_00a7;
					IL_00a7:
					num2 = 8;
					NumberFormat = range.NumberFormat.ToString();
					goto IL_00bc;
					IL_00bc:
					num2 = 9;
					RowHeight = Conversions.ToSingle(range.RowHeight);
					goto IL_00d2;
					IL_00d2:
					num2 = 10;
					ColumnWidth = Conversions.ToSingle(range.ColumnWidth);
					goto IL_00ea;
					IL_00ea:
					num2 = 11;
					if (KH.A.UndoAlignment)
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
						goto IL_0103;
					}
					goto IL_016b;
					IL_0232:
					num2 = 24;
					borderLeft = BorderLeft;
					A3 = range.Borders;
					borderLeft.A(ref A3, XlBordersIndex.xlEdgeLeft);
					goto IL_024f;
					IL_0103:
					num2 = 12;
					IndentLevel = Conversions.ToInteger(range.IndentLevel);
					goto IL_0117;
					IL_0117:
					num2 = 13;
					HorizontalAlignment = (XlHAlign)Conversions.ToInteger(range.HorizontalAlignment);
					goto IL_012d;
					IL_012d:
					num2 = 14;
					VerticalAlignment = (XlVAlign)Conversions.ToInteger(range.VerticalAlignment);
					goto IL_0141;
					IL_0141:
					num2 = 15;
					Orientation = (XlOrientation)Conversions.ToInteger(range.Orientation);
					goto IL_0155;
					IL_0155:
					num2 = 16;
					WrapText = Conversions.ToBoolean(range.WrapText);
					goto IL_016b;
					IL_016b:
					num2 = 17;
					if (KH.A.UndoFont)
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
						goto IL_0186;
					}
					goto IL_01a2;
					IL_024f:
					num2 = 25;
					borderRight = BorderRight;
					A3 = range.Borders;
					borderRight.A(ref A3, XlBordersIndex.xlEdgeRight);
					goto IL_026b;
					IL_0186:
					num2 = 18;
					font = Font;
					A4 = range.Font;
					font.A(ref A4);
					goto IL_01a2;
					IL_01a2:
					num2 = 19;
					if (KH.A.UndoFill)
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
						goto IL_01bd;
					}
					goto IL_01d9;
					IL_026b:
					num2 = 26;
					borderDiagUp = BorderDiagUp;
					A3 = range.Borders;
					borderDiagUp.A(ref A3, XlBordersIndex.xlDiagonalUp);
					goto IL_0284;
					end_IL_0000_2:
					break;
				}
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 808;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	internal void B(ref Range A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				JG borderDiagDown;
				JG borderTop;
				JG borderBottom;
				JG borderLeft;
				JG borderRight;
				KG font;
				Font A3;
				JG borderDiagUp;
				Borders A2;
				LG interior;
				Interior A4;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 767:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000c;
						case 4:
							goto IL_0034;
						case 5:
							goto IL_0044;
						case 6:
							goto IL_0054;
						case 7:
							goto IL_006b;
						case 8:
							goto IL_007b;
						case 9:
							goto IL_0089;
						case 10:
							goto IL_009f;
						case 11:
							goto IL_00b3;
						case 12:
							goto IL_00da;
						case 13:
							goto IL_00f0;
						case 14:
							goto IL_0106;
						case 15:
							goto IL_011a;
						case 16:
							goto IL_0130;
						case 17:
							goto IL_0146;
						case 18:
							goto IL_0161;
						case 19:
							goto IL_017b;
						case 20:
							goto IL_0196;
						case 21:
							goto IL_01b2;
						case 22:
							goto IL_01ce;
						case 23:
							goto IL_01e9;
						case 24:
							goto IL_0205;
						case 25:
							goto IL_0220;
						case 26:
							goto IL_023c;
						case 27:
							goto IL_0257;
						case 28:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 29:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0257:
					num2 = 27;
					borderDiagDown = BorderDiagDown;
					A2 = range.Borders;
					borderDiagDown.B(ref A2, XlBordersIndex.xlDiagonalDown);
					break;
					IL_0007:
					num2 = 2;
					range = A;
					goto IL_000c;
					IL_000c:
					num2 = 3;
					NewLateBinding.LateSet(range, null, VH.A(1998), new object[1] { Formula }, null, null);
					goto IL_0034;
					IL_0034:
					num2 = 4;
					if (Information.Err().Number != 0)
					{
						goto IL_0044;
					}
					goto IL_0054;
					IL_0044:
					num2 = 5;
					range.Formula = Formula;
					goto IL_0054;
					IL_0054:
					num2 = 6;
					if (Operators.CompareString(FormulaArray, "", TextCompare: false) != 0)
					{
						goto IL_006b;
					}
					goto IL_007b;
					IL_006b:
					num2 = 7;
					range.FormulaArray = FormulaArray;
					goto IL_007b;
					IL_007b:
					num2 = 8;
					range.NumberFormat = NumberFormat;
					goto IL_0089;
					IL_0089:
					num2 = 9;
					range.RowHeight = RowHeight;
					goto IL_009f;
					IL_009f:
					num2 = 10;
					range.ColumnWidth = ColumnWidth;
					goto IL_00b3;
					IL_00b3:
					num2 = 11;
					if (KH.A.UndoAlignment)
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
						goto IL_00da;
					}
					goto IL_0146;
					IL_01ce:
					num2 = 22;
					borderTop = BorderTop;
					A2 = range.Borders;
					borderTop.B(ref A2, XlBordersIndex.xlEdgeTop);
					goto IL_01e9;
					IL_01e9:
					num2 = 23;
					borderBottom = BorderBottom;
					A2 = range.Borders;
					borderBottom.B(ref A2, XlBordersIndex.xlEdgeBottom);
					goto IL_0205;
					IL_0205:
					num2 = 24;
					borderLeft = BorderLeft;
					A2 = range.Borders;
					borderLeft.B(ref A2, XlBordersIndex.xlEdgeLeft);
					goto IL_0220;
					IL_00da:
					num2 = 12;
					range.IndentLevel = IndentLevel;
					goto IL_00f0;
					IL_00f0:
					num2 = 13;
					range.VerticalAlignment = VerticalAlignment;
					goto IL_0106;
					IL_0106:
					num2 = 14;
					range.HorizontalAlignment = HorizontalAlignment;
					goto IL_011a;
					IL_011a:
					num2 = 15;
					range.Orientation = Orientation;
					goto IL_0130;
					IL_0130:
					num2 = 16;
					range.WrapText = WrapText;
					goto IL_0146;
					IL_0146:
					num2 = 17;
					if (KH.A.UndoFont)
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
						goto IL_0161;
					}
					goto IL_017b;
					IL_0220:
					num2 = 25;
					borderRight = BorderRight;
					A2 = range.Borders;
					borderRight.B(ref A2, XlBordersIndex.xlEdgeRight);
					goto IL_023c;
					IL_0161:
					num2 = 18;
					font = Font;
					A3 = range.Font;
					font.B(ref A3);
					goto IL_017b;
					IL_017b:
					num2 = 19;
					if (KH.A.UndoFill)
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
						goto IL_0196;
					}
					goto IL_01b2;
					IL_023c:
					num2 = 26;
					borderDiagUp = BorderDiagUp;
					A2 = range.Borders;
					borderDiagUp.B(ref A2, XlBordersIndex.xlDiagonalUp);
					goto IL_0257;
					IL_0196:
					num2 = 20;
					interior = Interior;
					A4 = range.Interior;
					interior.B(ref A4);
					goto IL_01b2;
					IL_01b2:
					num2 = 21;
					if (!KH.A.UndoBorders)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_01ce;
					end_IL_0000_2:
					break;
				}
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 767;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	internal void A(ref CellProp A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		CellProp cellProp = default(CellProp);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 568:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000c;
						case 4:
							goto IL_001a;
						case 5:
							goto IL_0046;
						case 6:
							goto IL_0056;
						case 7:
							goto IL_0066;
						case 8:
							goto IL_0076;
						case 9:
							goto IL_0086;
						case 10:
							goto IL_0095;
						case 11:
							goto IL_00a6;
						case 12:
							goto IL_00b7;
						case 13:
							goto IL_00c8;
						case 14:
							goto IL_00d9;
						case 15:
							goto IL_00ea;
						case 16:
							goto IL_0103;
						case 17:
							goto IL_0114;
						case 18:
							goto IL_0125;
						case 19:
							goto IL_0136;
						case 20:
							goto IL_0151;
						case 21:
							goto IL_0162;
						case 22:
							goto IL_0173;
						case 23:
							goto IL_0184;
						case 24:
							goto IL_0193;
						case 25:
							goto IL_01a4;
						case 26:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 27:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0151:
					num2 = 20;
					BorderTop = cellProp.BorderTop;
					goto IL_0162;
					IL_0007:
					num2 = 2;
					cellProp = A;
					goto IL_000c;
					IL_000c:
					num2 = 3;
					Formula = cellProp.Formula;
					goto IL_001a;
					IL_001a:
					num2 = 4;
					if (Operators.CompareString(cellProp.FormulaArray, "", TextCompare: false) != 0)
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
						goto IL_0046;
					}
					goto IL_0056;
					IL_0162:
					num2 = 21;
					BorderBottom = cellProp.BorderBottom;
					goto IL_0173;
					IL_0173:
					num2 = 22;
					BorderLeft = cellProp.BorderLeft;
					goto IL_0184;
					IL_0184:
					num2 = 23;
					BorderRight = cellProp.BorderRight;
					goto IL_0193;
					IL_0046:
					num2 = 5;
					FormulaArray = cellProp.FormulaArray;
					goto IL_0056;
					IL_0056:
					num2 = 6;
					NumberFormat = cellProp.NumberFormat;
					goto IL_0066;
					IL_0066:
					num2 = 7;
					RowHeight = cellProp.RowHeight;
					goto IL_0076;
					IL_0076:
					num2 = 8;
					ColumnWidth = cellProp.ColumnWidth;
					goto IL_0086;
					IL_0086:
					num2 = 9;
					if (KH.A.UndoAlignment)
					{
						goto IL_0095;
					}
					goto IL_00ea;
					IL_0095:
					num2 = 10;
					IndentLevel = cellProp.IndentLevel;
					goto IL_00a6;
					IL_00a6:
					num2 = 11;
					HorizontalAlignment = cellProp.HorizontalAlignment;
					goto IL_00b7;
					IL_00b7:
					num2 = 12;
					VerticalAlignment = cellProp.VerticalAlignment;
					goto IL_00c8;
					IL_00c8:
					num2 = 13;
					Orientation = cellProp.Orientation;
					goto IL_00d9;
					IL_00d9:
					num2 = 14;
					WrapText = cellProp.WrapText;
					goto IL_00ea;
					IL_00ea:
					num2 = 15;
					if (KH.A.UndoFont)
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
						goto IL_0103;
					}
					goto IL_0114;
					IL_0193:
					num2 = 24;
					BorderDiagUp = cellProp.BorderDiagUp;
					goto IL_01a4;
					IL_0103:
					num2 = 16;
					Font = cellProp.Font;
					goto IL_0114;
					IL_0114:
					num2 = 17;
					if (KH.A.UndoFill)
					{
						goto IL_0125;
					}
					goto IL_0136;
					IL_0125:
					num2 = 18;
					Interior = cellProp.Interior;
					goto IL_0136;
					IL_0136:
					num2 = 19;
					if (!KH.A.UndoBorders)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_0151;
					IL_01a4:
					num2 = 25;
					BorderDiagDown = cellProp.BorderDiagDown;
					break;
					end_IL_0000_2:
					break;
				}
				cellProp = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 568;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}
