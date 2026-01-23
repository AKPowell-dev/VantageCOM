using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Formulas;
using MacabacusMacros.Links;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Workbook
{
	[CompilerGenerated]
	internal sealed class Q
	{
		public Analysis A;

		public Settings A;

		public Microsoft.Office.Interop.Excel.Workbook A;

		[SpecialName]
		internal void A(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.A();
		}

		[SpecialName]
		internal void A()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			Workbook.A(ref A, this.A.LegacyFileType, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool A()
		{
			return this.A.LegacyFileType != Severity.Ignore;
		}

		[SpecialName]
		internal bool B()
		{
			return this.A.OldFile != Severity.Ignore;
		}

		[SpecialName]
		internal bool C()
		{
			return this.A.LargeFileSize != Severity.Ignore;
		}

		[SpecialName]
		internal void B(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.B();
		}

		[SpecialName]
		internal void B()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			BestPractices.A(ref A, this.A.CoverMissing, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool D()
		{
			return this.A.CoverMissing != Severity.Ignore;
		}

		[SpecialName]
		internal void C(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.C();
		}

		[SpecialName]
		internal void C()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			Workbook.A(ref A, this.A, this.A, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool E()
		{
			return this.A.ExcessNames != Severity.Ignore;
		}

		[SpecialName]
		internal void D(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.D();
		}

		[SpecialName]
		internal void D()
		{
			Workbook.A(this.A.Observations, this.A, this.A, this.A);
		}

		[SpecialName]
		internal bool F()
		{
			return this.A.HiddenNames != Severity.Ignore;
		}

		[SpecialName]
		internal void E(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.E();
		}

		[SpecialName]
		internal void E()
		{
			Workbook.B(this.A.Observations, this.A, this.A, this.A);
		}

		[SpecialName]
		internal bool G()
		{
			return this.A.UnusedNames != Severity.Ignore;
		}

		[SpecialName]
		internal void F(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.F();
		}

		[SpecialName]
		internal void F()
		{
			Workbook.C(this.A.Observations, this.A, this.A, this.A);
		}

		[SpecialName]
		internal bool H()
		{
			return this.A.NamesWithExternalReferences != Severity.Ignore;
		}

		[SpecialName]
		internal void G(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.G();
		}

		[SpecialName]
		internal void G()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			Workbook.A(ref A, this.A, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool I()
		{
			return this.A.ExcessStyles != Severity.Ignore;
		}

		[SpecialName]
		internal void H(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.H();
		}

		[SpecialName]
		internal void H()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			Workbook.B(ref A, this.A.DisplayDrawingObjects, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool J()
		{
			return this.A.DisplayDrawingObjects != Severity.Ignore;
		}

		[SpecialName]
		internal void I(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.I();
		}

		[SpecialName]
		internal void I()
		{
			CircularReferences.A(this.A, this.A.CircularReferences, this.A);
		}

		[SpecialName]
		internal bool K()
		{
			return this.A.CircularReferences != Severity.Ignore;
		}
	}

	[CompilerGenerated]
	internal sealed class R
	{
		public DateTime A;

		public Exception A;

		public long A;

		public Q A;

		[SpecialName]
		internal void A(Microsoft.Office.Interop.Excel.Workbook A)
		{
			this.A();
		}

		[SpecialName]
		internal void A()
		{
			Analysis analysis = this.A.A;
			List<Observation> A = analysis.Observations;
			Workbook.A(ref A, this.A.A, this.A, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal void B(Microsoft.Office.Interop.Excel.Workbook A)
		{
			B();
		}

		[SpecialName]
		internal void B()
		{
			Analysis analysis = this.A.A;
			List<Observation> A = analysis.Observations;
			Workbook.A(ref A, this.A.A, this.A, this.A);
			analysis.Observations = A;
		}
	}

	[CompilerGenerated]
	internal sealed class S
	{
		public List<Observation> A;

		public Settings A;

		[SpecialName]
		internal bool A(Name A)
		{
			if (ExcelAddIn1.Formulas.Names.B(A))
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
				this.A.Add(new ExternalNameReference(this.A.NamesWithExternalReferences, A));
			}
			return true;
		}
	}

	[CompilerGenerated]
	internal sealed class T
	{
		public int A;

		[SpecialName]
		internal bool A(Name A)
		{
			checked
			{
				this.A++;
				return true;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class U
	{
		public List<Observation> A;

		public Settings A;

		[SpecialName]
		internal bool A(Name A)
		{
			if (A.Visible)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return true;
					}
				}
			}
			this.A.Add(new HiddenName(this.A.HiddenNames));
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class V
	{
		public List<Observation> A;

		public Settings A;

		[SpecialName]
		internal bool A(Name A)
		{
			if (!ExcelAddIn1.Formulas.Names.A(A, 200L))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return true;
					}
				}
			}
			if (!ExcelAddIn1.Formulas.Names.A(A, B: true))
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
				this.A.Add(new UnusedName(this.A.UnusedNames, A));
			}
			return true;
		}
	}

	internal static List<RB> A(Analysis A, Settings B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		Q CS_0024_003C_003E8__locals36 = new Q();
		CS_0024_003C_003E8__locals36.A = A;
		CS_0024_003C_003E8__locals36.A = B;
		CS_0024_003C_003E8__locals36.A = C;
		List<RB> list = new List<RB>();
		if (CS_0024_003C_003E8__locals36.A.Path.Length > 0)
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
			list.Add(new UB(VH.A(5829), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.A();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.LegacyFileType != Severity.Ignore));
			if (CS_0024_003C_003E8__locals36.A.OldFile == Severity.Ignore)
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
				if (CS_0024_003C_003E8__locals36.A.LargeFileSize == Severity.Ignore)
				{
					goto IL_017e;
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
			}
			R CS_0024_003C_003E8__locals20 = new R();
			CS_0024_003C_003E8__locals20.A = CS_0024_003C_003E8__locals36;
			CS_0024_003C_003E8__locals20.A = null;
			try
			{
				FileInfo fileInfo = new FileInfo(CS_0024_003C_003E8__locals20.A.A.FullName);
				CS_0024_003C_003E8__locals20.A = fileInfo.CreationTimeUtc;
				CS_0024_003C_003E8__locals20.A = fileInfo.Length;
				fileInfo = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a = ex;
				CS_0024_003C_003E8__locals20.A = new DB(a);
				ProjectData.ClearProjectError();
			}
			list.Add(new UB(VH.A(5862), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals20.A();
			}, [SpecialName] () => CS_0024_003C_003E8__locals20.A.A.OldFile != Severity.Ignore));
			list.Add(new UB(VH.A(5879), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals20.B();
			}, [SpecialName] () => CS_0024_003C_003E8__locals20.A.A.LargeFileSize != Severity.Ignore));
		}
		goto IL_017e;
		IL_017e:
		list.AddRange(new RB[8]
		{
			new UB(VH.A(5898), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.B();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.CoverMissing != Severity.Ignore),
			new UB(VH.A(5925), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.C();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.ExcessNames != Severity.Ignore),
			new UB(VH.A(5950), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.D();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.HiddenNames != Severity.Ignore),
			new UB(VH.A(5975), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.E();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.UnusedNames != Severity.Ignore),
			new UB(VH.A(6000), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.F();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.NamesWithExternalReferences != Severity.Ignore),
			new UB(VH.A(6049), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.G();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.ExcessStyles != Severity.Ignore),
			new UB(VH.A(6076), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.H();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.DisplayDrawingObjects != Severity.Ignore),
			new UB(VH.A(6105), [SpecialName] (Microsoft.Office.Interop.Excel.Workbook workbook) =>
			{
				CS_0024_003C_003E8__locals36.I();
			}, [SpecialName] () => CS_0024_003C_003E8__locals36.A.CircularReferences != Severity.Ignore)
		});
		return list;
	}

	private static void A(ref List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
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
			if (Operators.CompareString(Path.GetExtension(C.FullName), VH.A(6144), TextCompare: false) != 0)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				A.Add(new LegacyFileType(B));
				return;
			}
		}
	}

	private static void A(ref List<Observation> A, Settings B, DateTime C, Exception D)
	{
		if (B.OldFile == Severity.Ignore)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (D != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						throw D;
					}
				}
			}
			if (DateTime.Compare(DateTime.UtcNow, C.AddMonths(B.MaxFileAgeInMonths)) > 0)
			{
				A.Add(new FileTooOld(B.OldFile, C));
			}
			return;
		}
	}

	private static void A(ref List<Observation> A, Settings B, long C, Exception D)
	{
		if (B.LargeFileSize == Severity.Ignore)
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
			if (D != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						throw D;
					}
				}
			}
			if (C <= B.MaxFileSize)
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
				A.Add(new LargeFileSize(B.LargeFileSize, C));
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A, Analysis B, Func<Name, bool> C)
	{
		int count = A.Names.Count;
		Name name;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			if (B.A())
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			name = A.Names.Item(i, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (Workbook.A(name, B: false))
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
			if (ExcelAddIn1.Formulas.Names.D(name))
			{
				continue;
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
			if (ExcelAddIn1.Formulas.Names.E(name))
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
				break;
			}
			if (!C(name))
			{
				break;
			}
		}
		name = null;
	}

	internal static bool A(Name A, bool B = true)
	{
		bool result;
		try
		{
			if (!A.Name.StartsWith(Base.LINK_PREFIX))
			{
				result = false;
			}
			else if (B)
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
					Range refersToRange = A.RefersToRange;
					Microsoft.Office.Interop.Excel.Worksheet worksheet = refersToRange.Worksheet;
					Range cells = refersToRange.Cells;
					Range cells2 = worksheet.Cells;
					result = Conversions.ToLong(cells.CountLarge) != Conversions.ToLong(cells2.CountLarge);
					break;
				}
			}
			else
			{
				result = true;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
			Range cells2 = null;
		}
		return result;
	}

	private static void A(ref List<Observation> A, Settings B, Microsoft.Office.Interop.Excel.Workbook C, Analysis D)
	{
		if (B.ExcessNames == Severity.Ignore)
		{
			return;
		}
		int A2;
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
			A2 = 0;
			Workbook.A(C, D, checked([SpecialName] (Name name) =>
			{
				A2++;
				return true;
			}));
			if (A2 > B.MaxNamesCount)
			{
				A.Add(new ExcessNames(B.ExcessNames, A2));
			}
			return;
		}
	}

	private static void A(List<Observation> A, Settings B, Microsoft.Office.Interop.Excel.Workbook C, Analysis D)
	{
		if (B.HiddenNames == Severity.Ignore)
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
			Workbook.A(C, D, [SpecialName] (Name name) =>
			{
				if (name.Visible)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return true;
						}
					}
				}
				A.Add(new HiddenName(B.HiddenNames));
				return false;
			});
			return;
		}
	}

	private static void B(List<Observation> A, Settings B, Microsoft.Office.Interop.Excel.Workbook C, Analysis D)
	{
		if (B.UnusedNames == Severity.Ignore)
		{
			return;
		}
		Workbook.A(C, D, [SpecialName] (Name name) =>
		{
			if (!ExcelAddIn1.Formulas.Names.A(name, 200L))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return true;
					}
				}
			}
			if (!ExcelAddIn1.Formulas.Names.A(name, B: true))
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
				A.Add(new UnusedName(B.UnusedNames, name));
			}
			return true;
		});
	}

	private static void C(List<Observation> A, Settings B, Microsoft.Office.Interop.Excel.Workbook C, Analysis D)
	{
		if (B.NamesWithExternalReferences == Severity.Ignore)
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
			Workbook.A(C, D, [SpecialName] (Name name) =>
			{
				if (ExcelAddIn1.Formulas.Names.B(name))
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
					A.Add(new ExternalNameReference(B.NamesWithExternalReferences, name));
				}
				return true;
			});
			return;
		}
	}

	private static void A(ref List<Observation> A, Settings B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		if (B.ExcessStyles == Severity.Ignore)
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
			int count = C.Styles.Count;
			if (C.Styles.Count <= B.MaxStylesCount)
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
				A.Add(new ExcessStyles(B.ExcessStyles, count));
				return;
			}
		}
	}

	private static void B(ref List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (C.DisplayDrawingObjects == XlDisplayDrawingObjects.xlDisplayShapes)
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
				A.Add(new HiddenObjects(B));
				return;
			}
		}
	}
}
