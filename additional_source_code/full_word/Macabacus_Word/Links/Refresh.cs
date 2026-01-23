using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Refresh
{
	public static void SelectedLinks()
	{
		if (!Access.AllowWordOperation((PlanType)5, (Restriction)1, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
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
			Selection selection = PC.A.Application.Selection;
			List<InlineShape> a = new List<InlineShape>();
			List<Microsoft.Office.Interop.Word.Shape> b = new List<Microsoft.Office.Interop.Word.Shape>();
			List<Table> c = new List<Table>();
			List<ContentControl> d = new List<ContentControl>();
			WdSelectionType type = selection.Type;
			if (type != WdSelectionType.wdSelectionInlineShape)
			{
				if (type != WdSelectionType.wdSelectionShape)
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
					try
					{
						a = selection.Range.InlineShapes.Cast<InlineShape>().ToList();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					try
					{
						try
						{
							enumerator = selection.ChildShapeRange.GetEnumerator();
							while (enumerator.MoveNext())
							{
								A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, b);
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (5)
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
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					if (Common.IsContentControlSelected(selection))
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
						try
						{
							d = Common.LinkedContentControlsInSelection(selection);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
					}
					else if (Common.IsTableSelected(selection))
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
							c = selection.Tables.Cast<Table>().ToList();
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
					}
				}
				else
				{
					try
					{
						try
						{
							enumerator2 = Helpers.SelectedShapes(selection).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								A((Microsoft.Office.Interop.Word.Shape)enumerator2.Current, b);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0140;
								}
								continue;
								end_IL_0140:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (7)
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
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						ProjectData.ClearProjectError();
					}
				}
			}
			else
			{
				try
				{
					a = selection.Range.InlineShapes.Cast<InlineShape>().ToList();
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
				try
				{
					enumerator3 = selection.ChildShapeRange.GetEnumerator();
					try
					{
						while (enumerator3.MoveNext())
						{
							A((Microsoft.Office.Interop.Word.Shape)enumerator3.Current, b);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00d6;
							}
							continue;
							end_IL_00d6:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator3 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					ProjectData.ClearProjectError();
				}
			}
			A(a, b, c, d, E: false);
			selection = null;
			a = null;
			b = null;
			c = null;
			d = null;
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, List<Microsoft.Office.Interop.Word.Shape> B)
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
					B.Add(A);
					return;
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Refresh.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, B);
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static void UpdateAllLinks()
	{
		if (!Access.AllowWordOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		Document activeDocument = PC.A.Application.ActiveDocument;
		List<InlineShape> B = new List<InlineShape>();
		List<Microsoft.Office.Interop.Word.Shape> C = new List<Microsoft.Office.Interop.Word.Shape>();
		List<Table> D = new List<Table>();
		List<ContentControl> E = new List<ContentControl>();
		_ = activeDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activeDocument.StoryRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Range range = (Microsoft.Office.Interop.Word.Range)enumerator.Current;
				do
				{
					A(range, ref B, ref C, ref D, ref E);
					range = range.NextStoryRange;
				}
				while (range != null);
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
			}
			while (true)
			{
				switch (2)
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
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		A(B, C, D, E, E: true);
		activeDocument = null;
		B = null;
		C = null;
		D = null;
		E = null;
	}

	private static void A(Microsoft.Office.Interop.Word.Range A, ref List<InlineShape> B, ref List<Microsoft.Office.Interop.Word.Shape> C, ref List<Table> D, ref List<ContentControl> E)
	{
		try
		{
			B.AddRange(A.InlineShapes.Cast<InlineShape>());
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			C.AddRange(A.ShapeRange.Cast<Microsoft.Office.Interop.Word.Shape>());
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			D.AddRange(A.Tables.Cast<Table>());
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			E.AddRange(A.ContentControls.Cast<ContentControl>());
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(List<InlineShape> A, List<Microsoft.Office.Interop.Word.Shape> B, List<Table> C, List<ContentControl> D, bool E)
	{
		new wpfLinkRefresh(A, B, C, D, E).Show();
		_ = null;
	}

	public static object UpdateShapeLink(Microsoft.Office.Interop.Word.Shape shp, ref RefreshInstance refreshInstance, ref bool blnSelectObject, CopierAsPicture copierAsPic)
	{
		return A(shp, ref refreshInstance, blnSelectObject, copierAsPic);
	}

	public static object UpdateShapeLink(InlineShape shp, ref RefreshInstance refreshInstance, ref bool blnSelectObject, CopierAsPicture copierAsPic)
	{
		return A(shp, ref refreshInstance, blnSelectObject, copierAsPic);
	}

	public static object UpdateShapeLink(Table tbl, ref RefreshInstance refreshInstance, ref bool blnSelectObject, CopierAsPicture copierAsPic)
	{
		return A(tbl, ref refreshInstance, blnSelectObject, copierAsPic);
	}

	public static object UpdateShapeLink(ContentControl cc, ref RefreshInstance refreshInstance, ref bool blnSelectObject, CopierAsPicture copierAsPic)
	{
		return A(cc, ref refreshInstance, blnSelectObject, copierAsPic);
	}

	private static object A(object A, ref RefreshInstance B, bool C, CopierAsPicture D)
	{
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_01db: Unknown result type (might be due to invalid IL or missing references)
		//IL_01dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01df: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Expected O, but got Unknown
		//IL_0d6a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d56: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_032d: Unknown result type (might be due to invalid IL or missing references)
		//IL_033a: Unknown result type (might be due to invalid IL or missing references)
		//IL_033c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0342: Invalid comparison between Unknown and I4
		//IL_0473: Unknown result type (might be due to invalid IL or missing references)
		//IL_0475: Unknown result type (might be due to invalid IL or missing references)
		//IL_047b: Invalid comparison between Unknown and I4
		//IL_049f: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a7: Invalid comparison between Unknown and I4
		//IL_04d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d8: Invalid comparison between Unknown and I4
		//IL_0501: Unknown result type (might be due to invalid IL or missing references)
		//IL_0503: Unknown result type (might be due to invalid IL or missing references)
		//IL_0509: Invalid comparison between Unknown and I4
		//IL_04b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_04bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_049a: Unknown result type (might be due to invalid IL or missing references)
		//IL_057f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0584: Unknown result type (might be due to invalid IL or missing references)
		//IL_0588: Unknown result type (might be due to invalid IL or missing references)
		//IL_058a: Unknown result type (might be due to invalid IL or missing references)
		//IL_058f: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_05fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_060b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0619: Unknown result type (might be due to invalid IL or missing references)
		//IL_0627: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_0659: Unknown result type (might be due to invalid IL or missing references)
		//IL_0660: Unknown result type (might be due to invalid IL or missing references)
		//IL_0662: Unknown result type (might be due to invalid IL or missing references)
		//IL_0667: Unknown result type (might be due to invalid IL or missing references)
		//IL_0669: Unknown result type (might be due to invalid IL or missing references)
		//IL_066c: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a2: Expected I4, but got Unknown
		//IL_063d: Unknown result type (might be due to invalid IL or missing references)
		//IL_063f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0644: Unknown result type (might be due to invalid IL or missing references)
		//IL_064b: Unknown result type (might be due to invalid IL or missing references)
		//IL_06f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_073c: Unknown result type (might be due to invalid IL or missing references)
		//IL_07d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_07eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_07fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_072c: Unknown result type (might be due to invalid IL or missing references)
		//IL_07c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0760: Unknown result type (might be due to invalid IL or missing references)
		//IL_0767: Unknown result type (might be due to invalid IL or missing references)
		//IL_0829: Unknown result type (might be due to invalid IL or missing references)
		//IL_0535: Unknown result type (might be due to invalid IL or missing references)
		//IL_053c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0780: Unknown result type (might be due to invalid IL or missing references)
		//IL_0782: Unknown result type (might be due to invalid IL or missing references)
		//IL_0787: Unknown result type (might be due to invalid IL or missing references)
		//IL_0789: Unknown result type (might be due to invalid IL or missing references)
		//IL_0557: Unknown result type (might be due to invalid IL or missing references)
		//IL_0559: Unknown result type (might be due to invalid IL or missing references)
		//IL_055e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0560: Unknown result type (might be due to invalid IL or missing references)
		//IL_0aaf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab5: Expected O, but got Unknown
		//IL_0ae5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0aeb: Expected I4, but got Unknown
		//IL_095e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0964: Expected O, but got Unknown
		//IL_0992: Unknown result type (might be due to invalid IL or missing references)
		//IL_0998: Expected I4, but got Unknown
		Workbook workbook = null;
		Microsoft.Office.Interop.Excel.Range range = null;
		Microsoft.Office.Interop.Excel.Chart chart = null;
		object obj = null;
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		Type typeFromHandle = typeof(Common);
		string memberName = XC.A(11777);
		object[] obj2 = new object[1] { A };
		object[] array = obj2;
		bool[] obj3 = new bool[1] { true };
		bool[] array2 = obj3;
		object obj4 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj2, null, null, obj3);
		if (array2[0])
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		_003F val;
		if (obj4 == null)
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
			val = default(Link);
		}
		else
		{
			val = (Link)obj4;
		}
		Link val2 = (Link)val;
		bool flag4 = Base.SourceIsRange(val2);
		B.LocateSource(ref workbook, ref range, ref chart, ref val2, ref flag, ref flag2, flag4);
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
			Type typeFromHandle2 = typeof(Update);
			string memberName2 = XC.A(13872);
			object[] obj5 = new object[4] { A, B, null, null };
			ref string source = ref val2.Source;
			ref string reference = ref source;
			obj5[2] = source;
			obj5[3] = false;
			array = obj5;
			bool[] obj6 = new bool[4] { true, true, true, false };
			array2 = obj6;
			NewLateBinding.LateCall(null, typeFromHandle2, memberName2, obj5, null, null, obj6, IgnoreReturn: true);
			if (array2[0])
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
				A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (array2[1])
			{
				B = (RefreshInstance)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(RefreshInstance));
			}
			if (array2[2])
			{
				reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(string));
			}
		}
		string strAddress;
		bool flag5;
		Link val3;
		Link val4;
		CustomXMLPart customXMLPart;
		int num;
		bool flag6;
		string text;
		Microsoft.Office.Interop.Word.Application application;
		Worksheet worksheet;
		Name name;
		if (!flag)
		{
			if (range == null)
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
				if (chart == null)
				{
					if (flag4)
					{
						workbook = null;
						throw new UpdateLinkException(XC.A(14110));
					}
					workbook = null;
					throw new UpdateLinkException(XC.A(14211));
				}
			}
			application = PC.A.Application;
			text = "";
			strAddress = "";
			name = null;
			worksheet = null;
			flag5 = false;
			flag6 = false;
			val3 = val2;
			val4 = val2;
			object obj7 = NewLateBinding.LateGet(null, typeof(Macabacus_Word.CustomXML), XC.A(3113), array = new object[1] { A }, null, null, array2 = new bool[1] { true });
			if (array2[0])
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
				A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			customXMLPart = (CustomXMLPart)obj7;
			if (!(A is InlineShape))
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
				if (!(A is Microsoft.Office.Interop.Word.Shape))
				{
					num = 0;
					goto IL_0291;
				}
			}
			num = (Conversions.ToBoolean(Operators.CompareObjectEqual(NewLateBinding.LateGet(A, null, XC.A(13885), new object[0], null, null, null), MsoTriState.msoTrue, TextCompare: false)) ? 1 : 0);
			goto IL_0291;
		}
		goto IL_0d70;
		IL_0291:
		flag6 = Conversions.ToBoolean((byte)num != 0);
		XlSheetVisibility xlSheetVisibility = default(XlSheetVisibility);
		if (flag4)
		{
			B.SourceRange(val4, workbook, ref name, ref worksheet, ref range, ref strAddress, ref xlSheetVisibility);
		}
		else
		{
			if (flag6 && Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, XC.A(13902), new object[0], null, null, null), null, XC.A(13913), new object[1] { Microsoft.Office.Core.XlAxisType.xlValue }, null, null, null), MsoTriState.msoTrue, TextCompare: false))
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
				text = Refresh.A(RuntimeHelpers.GetObjectValue(A));
			}
			B.SourceChart(val4, workbook, ref chart);
			flag5 = true;
		}
		if ((int)val4.Type != 4)
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
			Type typeFromHandle3 = typeof(Navigate);
			string memberName3 = XC.A(13928);
			object[] obj8 = new object[1] { A };
			array = obj8;
			bool[] obj9 = new bool[1] { true };
			array2 = obj9;
			object operand = NewLateBinding.LateGet(null, typeFromHandle3, memberName3, obj8, null, null, obj9);
			if (array2[0])
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
				A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (Conversions.ToBoolean(Operators.NotObject(operand)))
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
				NewLateBinding.LateCall(A, null, XC.A(12515), new object[0], null, null, null, IgnoreReturn: true);
			}
			else
			{
				Microsoft.Office.Interop.Word.Range range2 = null;
				if (A is InlineShape)
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
					range2 = ((InlineShape)A).Range;
				}
				else if (A is Microsoft.Office.Interop.Word.Shape)
				{
					range2 = ((Microsoft.Office.Interop.Word.Shape)A).Anchor;
				}
				else if (A is Table)
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
					range2 = ((Table)A).Range;
				}
				if (range2 != null)
				{
					Navigate.B(range2);
					NewLateBinding.LateCall(A, null, XC.A(12515), new object[0], null, null, null, IgnoreReturn: true);
					range2 = null;
				}
			}
		}
		if ((int)val4.Type == 2)
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
			if (range.Cells.Count == 1)
			{
				val4.Type = (ImportType)4;
			}
		}
		if ((int)val4.Type == 4)
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
			if (val4.Type == val3.Type)
			{
				goto IL_0657;
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
		}
		if ((int)val4.Type == 5)
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
			if (val4.Type == val3.Type)
			{
				goto IL_0657;
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
		}
		if ((int)val4.Type == 2)
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
			if (!NC.A.RebuildTables)
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
				if (A is Table && Operators.CompareString(val4.Name, val3.Name, TextCompare: false) == 0)
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
					if (val4.Type == val3.Type)
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
						flag3 = true;
						goto IL_0657;
					}
				}
			}
		}
		SelectionParameters val5 = Refresh.A(RuntimeHelpers.GetObjectValue(A));
		PasteParameters val6 = default(PasteParameters);
		val6.PasteType = val4.Type;
		val6.MatchSize = clsImportExport.GetMatchSize(N.Settings.ImportMatchDestinationWidth, N.Settings.ImportMatchDestinationHeight);
		val6.SourceRange = range;
		val6.SourceChart = chart;
		val6.CallingApplication = (CallingApp)3;
		val6.TargetApplication = application;
		val6.TargetWidth = val5.Width;
		val6.TargetHeight = val5.Height;
		val6.TargetTop = val5.Top;
		val6.TargetLeft = val5.Left;
		val6.TargetZOrder = val5.ZOrder;
		val6.Placement = val5.Placement;
		if (A is Microsoft.Office.Interop.Word.Shape)
		{
			val6.WrapFormat = val5.WrapFormat;
			val6.RelVertPosn = val5.RelVertPosn;
		}
		goto IL_0657;
		IL_0d70:
		if (obj != null)
		{
			return obj;
		}
		return A;
		IL_0657:
		B.CheckCalculationMode(val4);
		ImportType type = val4.Type;
		switch (type - 1)
		{
		case 3:
			if (A is ContentControl)
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
				((ContentControl)A).Range.Text = range.Text.ToString().Trim();
			}
			else
			{
				obj = CellAsWordText.Paste(application, range);
			}
			break;
		case 0:
		case 10:
			obj = RuntimeHelpers.GetObjectValue(CellsAsWordPicture.Paste(val6, ref B, D, true));
			break;
		case 1:
			if (flag3)
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
				Refresh.A((Table)A, range);
			}
			else
			{
				obj = CellsAsWordTable.Paste(val6);
			}
			break;
		case 2:
			obj = RuntimeHelpers.GetObjectValue(CellsAsWordEmbedded.Paste(val6, ref B));
			break;
		case 4:
			if (flag6)
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
				if (Operators.CompareString(val4.Name, val3.Name, TextCompare: false) == 0)
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
					if (val4.Type == val3.Type)
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
						Refresh.A((Microsoft.Office.Interop.Word.Chart)NewLateBinding.LateGet(A, null, XC.A(13902), new object[0], null, null, null), range);
						break;
					}
				}
			}
			obj = RuntimeHelpers.GetObjectValue(CellsAsWordChart.Paste(val6));
			break;
		case 5:
		case 11:
			obj = RuntimeHelpers.GetObjectValue(ChartAsWordPicture.Paste(val6, D, true));
			break;
		case 6:
			obj = RuntimeHelpers.GetObjectValue(ChartAsWordChart.Paste(val6));
			break;
		case 7:
			obj = RuntimeHelpers.GetObjectValue(ChartAsWordEmbedded.Paste(val6));
			break;
		}
		if (obj == null)
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
			CustomXML.UpdatePart(customXMLPart, B, workbook.FullName, val4, strAddress, Base.LastUpdate(), application.UserName);
			if (C)
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
				Refresh.A(RuntimeHelpers.GetObjectValue(A));
			}
		}
		else
		{
			if (flag5)
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
				Type typeFromHandle4 = typeof(Add);
				string memberName4 = XC.A(13961);
				object[] obj10 = new object[5] { obj, chart, B, null, null };
				ref ImportType type2 = ref val4.Type;
				ref ImportType reference2 = ref type2;
				obj10[3] = type2;
				ref string parentId = ref val4.ParentId;
				ref string reference = ref parentId;
				obj10[4] = parentId;
				array = obj10;
				bool[] obj11 = new bool[5] { true, true, true, true, true };
				array2 = obj11;
				NewLateBinding.LateCall(null, typeFromHandle4, memberName4, obj10, null, null, obj11, IgnoreReturn: true);
				if (array2[0])
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
					obj = RuntimeHelpers.GetObjectValue(array[0]);
				}
				if (array2[1])
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
					chart = (Microsoft.Office.Interop.Excel.Chart)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(Microsoft.Office.Interop.Excel.Chart));
				}
				if (array2[2])
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
					B = (RefreshInstance)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(RefreshInstance));
				}
				if (array2[3])
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
					reference2 = (ImportType)(int)(ImportType)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[3]), typeof(ImportType));
				}
				if (array2[4])
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
					reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[4]), typeof(string));
				}
			}
			else
			{
				Type typeFromHandle5 = typeof(Add);
				string memberName5 = XC.A(13992);
				object[] obj12 = new object[6] { obj, range, B, null, null, null };
				ref ImportType type3 = ref val4.Type;
				ref ImportType reference2 = ref type3;
				obj12[3] = type3;
				ref string parentId2 = ref val4.ParentId;
				ref string reference = ref parentId2;
				obj12[4] = parentId2;
				obj12[5] = name;
				array = obj12;
				bool[] obj13 = new bool[6] { true, true, true, true, true, true };
				array2 = obj13;
				NewLateBinding.LateCall(null, typeFromHandle5, memberName5, obj12, null, null, obj13, IgnoreReturn: true);
				if (array2[0])
				{
					obj = RuntimeHelpers.GetObjectValue(array[0]);
				}
				if (array2[1])
				{
					range = (Microsoft.Office.Interop.Excel.Range)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(Microsoft.Office.Interop.Excel.Range));
				}
				if (array2[2])
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
					B = (RefreshInstance)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(RefreshInstance));
				}
				if (array2[3])
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
					reference2 = (ImportType)(int)(ImportType)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[3]), typeof(ImportType));
				}
				if (array2[4])
				{
					reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[4]), typeof(string));
				}
				if (array2[5])
				{
					name = (Name)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[5]), typeof(Name));
				}
			}
			if (C)
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
				Refresh.A(RuntimeHelpers.GetObjectValue(obj));
			}
			if (customXMLPart != null)
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
				customXMLPart.Delete();
			}
			if (text.Length > 0)
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
				if (!(obj is InlineShape))
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
					if (!(obj is Microsoft.Office.Interop.Word.Shape))
					{
						goto IL_0cf5;
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
				if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(obj, null, XC.A(13885), new object[0], null, null, null), MsoTriState.msoTrue, TextCompare: false))
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
					if (new Regex(XC.A(14023)).IsMatch(text))
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
						Type typeFromHandle6 = typeof(Refresh);
						string memberName6 = XC.A(14050);
						object[] obj14 = new object[2] { obj, text };
						array = obj14;
						bool[] obj15 = new bool[2] { true, true };
						array2 = obj15;
						NewLateBinding.LateCall(null, typeFromHandle6, memberName6, obj14, null, null, obj15, IgnoreReturn: true);
						if (array2[0])
						{
							obj = RuntimeHelpers.GetObjectValue(array[0]);
						}
						if (array2[1])
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
							text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(string));
						}
						NewLateBinding.LateCall(null, typeof(clsCharts), XC.A(14087), array = new object[1] { obj }, null, null, array2 = new bool[1] { true }, IgnoreReturn: true);
						if (array2[0])
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
							obj = RuntimeHelpers.GetObjectValue(array[0]);
						}
					}
				}
			}
		}
		goto IL_0cf5;
		IL_0cf5:
		if (worksheet != null)
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
			if (worksheet.Visible != xlSheetVisibility)
			{
				worksheet.Visible = xlSheetVisibility;
			}
		}
		B.RestoreExcel();
		RefreshInstance val7;
		(val7 = B).RefreshedLinkCount = checked(val7.RefreshedLinkCount + 1);
		application = null;
		worksheet = null;
		name = null;
		range = null;
		chart = null;
		workbook = null;
		goto IL_0d70;
	}

	private static void A(object A)
	{
		Type typeFromHandle = typeof(Navigate);
		string memberName = XC.A(13928);
		object[] obj = new object[1] { A };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object operand = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
		if (array2[0])
		{
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		int num;
		if (!Conversions.ToBoolean(Operators.NotObject(operand)))
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
			num = (Navigate.A(PC.A.Application) ? 1 : 0);
		}
		else
		{
			num = 1;
		}
		if (!Conversions.ToBoolean((byte)num != 0))
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
			B(RuntimeHelpers.GetObjectValue(A));
			return;
		}
	}

	private static void B(object A)
	{
		if (A is ContentControl)
		{
			((ContentControl)A).Range.Select();
			return;
		}
		if (A is InlineShape)
		{
			((InlineShape)A).Select();
			return;
		}
		if (A is Microsoft.Office.Interop.Word.Shape)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Microsoft.Office.Interop.Word.Shape obj = (Microsoft.Office.Interop.Word.Shape)A;
					object Replace = RuntimeHelpers.GetObjectValue(Missing.Value);
					obj.Select(ref Replace);
					return;
				}
				}
			}
		}
		if (A is Table)
		{
			((Table)A).Select();
		}
	}

	private static SelectionParameters A(object A)
	{
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_006e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		Type typeFromHandle = typeof(ExcelToWord);
		string memberName = XC.A(14334);
		object[] obj = new object[1] { A };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
		if (array2[0])
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
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		if (obj3 == null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return default(SelectionParameters);
				}
			}
		}
		return (SelectionParameters)obj3;
	}

	private static void A(Table A, Microsoft.Office.Interop.Excel.Range B)
	{
		int count = B.Rows.Count;
		Rows rows = A.Rows;
		while (rows.Count < count)
		{
			try
			{
				Rows rows2 = rows;
				object BeforeRow = 2;
				rows2.Add(ref BeforeRow);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Rows rows3 = rows;
				object BeforeRow = RuntimeHelpers.GetObjectValue(Missing.Value);
				rows3.Add(ref BeforeRow);
				ProjectData.ClearProjectError();
			}
		}
		checked
		{
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
				while (rows.Count > count)
				{
					rows[rows.Count - 1].Delete();
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					rows = null;
					int count2 = B.Columns.Count;
					Columns columns = A.Columns;
					while (columns.Count < count2)
					{
						try
						{
							Columns columns2 = columns;
							object BeforeRow = 2;
							columns2.Add(ref BeforeRow);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							Columns columns3 = columns;
							object BeforeRow = RuntimeHelpers.GetObjectValue(Missing.Value);
							columns3.Add(ref BeforeRow);
							ProjectData.ClearProjectError();
						}
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
					while (columns.Count > count2)
					{
						columns[columns.Count - 1].Delete();
					}
					columns = null;
					int num = count;
					for (int i = 1; i <= num; i++)
					{
						int num2 = count2;
						for (int j = 1; j <= num2; j++)
						{
							A.Cell(i, j).Range.Text = Conversions.ToString(NewLateBinding.LateGet(B.Cells[i, j], null, XC.A(14361), new object[0], null, null, null));
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
			}
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Chart A, Microsoft.Office.Interop.Excel.Range B)
	{
		B.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
		Worksheet obj = (Worksheet)((Workbook)A.ChartData.Workbook).Worksheets[1];
		obj.Cells.Clear();
		NewLateBinding.LateCall(obj.Cells[1, 1], null, XC.A(14370), new object[1] { XlPasteType.xlPasteValuesAndNumberFormats }, null, null, null, IgnoreReturn: true);
		_ = null;
		_ = null;
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A)
	{
		Refresh.A(A.Chart);
	}

	private static void A(InlineShape A)
	{
		Refresh.A(A.Chart);
	}

	private static void A(Microsoft.Office.Interop.Word.Chart A)
	{
		A.ChartData.Activate();
	}

	public static void SetChartAxisFormat(Microsoft.Office.Interop.Word.Shape shp, string strFormat)
	{
		A(shp.Chart, strFormat);
	}

	public static void SetChartAxisFormat(InlineShape shp, string strFormat)
	{
		A(shp.Chart, strFormat);
	}

	private static void A(Microsoft.Office.Interop.Word.Chart A, string B)
	{
		try
		{
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(A.Axes(Microsoft.Office.Core.XlAxisType.xlValue), null, XC.A(14395), new object[0], null, null, null), null, XC.A(14416), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static string A(object A)
	{
		string result;
		try
		{
			result = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, XC.A(13902), new object[0], null, null, null), null, XC.A(14441), new object[1] { Microsoft.Office.Core.XlAxisType.xlValue }, null, null, null), null, XC.A(14395), new object[0], null, null, null), null, XC.A(14416), new object[0], null, null, null).ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
