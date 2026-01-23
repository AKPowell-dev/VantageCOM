using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class View
{
	public static void ViewSource()
	{
		try
		{
			Selection selection = PC.A.Application.ActiveWindow.Selection;
			WdSelectionType type = selection.Type;
			if (type != WdSelectionType.wdSelectionInlineShape)
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
				if (type == WdSelectionType.wdSelectionShape)
				{
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = Helpers.SelectedShapes(selection).GetEnumerator();
						while (enumerator.MoveNext())
						{
							A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0144;
							}
							continue;
							end_IL_0144:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (6)
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
				else if (Common.IsContentControlSelected(selection))
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
					List<ContentControl> list = Common.LinkedContentControlsInSelection(selection);
					using (List<ContentControl>.Enumerator enumerator2 = list.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							ContentControl current = enumerator2.Current;
							try
							{
								ViewSource(current);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Forms.ErrorMessage(ex2.Message);
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_01dc;
							}
							continue;
							end_IL_01dc:
							break;
						}
					}
					list = null;
				}
				else if (Common.IsTableSelected(selection))
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
					Tables tables = selection.Tables;
					IEnumerator enumerator3 = default(IEnumerator);
					try
					{
						enumerator3 = tables.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Table shp = (Table)enumerator3.Current;
							try
							{
								ViewSource(shp);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								Forms.ErrorMessage(ex4.Message);
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_026b;
							}
							continue;
							end_IL_026b:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
					tables = null;
				}
			}
			else
			{
				IEnumerator enumerator4 = default(IEnumerator);
				try
				{
					enumerator4 = selection.InlineShapes.GetEnumerator();
					while (enumerator4.MoveNext())
					{
						InlineShape shp2 = (InlineShape)enumerator4.Current;
						try
						{
							ViewSource(shp2);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							Forms.ErrorMessage(ex6.Message);
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0093;
						}
						continue;
						end_IL_0093:
						break;
					}
				}
				finally
				{
					if (enumerator4 is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator4 as IDisposable).Dispose();
							break;
						}
					}
				}
				{
					IEnumerator enumerator5 = selection.ChildShapeRange.GetEnumerator();
					try
					{
						while (enumerator5.MoveNext())
						{
							A((Microsoft.Office.Interop.Word.Shape)enumerator5.Current);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00ed;
							}
							continue;
							end_IL_00ed:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator5 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			Common.LogActivity(XC.A(14475));
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Selection selection = null;
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			try
			{
				ViewSource(A);
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
				return;
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				View.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
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
				return;
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

	public static void ViewSource(Microsoft.Office.Interop.Word.Shape shp)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		if (Common.IsLinked(shp))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A(Common.LinkDetails(shp), shp);
					return;
				}
			}
		}
		throw new Exception(XC.A(14498));
	}

	public static void ViewSource(InlineShape shp)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		if (Common.IsLinked(shp))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A(Common.LinkDetails(shp), shp);
					return;
				}
			}
		}
		throw new Exception(XC.A(14498));
	}

	public static void ViewSource(Table shp)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		if (Common.IsLinked(shp))
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
					A(Common.LinkDetails(shp), shp);
					return;
				}
			}
		}
		throw new Exception(XC.A(14498));
	}

	public static void ViewSource(ContentControl cc)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		if (Common.IsLinked(cc))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A(Common.LinkDetails(cc), cc);
					return;
				}
			}
		}
		throw new Exception(XC.A(14547));
	}

	private static void A(Link A, object B)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		if (Operators.CompareString(A.Name, string.Empty, TextCompare: false) == 0)
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
					throw new Exception(XC.A(14616));
				}
			}
		}
		string text = Source.View(A);
		if (Operators.CompareString(text, A.Source, TextCompare: false) == 0)
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
			Type typeFromHandle = typeof(Update);
			string memberName = XC.A(13872);
			object[] obj = new object[4] { B, null, text, true };
			object[] array = obj;
			bool[] obj2 = new bool[4] { true, false, true, false };
			bool[] array2 = obj2;
			NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
			if (array2[0])
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
				B = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (!array2[2])
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
				text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(string));
				return;
			}
		}
	}
}
