using System;
using System.Collections;
using System.Collections.Generic;
using A;
using MacabacusMacros.Links;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Break
{
	public static void SelectedLinks()
	{
		if (!Base.ConfirmBreakLink())
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
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
			Selection selection = PC.A.Application.ActiveWindow.Selection;
			WdSelectionType type = selection.Type;
			if (type != WdSelectionType.wdSelectionInlineShape)
			{
				if (type != WdSelectionType.wdSelectionShape)
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
					if (Common.IsContentControlSelected(selection))
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
						using List<ContentControl>.Enumerator enumerator = Common.LinkedContentControlsInSelection(selection).GetEnumerator();
						while (enumerator.MoveNext())
						{
							BreakLink(enumerator.Current);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_01b0;
							}
							continue;
							end_IL_01b0:
							break;
						}
					}
					else if (Common.IsTableSelected(selection))
					{
						{
							enumerator2 = selection.Tables.GetEnumerator();
							try
							{
								while (enumerator2.MoveNext())
								{
									Table shp = (Table)enumerator2.Current;
									try
									{
										BreakLink(shp);
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0217;
									}
									continue;
									end_IL_0217:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator2 as IDisposable;
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
					try
					{
						enumerator3 = Helpers.SelectedShapes(selection).GetEnumerator();
						while (enumerator3.MoveNext())
						{
							A((Microsoft.Office.Interop.Word.Shape)enumerator3.Current);
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			else
			{
				try
				{
					enumerator4 = selection.InlineShapes.GetEnumerator();
					while (enumerator4.MoveNext())
					{
						InlineShape shp2 = (InlineShape)enumerator4.Current;
						try
						{
							BreakLink(shp2);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
					}
				}
				finally
				{
					if (enumerator4 is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator4 as IDisposable).Dispose();
							break;
						}
					}
				}
				try
				{
					enumerator5 = selection.ChildShapeRange.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						A((Microsoft.Office.Interop.Word.Shape)enumerator5.Current);
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_00e9;
						}
						continue;
						end_IL_00e9:
						break;
					}
				}
				finally
				{
					if (enumerator5 is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator5 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			selection = null;
			clsRibbon.InvalidateLinkedItemControls();
			Common.LogActivity(XC.A(13035));
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					try
					{
						BreakLink(A);
						return;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
						return;
					}
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Break.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
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

	public static void BreakLink(Microsoft.Office.Interop.Word.Shape shp)
	{
		try
		{
			if (!Common.IsLinked(shp))
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
				Macabacus_Word.CustomXML.RemoveCustomXMLPart(shp);
				shp.AlternativeText = "";
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new Exception(XC.A(13056));
		}
	}

	public static void BreakLink(InlineShape shp)
	{
		try
		{
			if (!Common.IsLinked(shp))
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
				Macabacus_Word.CustomXML.RemoveCustomXMLPart(shp);
				shp.AlternativeText = "";
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new Exception(XC.A(13056));
		}
	}

	public static void BreakLink(Table shp)
	{
		try
		{
			if (!Common.IsLinked(shp))
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
				Macabacus_Word.CustomXML.RemoveCustomXMLPart(shp);
				shp.Descr = "";
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new Exception(XC.A(13056));
		}
	}

	public static void BreakLink(ContentControl cc)
	{
		try
		{
			if (Common.IsLinked(cc))
			{
				Macabacus_Word.CustomXML.RemoveCustomXMLPart(cc);
				cc.Delete();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			throw new Exception(XC.A(13056));
		}
	}
}
