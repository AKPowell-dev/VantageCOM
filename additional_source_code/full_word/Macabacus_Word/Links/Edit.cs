using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Edit
{
	public struct EditedShapes
	{
		public List<object> Objects;

		public List<bool> IsError;

		public List<string> Errors;

		public void ClearReferences()
		{
			ReleaseHelper.ClearListReferences<object>(ref Objects, false, (Action<object>)null);
		}
	}

	public static EditedShapes EditLink(List<object> listShapes)
	{
		EditedShapes result = new EditedShapes
		{
			Objects = listShapes,
			IsError = null,
			Errors = null
		};
		wpfLinkEdit wpfLinkEdit2;
		wpfLinkEdit obj = (wpfLinkEdit2 = new wpfLinkEdit(listShapes));
		if (Properties.EditLinksWidth > 0.0)
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
			wpfLinkEdit2.Width = Properties.EditLinksWidth;
		}
		Base.ShowDialogNotTopmost((System.Windows.Window)obj);
		if (wpfLinkEdit2.DialogResult == true)
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
			result = wpfLinkEdit2.ReturnValue;
		}
		wpfLinkEdit2 = null;
		GC.Collect();
		if (result.Errors != null)
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
			int count = result.Errors.Count;
			if (count > 0)
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
				int count2 = listShapes.Count;
				if (count2 == 1)
				{
					Forms.ErrorMessage(XC.A(13099) + result.Errors[0]);
				}
				else if (result.Errors.Distinct().ToList().Count == 1)
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
					Forms.ErrorMessage(count + XC.A(13138) + count2 + XC.A(13147) + result.Errors[0]);
				}
				else
				{
					Forms.ErrorMessage(count + XC.A(13138) + count2 + XC.A(13190));
				}
			}
		}
		return result;
	}

	public static void EditLink()
	{
		EditedShapes editedShapes = default(EditedShapes);
		List<object> list;
		try
		{
			IEnumerable<wpfManageLinks> source = System.Windows.Application.Current.Windows.OfType<wpfManageLinks>();
			if (source.Any())
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
					wpfManageLinks obj = source.ElementAt(0);
					obj.Topmost = false;
					Forms.WarningMessage(XC.A(13229));
					obj.Topmost = true;
					list = null;
					return;
				}
			}
			source = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Selection selection = PC.A.Application.ActiveWindow.Selection;
		list = new List<object>();
		WdSelectionType type = selection.Type;
		if (type != WdSelectionType.wdSelectionInlineShape)
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
				if (Common.IsContentControlSelected(selection))
				{
					using (List<ContentControl>.Enumerator enumerator = Common.LinkedContentControlsInSelection(selection).GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							ContentControl current = enumerator.Current;
							if (Common.IsLinked(current))
							{
								list.Add(current);
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0295;
							}
							continue;
							end_IL_0295:
							break;
						}
					}
					if (list.Count > 0)
					{
						editedShapes = EditLink(list);
					}
					else
					{
						Forms.WarningMessage(XC.A(13338));
					}
				}
				else if (Common.IsTableSelected(selection))
				{
					{
						IEnumerator enumerator2 = selection.Tables.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Table table = (Table)enumerator2.Current;
								if (Common.IsLinked(table))
								{
									list.Add(table);
								}
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0324;
								}
								continue;
								end_IL_0324:
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
					if (list.Count > 0)
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
						editedShapes = EditLink(list);
					}
					else
					{
						Forms.WarningMessage(XC.A(13338));
					}
				}
			}
			else
			{
				IEnumerator enumerator3 = default(IEnumerator);
				try
				{
					enumerator3 = Helpers.SelectedShapes(selection).GetEnumerator();
					while (enumerator3.MoveNext())
					{
						A((Microsoft.Office.Interop.Word.Shape)enumerator3.Current, ref list);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_01f0;
						}
						continue;
						end_IL_01f0:
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (list.Count > 0)
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
					editedShapes = EditLink(list);
				}
				else
				{
					Forms.WarningMessage(XC.A(13338));
				}
			}
		}
		else
		{
			{
				IEnumerator enumerator4 = selection.InlineShapes.GetEnumerator();
				try
				{
					while (enumerator4.MoveNext())
					{
						InlineShape inlineShape = (InlineShape)enumerator4.Current;
						if (Common.IsLinked(inlineShape))
						{
							list.Add(inlineShape);
						}
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0107;
						}
						continue;
						end_IL_0107:
						break;
					}
				}
				finally
				{
					IDisposable disposable2 = enumerator4 as IDisposable;
					if (disposable2 != null)
					{
						disposable2.Dispose();
					}
				}
			}
			IEnumerator enumerator5 = default(IEnumerator);
			try
			{
				enumerator5 = selection.ChildShapeRange.GetEnumerator();
				while (enumerator5.MoveNext())
				{
					A((Microsoft.Office.Interop.Word.Shape)enumerator5.Current, ref list);
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_015b;
					}
					continue;
					end_IL_015b:
					break;
				}
			}
			finally
			{
				if (enumerator5 is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator5 as IDisposable).Dispose();
						break;
					}
				}
			}
			if (list.Count > 0)
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
				editedShapes = EditLink(list);
			}
			else
			{
				Forms.WarningMessage(XC.A(13338));
			}
		}
		editedShapes.ClearReferences();
		ReleaseHelper.ReleaseObjectList<object>(ref list, false);
		selection = null;
		ReleaseHelper.DoGarbageCollection();
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, ref List<object> B)
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
					if (Common.IsLinked(A))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								B.Add(A);
								return;
							}
						}
					}
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Edit.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, ref B);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
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
