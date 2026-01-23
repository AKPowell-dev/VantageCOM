using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Common
{
	private static readonly object m_A = RuntimeHelpers.GetObjectValue(new object());

	public static Link LinkDetails(Microsoft.Office.Interop.Word.Shape shp)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		return A(shp.Anchor.Document, shp.AlternativeText);
	}

	public static Link LinkDetails(InlineShape shp)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		return A(shp.Range.Document, shp.AlternativeText);
	}

	public static Link LinkDetails(Table shp)
	{
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		return A(shp.Range.Document, shp.Descr);
	}

	public static Link LinkDetails(ContentControl cc)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		return A(cc.Range.Document, cc.Tag);
	}

	private static Link A(Document A, string B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b9: Unknown result type (might be due to invalid IL or missing references)
		Link result = default(Link);
		try
		{
			XmlDocument linkXML = CustomXML.GetLinkXML(A, B);
			result.Source = CloudStorage.FillPlaceholdersInPath(linkXML.GetElementsByTagName(CustomXML.XML_NODE_SOURCE).Item(0).InnerText);
			try
			{
				result.SourceModified = linkXML.GetElementsByTagName(CustomXML.XML_NODE_SOURCE_LAST_MOD).Item(0).InnerText;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result.SourceModified = "";
				ProjectData.ClearProjectError();
			}
			result.Name = linkXML.GetElementsByTagName(CustomXML.XML_NODE_NAME).Item(0).InnerText;
			result.Type = (ImportType)Conversions.ToInteger(linkXML.GetElementsByTagName(CustomXML.XML_NODE_TYPE).Item(0).InnerText);
			result.LastUpdate = linkXML.GetElementsByTagName(CustomXML.XML_NODE_UPDATED).Item(0).InnerText;
			try
			{
				result.LastUser = linkXML.GetElementsByTagName(CustomXML.XML_NODE_USER).Item(0).InnerText;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				result.LastUser = "";
				ProjectData.ClearProjectError();
			}
			try
			{
				result.Address = linkXML.GetElementsByTagName(CustomXML.XML_NODE_ADDRESS).Item(0).InnerText;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				result.Address = "";
				ProjectData.ClearProjectError();
			}
			try
			{
				result.ParentId = linkXML.GetElementsByTagName(CustomXML.XML_NODE_PARENT).Item(0).InnerText;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				result.ParentId = "";
				ProjectData.ClearProjectError();
			}
			linkXML = null;
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsLinked(Microsoft.Office.Interop.Word.Shape shp)
	{
		return A(shp.AlternativeText);
	}

	public static bool IsLinked(InlineShape shp)
	{
		return A(shp.AlternativeText);
	}

	public static bool IsLinked(Table shp)
	{
		return A(shp.Descr);
	}

	public static bool IsLinked(ContentControl cc)
	{
		return A(cc.Tag);
	}

	private static bool A(string A)
	{
		if (string.IsNullOrEmpty(A))
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
					return false;
				}
			}
		}
		return A.Contains(CustomXML.ALT_TEXT_TAG);
	}

	public static bool IsTableSelected(Selection sel)
	{
		return Conversions.ToBoolean(sel.get_Information(WdInformation.wdWithInTable));
	}

	public static bool IsContentControlSelected(Selection sel)
	{
		return Conversions.ToBoolean(sel.get_Information(WdInformation.wdInContentControl));
	}

	public static List<ContentControl> LinkedContentControlsInSelection(Selection sel, bool firstOnly = false)
	{
		List<ContentControl> list = new List<ContentControl>();
		try
		{
			ContentControls contentControls;
			if (Conversions.ToBoolean(sel.get_Information(WdInformation.wdInHeaderFooter)))
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
				contentControls = sel.HeaderFooter.Range.ContentControls;
			}
			else
			{
				contentControls = sel.Document.ContentControls;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = contentControls.GetEnumerator();
				while (enumerator.MoveNext())
				{
					ContentControl contentControl = (ContentControl)enumerator.Current;
					if (!sel.InRange(contentControl.Range))
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
						if (!contentControl.Range.InRange(sel.Range))
						{
							continue;
						}
					}
					if (!IsLinked(contentControl))
					{
						continue;
					}
					list.Add(contentControl);
					if (firstOnly)
					{
						break;
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
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			ContentControls contentControls = null;
		}
		return list;
	}

	public static bool IsLinkSelected()
	{
		if (clsRibbon.IsLinkSelectedResult.HasValue)
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
					return clsRibbon.IsLinkSelectedResult.Value;
				}
			}
		}
		object a = Common.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			bool flag = false;
			Selection selection;
			try
			{
				selection = PC.A.Application.ActiveWindow.Selection;
				WdSelectionType type = selection.Type;
				if (type != WdSelectionType.wdSelectionInlineShape)
				{
					IEnumerator enumerator = default(IEnumerator);
					IEnumerator enumerator2 = default(IEnumerator);
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (type == WdSelectionType.wdSelectionShape)
						{
							try
							{
								enumerator = Helpers.SelectedShapes(selection).GetEnumerator();
								while (true)
								{
									if (enumerator.MoveNext())
									{
										flag = A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
										if (flag)
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
										continue;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_01b9;
										}
										continue;
										end_IL_01b9:
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
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator as IDisposable).Dispose();
										break;
									}
								}
							}
							break;
						}
						if (IsContentControlSelected(selection))
						{
							flag = LinkedContentControlsInSelection(selection, firstOnly: true).Any();
							break;
						}
						if (!IsTableSelected(selection))
						{
							break;
						}
						try
						{
							enumerator2 = selection.Tables.GetEnumerator();
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									if (IsLinked((Table)enumerator2.Current))
									{
										flag = true;
										break;
									}
									continue;
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										goto end_IL_024f;
									}
									continue;
									end_IL_024f:
									break;
								}
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						break;
					}
				}
				else
				{
					{
						IEnumerator enumerator3 = selection.InlineShapes.GetEnumerator();
						try
						{
							do
							{
								if (enumerator3.MoveNext())
								{
									flag = IsLinked((InlineShape)enumerator3.Current);
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_00d0;
									}
									continue;
									end_IL_00d0:
									break;
								}
								break;
							}
							while (!flag);
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
					if (!flag)
					{
						IEnumerator enumerator4 = default(IEnumerator);
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							try
							{
								enumerator4 = selection.ChildShapeRange.GetEnumerator();
								while (true)
								{
									if (enumerator4.MoveNext())
									{
										flag = A((Microsoft.Office.Interop.Word.Shape)enumerator4.Current);
										if (flag)
										{
											break;
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
										continue;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_0147;
										}
										continue;
										end_IL_0147:
										break;
									}
									break;
								}
							}
							finally
							{
								if (enumerator4 is IDisposable)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator4 as IDisposable).Dispose();
										break;
									}
								}
							}
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			selection = null;
			clsRibbon.IsLinkSelectedResult = flag;
			return flag;
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					Monitor.Exit(a);
					break;
				}
			}
		}
	}

	private static bool A(Microsoft.Office.Interop.Word.Shape A)
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
					return IsLinked(A);
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		bool result = default(bool);
		try
		{
			if (enumerator.MoveNext())
			{
				result = Common.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
				return result;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_0059;
				}
				continue;
				end_IL_0059:
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
		return result;
	}

	public static void CloseHeaderFooterView(Application wdApp)
	{
		try
		{
			Microsoft.Office.Interop.Word.View view = wdApp.ActiveWindow.View;
			WdSeekView seekView = view.SeekView;
			if ((uint)(seekView - 1) > 5u)
			{
				if ((uint)(seekView - 9) > 1u)
				{
					goto IL_003c;
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
			}
			view.SeekView = WdSeekView.wdSeekMainDocument;
			goto IL_003c;
			IL_003c:
			view = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void NavigateLink(object obj)
	{
		try
		{
			Navigate.A(RuntimeHelpers.GetObjectValue(obj));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)10, strActivity);
	}
}
