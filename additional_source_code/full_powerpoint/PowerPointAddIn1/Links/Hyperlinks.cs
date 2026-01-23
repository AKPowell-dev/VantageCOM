using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class Hyperlinks
{
	[CompilerGenerated]
	internal sealed class AF
	{
		public int A;

		[SpecialName]
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Id == this.A;
		}
	}

	public static Link LinkDetails(Hyperlink hyp)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0144: Unknown result type (might be due to invalid IL or missing references)
		//IL_0146: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fd: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Link val = default(Link);
		XmlNode nd = default(XmlNode);
		XmlDocument xmlDocument = default(XmlDocument);
		IEnumerator enumerator = default(IEnumerator);
		Link result = default(Link);
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
				case 428:
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
							goto IL_0011;
						case 4:
							goto IL_001a;
						case 5:
							goto IL_002b;
						case 6:
							goto IL_005e;
						case 7:
							goto IL_0060;
						case 8:
							goto IL_0072;
						case 9:
							goto IL_0082;
						case 10:
							goto IL_0095;
						case 11:
							goto IL_00a8;
						case 12:
							goto IL_00c5;
						case 13:
							goto IL_00d6;
						case 14:
							goto IL_00e9;
						case 15:
						case 16:
							goto IL_00fa;
						case 17:
							goto IL_011c;
						case 18:
							goto IL_013e;
						case 19:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 20:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0082:
					num2 = 9;
					val.Name = Common.GetLinkId(nd);
					goto IL_0095;
					IL_0007:
					num2 = 2;
					val = default(Link);
					goto IL_0011;
					IL_0011:
					num2 = 3;
					xmlDocument = new XmlDocument();
					goto IL_001a;
					IL_001a:
					num2 = 4;
					xmlDocument.LoadXml(hyp.SubAddress);
					goto IL_002b;
					IL_002b:
					num2 = 5;
					enumerator = xmlDocument.SelectNodes(AH.A(94272)).GetEnumerator();
					goto IL_00fd;
					IL_00fd:
					if (enumerator.MoveNext())
					{
						nd = (XmlNode)enumerator.Current;
						goto IL_005e;
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
					goto IL_011c;
					IL_0095:
					num2 = 10;
					val.ParentId = Common.GetParentId(nd);
					goto IL_00a8;
					IL_00c5:
					num2 = 12;
					val.LastUpdate = Common.GetLinkTime(nd);
					goto IL_00d6;
					IL_00d6:
					num2 = 13;
					val.LastUser = Common.GetLinkUser(nd);
					goto IL_00e9;
					IL_011c:
					num2 = 17;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_013e;
					IL_00e9:
					num2 = 14;
					val.Address = Common.GetLinkAddress(nd);
					goto IL_00fa;
					IL_00fa:
					num2 = 16;
					goto IL_00fd;
					IL_013e:
					xmlDocument = null;
					break;
					IL_00a8:
					num2 = 11;
					val.Type = (ImportType)Conversions.ToInteger(Common.GetLinkOther(nd, Base.XML_NODE_TYPE));
					goto IL_00c5;
					IL_005e:
					num2 = 6;
					goto IL_0060;
					IL_0060:
					num2 = 7;
					val.Source = Common.GetLinkSource(nd);
					goto IL_0072;
					IL_0072:
					num2 = 8;
					val.SourceModified = Common.GetLinkSourceModified(nd);
					goto IL_0082;
					end_IL_0000_2:
					break;
				}
				num2 = 19;
				result = val;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 428;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsLinked(Hyperlink hyp)
	{
		bool result;
		try
		{
			int num;
			if (hyp.Type == MsoHyperlinkType.msoHyperlinkRange && Operators.CompareString(hyp.Address, Base.HYPERLINK_ADDRESS, TextCompare: false) == 0)
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
				num = (hyp.SubAddress.Contains(AH.A(94285)) ? 1 : 0);
			}
			else
			{
				num = 0;
			}
			result = (byte)num != 0;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void UpdateSource(Hyperlink hyp, RefreshInstance refreshInstance, string strFullName, bool blnUpdateLastModified)
	{
		B(hyp, Base.XML_NODE_SOURCE, CloudStorage.AddPlaceholdersToPath(strFullName));
		if (!blnUpdateLastModified)
		{
			return;
		}
		string lastModifiedTime;
		if (refreshInstance != null)
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
			lastModifiedTime = refreshInstance.GetLastModifiedTime(strFullName);
		}
		else
		{
			lastModifiedTime = Updates.GetLastModifiedTime(strFullName);
		}
		if (lastModifiedTime.Length <= 0)
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
			B(hyp, Base.XML_NODE_SOURCE_LAST_MOD, lastModifiedTime);
			return;
		}
	}

	private static void A(Hyperlink A)
	{
		B(A, Base.XML_NODE_TIME, Base.LastUpdate());
	}

	public static void UpdateUser(Hyperlink hyp, string strUser)
	{
		B(hyp, Base.XML_NODE_USER, strUser);
	}

	private static void A(Hyperlink A, string B)
	{
		Hyperlinks.B(A, Base.XML_NODE_ADDRESS, B);
	}

	public static void UpdateName(Hyperlink hyp, string strName)
	{
		B(hyp, Base.XML_NODE_LINK_ID, strName);
	}

	public static void UpdateParentId(Hyperlink hyp, string strParentId)
	{
		B(hyp, Base.XML_NODE_PARENT_ID, strParentId);
	}

	private static void B(Hyperlink A, string B, string C)
	{
		string subAddress = A.SubAddress;
		if (Operators.CompareString(subAddress, string.Empty, TextCompare: false) == 0)
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
			XmlDocument xmlDocument = Common.UpdateXml(subAddress, B, C);
			A.SubAddress = xmlDocument.OuterXml;
			subAddress = xmlDocument.OuterXml;
			xmlDocument = null;
			try
			{
				A.ScreenTip = AH.A(94302);
				return;
			}
			catch (COMException ex)
			{
				ProjectData.SetProjectError(ex);
				COMException ex2 = ex;
				TextRange textRange = HyperlinkParentTextRange(A);
				if (textRange.Text.EndsWith(AH.A(47331)))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							Hyperlinks.C(A);
							textRange = textRange.TrimText();
							Add.BuildHyperlink(textRange, subAddress);
							textRange = null;
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
				throw;
			}
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape GetParentShape(Hyperlink hyp, bool blnIgnoreTables)
	{
		return A(hyp, blnIgnoreTables);
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape GetParentShape(TextRange txtRng, bool blnIgnoreTables)
	{
		return A(txtRng, blnIgnoreTables);
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(object A, bool B)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(A, null, AH.A(28234), new object[0], null, null, null));
		while (!(objectValue is Microsoft.Office.Interop.PowerPoint.Shape))
		{
			objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, AH.A(28234), new object[0], null, null, null));
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
			if (!B)
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
				if (IsShapeInsideTableCell((Microsoft.Office.Interop.PowerPoint.Shape)objectValue))
				{
					return Hyperlinks.A((Microsoft.Office.Interop.PowerPoint.Shape)objectValue);
				}
			}
			return (Microsoft.Office.Interop.PowerPoint.Shape)objectValue;
		}
	}

	public static TextRange HyperlinkParentTextRange(Hyperlink hyp)
	{
		return (TextRange)NewLateBinding.LateGet(hyp.Parent, null, AH.A(28234), new object[0], null, null, null);
	}

	public static TextRange2 HyperlinkParentTextRange2(Hyperlink hyp)
	{
		Microsoft.Office.Interop.PowerPoint.Shape parentShape = GetParentShape(hyp, blnIgnoreTables: true);
		TextRange textRange = HyperlinkParentTextRange(hyp);
		TextRange2 result = parentShape.TextFrame2.TextRange.get_Characters(textRange.Start, textRange.Length);
		textRange = null;
		return result;
	}

	public static bool SelectedShapesContainHyperlink(Hyperlink hyp, List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes)
	{
		int A = GetParentShape(hyp, blnIgnoreTables: false).Id;
		return listShapes.Where([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Id == A).Count() > 0;
	}

	public static bool IsShapeInsideTableCell(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result;
		try
		{
			_ = shp.ZOrderPosition;
			result = false;
		}
		catch (NotImplementedException ex)
		{
			ProjectData.SetProjectError(ex);
			NotImplementedException ex2 = ex;
			result = true;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			Table table;
			try
			{
				enumerator = ((Slide)A.Parent).Shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (shape.HasTable != MsoTriState.msoTrue)
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
					table = shape.Table;
					int count = table.Rows.Count;
					for (int i = 1; i <= count; i++)
					{
						int count2 = table.Columns.Count;
						for (int j = 1; j <= count2; j++)
						{
							if (table.Cell(i, j).Shape != A)
							{
								continue;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								table = null;
								return shape;
							}
						}
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
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00d2;
					}
					continue;
					end_IL_00d2:
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
			table = null;
			return null;
		}
	}

	public static bool IsHyperlinkSelected(Hyperlink hyp, Selection sel)
	{
		TextRange textRange = (TextRange)((ActionSetting)hyp.Parent).Parent;
		TextRange textRange2 = sel.TextRange;
		int start = textRange.Start;
		int start2 = textRange2.Start;
		bool result;
		checked
		{
			if (start >= start2)
			{
				if (start <= start2 + textRange2.Length)
				{
					goto IL_006e;
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
			}
			if (start2 > start)
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
				if (start2 < start + textRange.Length)
				{
					goto IL_006e;
				}
			}
			result = false;
			goto IL_0076;
		}
		IL_0076:
		JG.A(textRange);
		JG.A(textRange2);
		return result;
		IL_006e:
		result = true;
		goto IL_0076;
	}

	public static List<Hyperlink> SelectedLinks(Selection sel)
	{
		List<Hyperlink> list = new List<Hyperlink>();
		Slide slide = null;
		List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes = null;
		List<Cell> list2 = null;
		checked
		{
			try
			{
				slide = sel.SlideRange[1];
				if (slide.Hyperlinks.Count > 0)
				{
					IEnumerator enumerator3 = default(IEnumerator);
					IEnumerator enumerator4 = default(IEnumerator);
					IEnumerator enumerator5 = default(IEnumerator);
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
						bool flag = sel.Type == PpSelectionType.ppSelectionText;
						listShapes = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
						list2 = new List<Cell>();
						if (sel.ShapeRange.Count == 1)
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
							if (sel.ShapeRange[1].HasTable == MsoTriState.msoTrue)
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
								Table table = sel.ShapeRange[1].Table;
								int count = table.Rows.Count;
								for (int i = 1; i <= count; i++)
								{
									int count2 = table.Columns.Count;
									for (int j = 1; j <= count2; j++)
									{
										if (!table.Cell(i, j).Selected)
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
										list2.Add(table.Cell(i, j));
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_0129;
										}
										continue;
										end_IL_0129:
										break;
									}
								}
								table = null;
							}
						}
						if (list2.Count > 0)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								foreach (Hyperlink hyperlink3 in slide.Hyperlinks)
								{
									if (!IsLinked(hyperlink3))
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
									Microsoft.Office.Interop.PowerPoint.Shape parentShape = GetParentShape(hyperlink3, blnIgnoreTables: true);
									using (List<Cell>.Enumerator enumerator2 = list2.GetEnumerator())
									{
										while (enumerator2.MoveNext())
										{
											if (enumerator2.Current.Shape != parentShape)
											{
												continue;
											}
											if (flag)
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
												if (!IsHyperlinkSelected(hyperlink3, sel))
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
												list.Add(hyperlink3);
											}
											else
											{
												list.Add(hyperlink3);
											}
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_01ff;
											}
											continue;
											end_IL_01ff:
											break;
										}
									}
									parentShape = null;
								}
								break;
							}
							break;
						}
						if (sel.HasChildShapeRange)
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
							try
							{
								enumerator3 = sel.ChildShapeRange.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									ProcessSelectedShape((Microsoft.Office.Interop.PowerPoint.Shape)RuntimeHelpers.GetObjectValue(enumerator3.Current), ref listShapes);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0291;
									}
									continue;
									end_IL_0291:
									break;
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (2)
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
						else
						{
							try
							{
								enumerator4 = sel.ShapeRange.GetEnumerator();
								while (enumerator4.MoveNext())
								{
									ProcessSelectedShape((Microsoft.Office.Interop.PowerPoint.Shape)RuntimeHelpers.GetObjectValue(enumerator4.Current), ref listShapes);
								}
							}
							finally
							{
								if (enumerator4 is IDisposable)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										(enumerator4 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						enumerator5 = slide.Hyperlinks.GetEnumerator();
						try
						{
							while (enumerator5.MoveNext())
							{
								Hyperlink hyperlink2 = (Hyperlink)enumerator5.Current;
								if (!IsLinked(hyperlink2))
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
								if (!SelectedShapesContainHyperlink(hyperlink2, listShapes))
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
								if (flag)
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
									if (!IsHyperlinkSelected(hyperlink2, sel))
									{
										continue;
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
									list.Add(hyperlink2);
								}
								else
								{
									list.Add(hyperlink2);
								}
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_03a5;
								}
								continue;
								end_IL_03a5:
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
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			ReleaseHelper.FullReleaseComObj((object)slide);
			slide = null;
			ReleaseHelper.ClearListReferences<Microsoft.Office.Interop.PowerPoint.Shape>(ref listShapes, false, (Action<Microsoft.Office.Interop.PowerPoint.Shape>)null);
			ReleaseHelper.ClearListReferences<Cell>(ref list2, false, (Action<Cell>)null);
			ReleaseHelper.DoGarbageCollection();
			return list;
		}
	}

	public static void ProcessSelectedShape(Microsoft.Office.Interop.PowerPoint.Shape shp, ref List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes)
	{
		if (shp.Type != MsoShapeType.msoGroup)
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
					listShapes.Add(shp);
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shp.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ProcessSelectedShape((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref listShapes);
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

	public static void RemoveHyperlinks()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		int num = 0;
		checked
		{
			try
			{
				if (application.Presentations.Count > 0 && MessageBox.Show(AH.A(94331), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						application.StartNewUndoEntry();
						foreach (Slide slide in application.ActivePresentation.Slides)
						{
							for (int i = slide.Hyperlinks.Count; i >= 1; i += -1)
							{
								if (IsLinked(slide.Hyperlinks[i]))
								{
									C(slide.Hyperlinks[i]);
									num++;
								}
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
						}
						if (num > 0)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								Forms.SuccessMessage(AH.A(94493) + num + AH.A(94510));
								break;
							}
						}
						else
						{
							Forms.InfoMessage(AH.A(94605));
						}
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
			application = null;
		}
	}

	public static Hyperlink Refresh(Hyperlink hyp, bool blnAll, ref List<string> listUpdatedShapeNames, ref RefreshInstance refreshInstance)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		if (blnAll)
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
			Common.NavigateToSlide(hyp, NG.A.Application);
		}
		HyperlinkParentTextRange(hyp).Select();
		A(ref hyp, ref refreshInstance, LinkDetails(hyp));
		listUpdatedShapeNames.Add(GetParentShape(hyp, blnIgnoreTables: false).Name);
		return hyp;
	}

	private static void A(ref Hyperlink A, ref RefreshInstance B, Link C)
	{
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_0163: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		Range range = null;
		bool flag = false;
		bool flag2 = false;
		Name name = null;
		Worksheet worksheet = null;
		string text = "";
		RefreshInstance obj = B;
		Microsoft.Office.Interop.Excel.Chart chart = null;
		obj.LocateSource(ref workbook, ref range, ref chart, ref C, ref flag, ref flag2, true);
		if (flag2)
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
			UpdateSource(A, B, C.Source, blnUpdateLastModified: false);
		}
		if (flag)
		{
			return;
		}
		XlSheetVisibility xlSheetVisibility = default(XlSheetVisibility);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (range != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
					{
						B.SourceRange(LinkDetails(A), workbook, ref name, ref worksheet, ref range, ref text, ref xlSheetVisibility);
						if (range.Cells.Count > 1)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									throw new UpdateLinkException(AH.A(94708));
								}
							}
						}
						TextRange textRange = HyperlinkParentTextRange(A);
						textRange.Text = range.Text.ToString().Trim();
						Add.CreateHyperlink(HyperlinkParentTextRange2(A), textRange, workbook, range, name, C.ParentId, B.HyperlinkColors);
						if (worksheet != null)
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
							if (worksheet.Visible != xlSheetVisibility)
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
								worksheet.Visible = xlSheetVisibility;
							}
						}
						B.RestoreExcel();
						RefreshInstance val;
						(val = B).RefreshedLinkCount = checked(val.RefreshedLinkCount + 1);
						range = null;
						worksheet = null;
						name = null;
						workbook = null;
						return;
					}
					}
				}
			}
			workbook = null;
			throw new UpdateLinkException(AH.A(94838));
		}
	}

	public static void ViewSource(Selection sel, ref int intViewed)
	{
		checked
		{
			using List<Hyperlink>.Enumerator enumerator = SelectedLinks(sel).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Hyperlink current = enumerator.Current;
				try
				{
					ViewSource(current);
					intViewed++;
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

	public static void ViewSource(Hyperlink hyp)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		if (IsLinked(hyp))
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
					Link val = LinkDetails(hyp);
					string text = Source.View(val);
					if (Operators.CompareString(text, val.Source, TextCompare: false) != 0)
					{
						try
						{
							UpdateSource(hyp, null, text, blnUpdateLastModified: true);
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
					return;
				}
				}
			}
		}
		throw new Exception(AH.A(94939));
	}

	public static void BreakLinks(Selection sel)
	{
		foreach (Hyperlink item in SelectedLinks(sel))
		{
			BreakLink(item);
		}
	}

	public static void BreakLink(Hyperlink hyp)
	{
		if (!IsLinked(hyp))
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
			C(hyp);
			return;
		}
	}

	public static bool PromptToConvert()
	{
		return Forms.YesNoMessage(AH.A(94986)) == DialogResult.Yes;
	}

	public static void ConvertFromLegacyLinks(ref List<TextLink> listTextLinks)
	{
		using (List<TextLink>.Enumerator enumerator = listTextLinks.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				A(enumerator.Current);
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
				break;
			}
		}
		listTextLinks.Clear();
	}

	private static void A(TextLink A)
	{
		//IL_0129: Unknown result type (might be due to invalid IL or missing references)
		//IL_0134: Expected O, but got Unknown
		string text = A.Xml;
		string text2 = "";
		TextLink val = A;
		TextRange textRange = val.Shape.TextFrame.TextRange.Characters(val.TextRange.Start, val.TextRange.Length);
		val = null;
		try
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(text);
			string[] array = new string[4]
			{
				Text.XML_NODE_TEXT,
				Text.XML_NODE_VALUE,
				Text.XML_NODE_START,
				Text.XML_NODE_RANGE_ID
			};
			foreach (string text3 in array)
			{
				XmlNode xmlNode = xmlDocument.DocumentElement.SelectSingleNode(text3);
				if (Operators.CompareString(text3, Text.XML_NODE_RANGE_ID, TextCompare: false) == 0)
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
					text2 = xmlNode.InnerText;
				}
				xmlNode.ParentNode.RemoveChild(xmlNode);
				xmlNode = null;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				text = xmlDocument.OuterXml;
				xmlDocument = null;
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Add.BuildHyperlink(textRange, text);
		Add.FormatHyperlink(A.TextRange, new HyperlinkColors(), false);
		XmlNodeList xmlNodeList;
		if (text2.Length > 0)
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
			XmlDocument xmlDocument2 = new XmlDocument();
			xmlDocument2.LoadXml(A.Shape.Tags[Text.TAG_TEXT_LINK_XML]);
			XmlNode xmlNode = xmlDocument2.SelectSingleNode(Text.XpathNodeById(text2));
			xmlNode.ParentNode.RemoveChild(xmlNode);
			xmlNode = null;
			xmlNodeList = xmlDocument2.SelectNodes(Text.XpathQuery());
			if (xmlNodeList != null && xmlNodeList.Count > 0)
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
				Text.SetXml(A.Shape, xmlDocument2.OuterXml);
			}
			else
			{
				Text.BreakLink(A.Shape);
			}
			xmlDocument2 = null;
		}
		textRange = null;
		xmlNodeList = null;
	}

	public static void SelectionChange(Selection Sel)
	{
		TextRange textRange = null;
		checked
		{
			try
			{
				Selection selection = Sel;
				if (selection.Application.ActiveWindow.ActivePane.ViewType == PpViewType.ppViewSlide)
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
					if (selection.Type == PpSelectionType.ppSelectionText)
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
						if (selection.TextRange2.Length == 0)
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
							Hyperlink hyperlink = selection.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink;
							if (hyperlink != null)
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
								if (IsLinked(hyperlink))
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
										textRange = ((Microsoft.Office.Interop.PowerPoint.ShapeRange)NewLateBinding.LateGet(selection.TextRange2.Parent, null, AH.A(28234), new object[0], null, null, null))[1].TextFrame.TextRange;
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										try
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape = selection.ShapeRange[1];
											if (shape.HasTable == MsoTriState.msoTrue)
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
												Table table = shape.Table;
												int count = table.Rows.Count;
												int num = 1;
												while (true)
												{
													if (num <= count)
													{
														int count2 = table.Columns.Count;
														int num2 = 1;
														while (true)
														{
															if (num2 <= count2)
															{
																Cell cell = table.Cell(num, num2);
																if (cell.Selected)
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
																	textRange = cell.Shape.TextFrame.TextRange;
																	break;
																}
																cell = null;
																num2++;
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
															break;
														}
														if (textRange != null)
														{
															break;
														}
														while (true)
														{
															switch (5)
															{
															case 0:
																break;
															default:
																goto end_IL_01cd;
															}
															continue;
															end_IL_01cd:
															break;
														}
														num++;
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
													break;
												}
												table = null;
											}
											shape = null;
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										ProjectData.ClearProjectError();
									}
									if (textRange != null)
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
										int start = selection.TextRange2.Start;
										if (start > textRange.Characters().Count || start == 1)
										{
											return;
										}
										int num3 = start;
										int num4 = start - 1;
										while (true)
										{
											if (num4 >= 1)
											{
												try
												{
													if (A(textRange, num4))
													{
														num3--;
														goto IL_027e;
													}
												}
												catch (Exception ex5)
												{
													ProjectData.SetProjectError(ex5);
													Exception ex6 = ex5;
													ProjectData.ClearProjectError();
												}
												break;
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
											break;
											IL_027e:
											num4 += -1;
										}
										int num5 = num3;
										if (start == num5)
										{
											return;
										}
										num3 = start;
										int count3 = textRange.Characters().Count;
										for (int i = start; i <= count3; i++)
										{
											try
											{
												if (!A(textRange, i))
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
													num3++;
													break;
												}
												continue;
											}
											catch (Exception ex7)
											{
												ProjectData.SetProjectError(ex7);
												Exception ex8 = ex7;
												ProjectData.ClearProjectError();
											}
											break;
										}
										int length = num3 - num5;
										new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(selection.Application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
										try
										{
											textRange.Characters(num5, length).Select();
										}
										catch (Exception ex9)
										{
											ProjectData.SetProjectError(ex9);
											Exception ex10 = ex9;
											ProjectData.ClearProjectError();
										}
										new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(selection.Application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
										textRange = null;
									}
								}
							}
							hyperlink = null;
						}
					}
				}
				selection = null;
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			finally
			{
				Hyperlink hyperlink = null;
			}
		}
	}

	private static bool A(TextRange A, int B)
	{
		return IsLinked(A.Characters(B, 0).ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink);
	}

	internal static void C(Hyperlink A)
	{
		try
		{
			bool flag = false;
			try
			{
				object obj = HyperlinkParentTextRange(A).Text;
				if (obj == null)
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
					obj = "";
				}
				string text = (string)obj;
				string[] source = new string[2]
				{
					AH.A(47331),
					AH.A(47334)
				};
				int num;
				if (Operators.CompareString(text, "", TextCompare: false) != 0)
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
					if (!source.Contains(Conversions.ToString(text[0])))
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
						num = (source.Contains(Conversions.ToString(text.Last())) ? 1 : 0);
					}
					else
					{
						num = 1;
					}
				}
				else
				{
					num = 0;
				}
				flag = (byte)num != 0;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			A.Delete();
			if (!flag)
			{
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
			checked
			{
				try
				{
					Microsoft.Office.Interop.PowerPoint.Shape parentShape = GetParentShape(A, blnIgnoreTables: true);
					Slide slideFromShape = clsPowerPoint.GetSlideFromShape(parentShape);
					TextRange textRange = parentShape.TextFrame.TextRange;
					int num2 = textRange.Runs().Count + 1;
					while (num2 > 1)
					{
						TextRange textRange2 = null;
						num2 = Math.Min(num2 - 1, textRange.Runs().Count);
						if (num2 == 0)
						{
							break;
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
						textRange2 = textRange.Runs(num2);
						string text2 = textRange2.Text ?? "";
						if (Operators.CompareString(text2, "", TextCompare: false) == 0)
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
						if (Operators.CompareString(text2.Replace(AH.A(47331), "").Replace(AH.A(47334), ""), "", TextCompare: false) != 0)
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
						if (!Hyperlinks.A(Hyperlinks.A(slideFromShape), textRange2))
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
						if (num2 == 1)
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
							textRange2.Delete();
							if (textRange.Runs().Count > 0)
							{
								textRange.Runs(1).InsertBefore(text2);
							}
							else
							{
								textRange.Text = text2;
							}
						}
						else if (num2 == textRange.Runs().Count)
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
							textRange2.Delete();
						}
						else
						{
							textRange.Runs(num2 - 1).InsertAfter(text2);
							textRange2.Delete();
						}
					}
				}
				catch (Exception projectError2)
				{
					ProjectData.SetProjectError(projectError2);
					ProjectData.ClearProjectError();
				}
				finally
				{
					TextRange textRange2 = null;
					TextRange textRange = null;
					Slide slideFromShape = null;
				}
			}
		}
		finally
		{
		}
	}

	private static List<TextRange> A(Slide A)
	{
		List<TextRange> list = new List<TextRange>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Hyperlinks.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Hyperlink hyp = (Hyperlink)enumerator.Current;
				list.Add(HyperlinkParentTextRange(hyp));
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
				break;
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
		return list;
	}

	private static bool A(List<TextRange> A, TextRange B)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			using (List<TextRange>.Enumerator enumerator = A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					TextRange current = enumerator.Current;
					try
					{
						object obj = current.Text;
						if (obj == null)
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
							obj = "";
						}
						object obj2 = B.Text;
						if (obj2 == null)
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
							obj2 = "";
						}
						if (Operators.CompareString((string)obj, (string)obj2, TextCompare: false) != 0)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0061;
								}
								continue;
								end_IL_0061:
								break;
							}
							continue;
						}
						if (current.Start != B.Start)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_007f;
								}
								continue;
								end_IL_007f:
								break;
							}
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.Shape parentShape = GetParentShape(current, blnIgnoreTables: true);
						if (shape == null)
						{
							shape = GetParentShape(B, blnIgnoreTables: true);
						}
						if (parentShape != shape)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_00a2;
								}
								continue;
								end_IL_00a2:
								break;
							}
							continue;
						}
						return true;
					}
					finally
					{
						current = null;
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00c3;
					}
					continue;
					end_IL_00c3:
					break;
				}
			}
			return false;
		}
		finally
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		}
	}
}
