using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
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

public sealed class Text
{
	public static readonly string TAG_TEXT_LINK_XML = AH.A(95963);

	public static readonly string XML_NODE_TEXT = AH.A(70464);

	public static readonly string XML_NODE_VALUE = AH.A(93748);

	public static readonly string XML_NODE_START = AH.A(95998);

	public static readonly string XML_NODE_RANGE_ID = AH.A(96009);

	private static Microsoft.Office.Interop.PowerPoint.Shape m_A = null;

	private static Microsoft.Office.Interop.PowerPoint.Shape A
	{
		get
		{
			return Text.m_A;
		}
		set
		{
			Text.m_A = value;
		}
	}

	public static List<Link> LinkDetails(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_0145: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Link item = default(Link);
		XmlNode nd = default(XmlNode);
		List<Link> list = default(List<Link>);
		string text = default(string);
		XmlDocument xmlDocument = default(XmlDocument);
		XmlNodeList xmlNodeList = default(XmlNodeList);
		IEnumerator enumerator = default(IEnumerator);
		List<Link> result = default(List<Link>);
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
				case 538:
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
							goto IL_0010;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_003e;
						case 6:
							goto IL_0047;
						case 7:
							goto IL_0052;
						case 8:
							goto IL_0062;
						case 9:
							goto IL_006b;
						case 10:
							goto IL_008c;
						case 11:
							goto IL_0097;
						case 12:
							goto IL_009a;
						case 13:
							goto IL_00ad;
						case 14:
							goto IL_00c0;
						case 15:
							goto IL_00d3;
						case 16:
							goto IL_00e6;
						case 17:
							goto IL_00f1;
						case 18:
							goto IL_0102;
						case 19:
							goto IL_0115;
						case 20:
							goto IL_0128;
						case 21:
						case 22:
							goto IL_0140;
						case 23:
							goto IL_014c;
						case 24:
							goto IL_016e;
						case 25:
							goto IL_0186;
						case 26:
							goto IL_018c;
						case 27:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 28:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0102:
					num2 = 18;
					item.LastUser = Common.GetLinkUser(nd);
					goto IL_0115;
					IL_0007:
					num2 = 2;
					list = new List<Link>();
					goto IL_0010;
					IL_0010:
					num2 = 3;
					text = shp.Tags[TAG_TEXT_LINK_XML];
					goto IL_0028;
					IL_0028:
					num2 = 4;
					if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
					{
						break;
					}
					goto IL_003e;
					IL_003e:
					num2 = 5;
					xmlDocument = new XmlDocument();
					goto IL_0047;
					IL_0047:
					num2 = 6;
					xmlDocument.LoadXml(text);
					goto IL_0052;
					IL_0052:
					num2 = 7;
					xmlNodeList = xmlDocument.SelectNodes(XpathQuery());
					goto IL_0062;
					IL_0062:
					num2 = 8;
					if (xmlNodeList != null)
					{
						goto IL_006b;
					}
					goto IL_018c;
					IL_006b:
					num2 = 9;
					enumerator = xmlNodeList.GetEnumerator();
					goto IL_014f;
					IL_014f:
					if (enumerator.MoveNext())
					{
						nd = (XmlNode)enumerator.Current;
						goto IL_008c;
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
					goto IL_016e;
					IL_0115:
					num2 = 19;
					item.Address = Common.GetLinkAddress(nd);
					goto IL_0128;
					IL_0140:
					num2 = 22;
					list.Add(item);
					goto IL_014c;
					IL_014c:
					num2 = 23;
					goto IL_014f;
					IL_016e:
					num2 = 24;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0186;
					IL_018c:
					xmlDocument = null;
					break;
					IL_0186:
					num2 = 25;
					xmlNodeList = null;
					goto IL_018c;
					IL_0128:
					num2 = 20;
					item.RangeId = Common.GetLinkOther(nd, XML_NODE_RANGE_ID);
					goto IL_0140;
					IL_008c:
					num2 = 10;
					item = default(Link);
					goto IL_0097;
					IL_0097:
					num2 = 11;
					goto IL_009a;
					IL_009a:
					num2 = 12;
					item.Source = Common.GetLinkSource(nd);
					goto IL_00ad;
					IL_00ad:
					num2 = 13;
					item.SourceModified = Common.GetLinkSourceModified(nd);
					goto IL_00c0;
					IL_00c0:
					num2 = 14;
					item.Name = Common.GetLinkId(nd);
					goto IL_00d3;
					IL_00d3:
					num2 = 15;
					item.ParentId = Common.GetParentId(nd);
					goto IL_00e6;
					IL_00e6:
					num2 = 16;
					item.Type = (ImportType)4;
					goto IL_00f1;
					IL_00f1:
					num2 = 17;
					item.LastUpdate = Common.GetLinkTime(nd);
					goto IL_0102;
					end_IL_0000_2:
					break;
				}
				num2 = 27;
				result = list;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 538;
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
				switch (5)
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

	public static Link LinkDetails(TextLink textLink)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b1: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		XmlDocument xmlDocument = default(XmlDocument);
		Link val = default(Link);
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
				case 343:
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
							goto IL_0029;
						case 6:
							goto IL_0040;
						case 7:
							goto IL_0055;
						case 8:
							goto IL_006c;
						case 9:
							goto IL_0085;
						case 10:
							goto IL_0090;
						case 11:
							goto IL_00a8;
						case 12:
							goto IL_00c0;
						case 13:
							goto IL_00da;
						case 14:
							goto IL_00f9;
						case 15:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 16:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00f9:
					xmlDocument = null;
					break;
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
					xmlDocument.LoadXml(textLink.Xml);
					goto IL_0029;
					IL_0029:
					num2 = 5;
					val.Source = Common.GetLinkSource(xmlDocument.DocumentElement);
					goto IL_0040;
					IL_0040:
					num2 = 6;
					val.SourceModified = Common.GetLinkSourceModified(xmlDocument.DocumentElement);
					goto IL_0055;
					IL_0055:
					num2 = 7;
					val.Name = Common.GetLinkId(xmlDocument.DocumentElement);
					goto IL_006c;
					IL_006c:
					num2 = 8;
					val.ParentId = Common.GetParentId(xmlDocument.DocumentElement);
					goto IL_0085;
					IL_0085:
					num2 = 9;
					val.Type = (ImportType)4;
					goto IL_0090;
					IL_0090:
					num2 = 10;
					val.LastUpdate = Common.GetLinkTime(xmlDocument.DocumentElement);
					goto IL_00a8;
					IL_00a8:
					num2 = 11;
					val.LastUser = Common.GetLinkUser(xmlDocument.DocumentElement);
					goto IL_00c0;
					IL_00c0:
					num2 = 12;
					val.Address = Common.GetLinkAddress(xmlDocument.DocumentElement);
					goto IL_00da;
					IL_00da:
					num2 = 13;
					val.RangeId = Common.GetLinkOther(xmlDocument.DocumentElement, XML_NODE_RANGE_ID);
					goto IL_00f9;
					end_IL_0000_2:
					break;
				}
				num2 = 15;
				result = val;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 343;
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool ContainsLinks(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string text = default(string);
		bool result = default(bool);
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
				case 78:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0007:
					num2 = 2;
					text = shp.Tags[TAG_TEXT_LINK_XML];
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 3;
				result = text.Length > 0;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 78;
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static string A(XmlNode A)
	{
		return A.SelectSingleNode(XML_NODE_TEXT).InnerText;
	}

	private static int A(XmlNode A)
	{
		return Conversions.ToInteger(A.SelectSingleNode(XML_NODE_START).InnerText);
	}

	public static void UpdateSource(TextLink tl, string strRangeId, string strSource, bool blnUpdateLastModified)
	{
		A(tl, strRangeId, Base.XML_NODE_SOURCE, CloudStorage.AddPlaceholdersToPath(strSource));
		if (!blnUpdateLastModified)
		{
			return;
		}
		string lastModifiedTime = Updates.GetLastModifiedTime(strSource);
		if (lastModifiedTime.Length <= 0)
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
			A(tl, strRangeId, Base.XML_NODE_SOURCE_LAST_MOD, lastModifiedTime);
			return;
		}
	}

	public static void UpdateName(TextLink tl, string strRangeId, string strName)
	{
		A(tl, strRangeId, Base.XML_NODE_LINK_ID, strName);
	}

	public static void UpdateUser(TextLink tl, string strRangeId, string strUser)
	{
		A(tl, strRangeId, Base.XML_NODE_USER, strUser);
	}

	public static void UpdateParentId(TextLink tl, string strRangeId, string strParentId)
	{
		A(tl, strRangeId, Base.XML_NODE_PARENT_ID, strParentId);
	}

	private static void A(TextLink A, string B, string C, string D)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(A.Shape.Tags[TAG_TEXT_LINK_XML]);
		XmlNode xmlNode = xmlDocument.SelectSingleNode(XpathNodeById(B));
		if (xmlNode != null)
		{
			xmlNode.SelectSingleNode(C).InnerText = D;
			A.Xml = xmlNode.OuterXml;
			xmlNode = null;
		}
		SetXml(A.Shape, xmlDocument.OuterXml);
		xmlDocument = null;
	}

	private static void A(TextLink A, string B, Dictionary<string, string> C)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(A.Shape.Tags[TAG_TEXT_LINK_XML]);
		XmlNode xmlNode = xmlDocument.SelectSingleNode(XpathNodeById(B));
		if (xmlNode != null)
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
			using (Dictionary<string, string>.Enumerator enumerator = C.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					KeyValuePair<string, string> current = enumerator.Current;
					xmlNode.SelectSingleNode(current.Key).InnerText = current.Value;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_008c;
					}
					continue;
					end_IL_008c:
					break;
				}
			}
			A.Xml = xmlNode.OuterXml;
			xmlNode = null;
		}
		SetXml(A.Shape, xmlDocument.OuterXml);
		xmlDocument = null;
	}

	public static void SelectionChange(Selection Sel)
	{
		bool A = false;
		if (Text.A != null)
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			try
			{
				shape = Sel.ShapeRange[1];
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				shape = null;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (shape == null)
				{
					goto IL_0077;
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
				if (shape != Text.A)
				{
					goto IL_0077;
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
				if (Sel.Type != PpSelectionType.ppSelectionText)
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
					goto IL_0077;
				}
				goto end_IL_0037;
				IL_0077:
				if (Index.IndexedLinks != null)
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
						string text = Text.A(Text.A);
						if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								XmlDocument xmlDocument = new XmlDocument();
								xmlDocument.LoadXml(text);
								using (Dictionary<string, TextLink>.Enumerator enumerator = Index.IndexedLinks.GetEnumerator())
								{
									while (enumerator.MoveNext())
									{
										KeyValuePair<string, TextLink> current = enumerator.Current;
										XmlNode A2 = xmlDocument.SelectSingleNode(XpathNodeById(current.Key));
										if (A2 == null)
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
										TextRange2 textRange = current.Value.TextRange;
										try
										{
											if (textRange.Text.Length > 0)
											{
												while (true)
												{
													switch (4)
													{
													case 0:
														continue;
													}
													Text.A(ref A2, textRange);
													break;
												}
											}
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											A2.ParentNode.RemoveChild(A2);
											ProjectData.ClearProjectError();
										}
										A2 = null;
										textRange = null;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_016c;
										}
										continue;
										end_IL_016c:
										break;
									}
								}
								if (Operators.CompareString(text, xmlDocument.OuterXml, TextCompare: false) == 0)
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
									Text.A(ref A);
									SetXml(Text.A, xmlDocument.OuterXml);
									break;
								}
								break;
							}
						}
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					Index.IndexedLinks = null;
				}
				Text.A = null;
				end_IL_0037:;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
		}
		checked
		{
			try
			{
				if (Sel.Application.ActiveWindow.ActivePane.ViewType != PpViewType.ppViewSlide || Sel.Type != PpSelectionType.ppSelectionText)
				{
					return;
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape = Sel.ShapeRange[1];
				if (shape != Text.A)
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
					string text = Text.A(shape);
					if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
						if (shape.HasTextFrame == MsoTriState.msoTrue)
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
							XmlDocument xmlDocument = new XmlDocument();
							xmlDocument.LoadXml(text);
							XmlNodeList xmlNodeList = xmlDocument.SelectNodes(XpathQuery());
							if (xmlNodeList != null)
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
								Index.IndexedLinks = new Dictionary<string, TextLink>();
								XmlNode A3;
								for (int i = xmlNodeList.Count - 1; i >= 0; i += -1)
								{
									A3 = xmlNodeList[i];
									TextRange2 textRange2;
									try
									{
										string text2 = Text.A(A3);
										string innerText = A3.SelectSingleNode(XML_NODE_RANGE_ID).InnerText;
										textRange2 = shape.TextFrame2.TextRange.get_Characters(Text.A(A3), text2.Length);
										if (Operators.CompareString(textRange2.Text, text2, TextCompare: false) == 0)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												Index.IndexedLinks.Add(innerText, Text.A(shape, textRange2, A3));
												break;
											}
										}
										else
										{
											TextRange2 C = null;
											int D = 0;
											Text.A(shape, text2, ref C, ref D);
											if (D == 1)
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
												Text.A(ref A3, C);
												Index.IndexedLinks.Add(innerText, Text.A(shape, C, A3));
											}
											else
											{
												A3.ParentNode.RemoveChild(A3);
											}
											C = null;
											Text.A(ref A);
											if (xmlDocument.DocumentElement.ChildNodes.Count > 0)
											{
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													SetXml(shape, xmlDocument.OuterXml);
													break;
												}
											}
											else
											{
												BreakLink(shape);
											}
										}
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										ProjectData.ClearProjectError();
									}
									textRange2 = null;
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
								A3 = null;
								xmlNodeList = null;
								Text.A = shape;
								TextRange2 textRange3 = Sel.TextRange2;
								int start = textRange3.Start;
								using (Dictionary<string, TextLink>.Enumerator enumerator2 = Index.IndexedLinks.GetEnumerator())
								{
									while (true)
									{
										if (enumerator2.MoveNext())
										{
											if (!Text.A(enumerator2.Current.Value.TextRange, textRange3, start))
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
												Ribbon.LinkSelected = Ribbon.LinkSelection.Yes;
												break;
											}
											break;
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_048e;
											}
											continue;
											end_IL_048e:
											break;
										}
										break;
									}
								}
								textRange3 = null;
							}
							xmlDocument = null;
						}
					}
				}
				if (Index.IndexedLinks != null)
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
					if (Sel.TextRange2.Length == 0)
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
						TextRange2 textRange3 = Sel.TextRange2;
						int start = textRange3.Start;
						using (Dictionary<string, TextLink>.Enumerator enumerator3 = Index.IndexedLinks.GetEnumerator())
						{
							while (true)
							{
								if (enumerator3.MoveNext())
								{
									TextRange2 textRange4 = enumerator3.Current.Value.TextRange;
									int start2 = textRange4.Start;
									if (start <= start2)
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
									if (start >= start2 + textRange4.Length)
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
										new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
										try
										{
											textRange4.Select();
										}
										catch (Exception ex11)
										{
											ProjectData.SetProjectError(ex11);
											Exception ex12 = ex11;
											ProjectData.ClearProjectError();
										}
										new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(clsRibbon.Application_WindowSelectionChange));
										break;
									}
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_05f4;
									}
									continue;
									end_IL_05f4:
									break;
								}
								break;
							}
						}
						textRange3 = null;
					}
				}
				shape = null;
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void A(ref bool A)
	{
		if (!A)
		{
			NG.A.Application.StartNewUndoEntry();
			A = true;
		}
	}

	public static string XpathQuery()
	{
		return AH.A(94272);
	}

	public static string XpathNodeById(string strRangeId)
	{
		return AH.A(95304) + XML_NODE_RANGE_ID + AH.A(95319) + strRangeId + AH.A(68449);
	}

	private static string A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		return A.Tags[TAG_TEXT_LINK_XML];
	}

	public static void SetXml(Microsoft.Office.Interop.PowerPoint.Shape shp, string strXml)
	{
		shp.Tags.Add(TAG_TEXT_LINK_XML, strXml);
	}

	public static bool LinkSelected(Selection sel)
	{
		bool result = false;
		if (Index.IndexedLinks != null)
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
			TextRange2 textRange = sel.TextRange2;
			int start = textRange.Start;
			using (Dictionary<string, TextLink>.Enumerator enumerator = Index.IndexedLinks.GetEnumerator())
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						if (!A(enumerator.Current.Value.TextRange, textRange, start))
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
							result = true;
							break;
						}
						break;
					}
					while (true)
					{
						switch (1)
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
					break;
				}
			}
			textRange = null;
		}
		return result;
	}

	private static List<TextLink> A(Selection A)
	{
		List<TextLink> list = new List<TextLink>();
		if (Index.IndexedLinks != null)
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
			TextRange2 textRange = A.TextRange2;
			int start = textRange.Start;
			using (Dictionary<string, TextLink>.Enumerator enumerator = Index.IndexedLinks.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					KeyValuePair<string, TextLink> current = enumerator.Current;
					if (!Text.A(current.Value.TextRange, textRange, start))
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
					list.Add(current.Value);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0086;
					}
					continue;
					end_IL_0086:
					break;
				}
			}
			textRange = null;
		}
		return list;
	}

	public static List<TextLink> SelectedLinks(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		List<TextLink> list = new List<TextLink>();
		bool A = false;
		try
		{
			string text = Text.A(shp);
			if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0 && shp.HasTextFrame == MsoTriState.msoTrue)
			{
				IEnumerator enumerator = default(IEnumerator);
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
					XmlDocument xmlDocument = new XmlDocument();
					xmlDocument.LoadXml(text);
					XmlNodeList xmlNodeList = xmlDocument.SelectNodes(XpathQuery());
					if (xmlNodeList != null)
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
							enumerator = xmlNodeList.GetEnumerator();
							while (enumerator.MoveNext())
							{
								XmlNode A2 = (XmlNode)enumerator.Current;
								TextRange2 textRange;
								TextRange2 C;
								try
								{
									string text2 = Text.A(A2);
									textRange = shp.TextFrame2.TextRange.get_Characters(Text.A(A2), text2.Length);
									if (Operators.CompareString(textRange.Text, text2, TextCompare: false) == 0)
									{
										list.Add(Text.A(shp, textRange, A2));
									}
									else
									{
										C = null;
										int D = 0;
										Text.A(shp, text2, ref C, ref D);
										if (D == 1)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												Text.A(ref A2, C);
												list.Add(Text.A(shp, C, A2));
												Text.A(ref A);
												SetXml(shp, xmlDocument.OuterXml);
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
								textRange = null;
								C = null;
							}
							while (true)
							{
								switch (4)
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
						xmlNodeList = null;
					}
					xmlDocument = null;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		return list;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, string B, ref TextRange2 C, ref int D)
	{
		int length = B.Length;
		D = 0;
		TextRange2 textRange = A.TextFrame2.TextRange;
		C = textRange.Find(B, 0, MsoTriState.msoTrue, MsoTriState.msoTrue);
		checked
		{
			TextRange2 textRange2;
			for (textRange2 = C; textRange2 != null; textRange2 = textRange.Find(B, textRange2.Start + length - 1, MsoTriState.msoTrue, MsoTriState.msoTrue))
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
				if (textRange2.Length != length)
				{
					break;
				}
				D++;
				if (D > 1)
				{
					break;
				}
			}
			textRange = null;
			textRange2 = null;
		}
	}

	private static bool A(TextRange2 A, TextRange2 B, int C)
	{
		int start = A.Start;
		checked
		{
			if (start >= C)
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
				if (start <= C + B.Length)
				{
					return true;
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
			if (C > start)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return C < start + A.Length;
					}
				}
			}
			return false;
		}
	}

	private static void A(ref XmlNode A, TextRange2 B)
	{
		A.SelectSingleNode(XML_NODE_START).InnerText = B.Start.ToString();
	}

	private static TextLink A(Microsoft.Office.Interop.PowerPoint.Shape A, TextRange2 B, XmlNode C)
	{
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Expected O, but got Unknown
		return new TextLink(A, B, C.OuterXml);
	}

	public static void RefreshLinks(Selection sel)
	{
		List<TextLink> listTextLinks = null;
		List<Hyperlink> b = null;
		try
		{
			listTextLinks = A(sel);
			if (listTextLinks.Count > 0)
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
				if (Hyperlinks.PromptToConvert())
				{
					Hyperlinks.ConvertFromLegacyLinks(ref listTextLinks);
				}
			}
			b = Hyperlinks.SelectedLinks(sel);
			A(listTextLinks, b, C: false);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		ReleaseHelper.ClearListReferences<TextLink>(ref listTextLinks, false, (Action<TextLink>)null);
		ReleaseHelper.ClearListReferences<Hyperlink>(ref b, false, (Action<Hyperlink>)null);
		ReleaseHelper.DoGarbageCollection();
	}

	private static void A(List<TextLink> A, List<Hyperlink> B, bool C)
	{
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005a: Expected O, but got Unknown
		//IL_005c: Expected O, but got Unknown
		//IL_01ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b4: Expected O, but got Unknown
		//IL_01b6: Expected O, but got Unknown
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		//IL_03d1: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<object, string> dictionary = new Dictionary<object, string>();
		RefreshInstance val = null;
		checked
		{
			int num = A.Count + B.Count;
			int num2 = 0;
			bool flag = false;
			Microsoft.Office.Interop.PowerPoint.Application application;
			List<string> listUpdatedShapeNames;
			if (num > 0)
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
				wpfLinkRefresh wpfLinkRefresh = new wpfLinkRefresh();
				try
				{
					val = new RefreshInstance(System.Windows.Window.GetWindow(wpfLinkRefresh));
				}
				catch (UpdateLinkException ex)
				{
					ProjectData.SetProjectError((Exception)ex);
					UpdateLinkException ex2 = ex;
					wpfLinkRefresh = null;
					dictionary = null;
					Forms.WarningMessage(((Exception)(object)ex2).Message);
					ProjectData.ClearProjectError();
					return;
				}
				wpfLinkRefresh.Show();
				listUpdatedShapeNames = new List<string>();
				application = NG.A.Application;
				application.StartNewUndoEntry();
				try
				{
					foreach (TextLink item in A)
					{
						if (wpfLinkRefresh.Canceled)
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
							if (val.Canceled)
							{
								break;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								num2++;
								Common.UpdateProgressStart(wpfLinkRefresh, num2, num);
								try
								{
									Refresh(item, C, ref listUpdatedShapeNames, ref val);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									dictionary.Add(item, ex4.Message);
									if (!C)
									{
										clsReporting.LogException(ex4);
									}
									ProjectData.ClearProjectError();
								}
								Common.UpdateProgressFinish(wpfLinkRefresh, num2, num);
								break;
							}
							goto IL_0132;
						}
						break;
						IL_0132:;
					}
					foreach (Hyperlink item2 in B)
					{
						if (wpfLinkRefresh.Canceled || val.Canceled)
						{
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							num2++;
							Common.UpdateProgressStart(wpfLinkRefresh, num2, num);
							try
							{
								Hyperlinks.Refresh(item2, C, ref listUpdatedShapeNames, ref val);
							}
							catch (UpdateLinkException ex5)
							{
								ProjectData.SetProjectError((Exception)ex5);
								UpdateLinkException ex6 = ex5;
								dictionary.Add(item2, ((Exception)(object)ex6).Message);
								ProjectData.ClearProjectError();
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								dictionary.Add(item2, ex8.Message);
								if (!C)
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
									clsReporting.LogException(ex8);
								}
								ProjectData.ClearProjectError();
							}
							Common.UpdateProgressFinish(wpfLinkRefresh, num2, num);
							break;
						}
					}
					Thread.Sleep(500);
					int num3;
					if (!wpfLinkRefresh.Canceled)
					{
						if (!val.Canceled)
						{
							num3 = 0;
							goto IL_0257;
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
					num3 = ((num2 < num) ? 1 : 0);
					goto IL_0257;
					IL_0257:
					flag = unchecked((byte)num3) != 0;
					wpfLinkRefresh.Close();
					wpfLinkRefresh = null;
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					clsReporting.LogException(ex10);
					ProjectData.ClearProjectError();
				}
				ExcelToPowerPoint.ActivatePowerPoint(application);
				Base.ReleaseRefreshInstance(ref val, true);
			}
			if (!dictionary.Any())
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
				if (num > 0)
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
						if (!flag)
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
							Forms.SuccessMessage(AH.A(95324));
						}
					}
				}
				else if (!C)
				{
					Forms.InfoMessage(AH.A(95401));
				}
				else
				{
					Forms.InfoMessage(AH.A(95450));
				}
			}
			else if (num == 1)
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
				Forms.ErrorMessage(dictionary.Values.ElementAtOrDefault(0));
			}
			else if (System.Windows.Forms.MessageBox.Show(AH.A(95533), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Hand) == DialogResult.OK)
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
				List<LinkError> list = new List<LinkError>();
				wpfLinkUpdateErrors wpfLinkUpdateErrors2 = new wpfLinkUpdateErrors();
				wpfLinkUpdateErrors2.colShapeOrSlide.Header = AH.A(70464);
				using (Dictionary<object, string>.Enumerator enumerator3 = dictionary.GetEnumerator())
				{
					while (enumerator3.MoveNext())
					{
						KeyValuePair<object, string> current3 = enumerator3.Current;
						string strName;
						if (current3.Key is TextLink)
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
							strName = ((TextLink)current3.Key).TextRange.Text;
						}
						else
						{
							strName = ((Hyperlink)current3.Key).ScreenTip;
						}
						list.Add(new LinkError(RuntimeHelpers.GetObjectValue(current3.Key), strName, current3.Value));
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0430;
						}
						continue;
						end_IL_0430:
						break;
					}
				}
				wpfLinkUpdateErrors2.lvErrors.ItemsSource = list;
				list = null;
				wpfLinkUpdateErrors2.ShowDialog();
				wpfLinkUpdateErrors2 = null;
			}
			application = null;
			listUpdatedShapeNames = null;
			dictionary = null;
			Common.LogActivity(AH.A(95681));
		}
	}

	public static TextLink Refresh(TextLink tl, bool blnAll, ref List<string> listUpdatedShapeNames, ref RefreshInstance refreshInstance)
	{
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		if (blnAll)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Common.NavigateToSlide(tl.TextRange, NG.A.Application);
		}
		tl.TextRange.Select();
		A(ref tl, ref refreshInstance, LinkDetails(tl));
		listUpdatedShapeNames.Add(tl.Shape.Name);
		return tl;
	}

	private static void A(ref TextLink A, ref RefreshInstance B, Link C)
	{
		//IL_01f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_006e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
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
			UpdateSource(A, C.RangeId, C.Source, blnUpdateLastModified: false);
		}
		if (flag)
		{
			return;
		}
		XlSheetVisibility xlSheetVisibility = default(XlSheetVisibility);
		if (range != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					B.SourceRange(C, workbook, ref name, ref worksheet, ref range, ref text, ref xlSheetVisibility);
					if (range.Cells.Count > 1)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								throw new UpdateLinkException(AH.A(94708));
							}
						}
					}
					string text2 = range.Text.ToString().Trim();
					A.TextRange.Text = text2;
					string rangeId = C.RangeId;
					UpdateSource(A, rangeId, workbook.FullName, blnUpdateLastModified: true);
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					Dictionary<string, string> dictionary2 = dictionary;
					dictionary2.Add(Base.XML_NODE_USER, workbook.Application.UserName);
					dictionary2.Add(Base.XML_NODE_TIME, Base.LastUpdate());
					dictionary2.Add(XML_NODE_TEXT, text2);
					dictionary2.Add(XML_NODE_VALUE, range.Value2.ToString());
					dictionary2.Add(XML_NODE_START, A.TextRange.Start.ToString());
					dictionary2.Add(Base.XML_NODE_ADDRESS, name.RefersTo.ToString());
					_ = null;
					Text.A(A, rangeId, dictionary);
					dictionary = null;
					if (worksheet != null)
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
						if (worksheet.Visible != xlSheetVisibility)
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

	public static Microsoft.Office.Interop.PowerPoint.Shape TextRangeParentShape(TextRange2 rng)
	{
		if (NewLateBinding.LateGet(rng.Parent, null, AH.A(28234), new object[0], null, null, null) is Microsoft.Office.Interop.PowerPoint.Shape)
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
					return (Microsoft.Office.Interop.PowerPoint.Shape)NewLateBinding.LateGet(rng.Parent, null, AH.A(28234), new object[0], null, null, null);
				}
			}
		}
		return ((Microsoft.Office.Interop.PowerPoint.ShapeRange)NewLateBinding.LateGet(rng.Parent, null, AH.A(28234), new object[0], null, null, null))[1];
	}

	public static void EditLinks(Selection sel)
	{
		Shapes.EditedShapes editedShapes = default(Shapes.EditedShapes);
		if (Common.IsManageLinksDialogOpen())
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
			List<object> list = new List<object>();
			list.AddRange(A(sel).ToArray());
			using (List<Hyperlink>.Enumerator enumerator = Hyperlinks.SelectedLinks(sel).GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					Hyperlink current = enumerator.Current;
					list.Add(current);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_006a;
					}
					continue;
					end_IL_006a:
					break;
				}
			}
			if (list.Any())
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
				editedShapes = EditLink(list);
				Common.LogActivity(AH.A(95720));
			}
			else
			{
				Forms.WarningMessage(AH.A(95753));
			}
			editedShapes.ClearReferences(doDeepClearing: false, collectGarbage: true);
			return;
		}
	}

	public static Shapes.EditedShapes EditLink(List<object> listLinks)
	{
		Shapes.EditedShapes result = new Shapes.EditedShapes
		{
			Objects = listLinks,
			IsError = null,
			Errors = null
		};
		wpfLinkEdit wpfLinkEdit2;
		wpfLinkEdit obj = (wpfLinkEdit2 = new wpfLinkEdit(listLinks));
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
			result = wpfLinkEdit2.ReturnValue;
		}
		wpfLinkEdit2 = null;
		GC.Collect();
		Common.LinkEditFailed(result.Errors, listLinks.Count);
		return result;
	}

	public static void ViewSource(Selection sel)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		int intViewed = 0;
		try
		{
			using List<TextLink>.Enumerator enumerator = A(sel).GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextLink current = enumerator.Current;
				Link val = LinkDetails(current);
				string text = Source.View(val);
				if (Operators.CompareString(text, val.Source, TextCompare: false) != 0)
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
					UpdateSource(current, val.RangeId, text, blnUpdateLastModified: true);
				}
				intViewed = checked(intViewed + 1);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_0072;
				}
				continue;
				end_IL_0072:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Hyperlinks.ViewSource(sel, ref intViewed);
		if (intViewed > 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					Common.LogActivity(AH.A(95790));
					return;
				}
			}
		}
		Forms.WarningMessage(AH.A(95827));
	}

	public static void ViewSource(TextLink tl)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		int num = 0;
		try
		{
			Link val = LinkDetails(tl);
			string text = Source.View(val);
			if (Operators.CompareString(text, val.Source, TextCompare: false) != 0)
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
				UpdateSource(tl, val.RangeId, text, blnUpdateLastModified: true);
			}
			num = checked(num + 1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (num > 0)
		{
			Common.LogActivity(AH.A(95790));
		}
		else
		{
			Forms.WarningMessage(AH.A(95827));
		}
	}

	public static void ViewSource(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
		if (ContainsLinks(shp))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using List<Link>.Enumerator enumerator = LinkDetails(shp).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Link current = enumerator.Current;
						string text = Source.View(current);
						if (Operators.CompareString(text, current.Source, TextCompare: false) != 0)
						{
							Common.UpdateSource(shp.Tags, null, text, blnUpdateLastModified: true);
						}
					}
					while (true)
					{
						switch (3)
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
		throw new Exception(AH.A(95914));
	}

	public static void BreakLinks(Selection sel, bool blnUpdateRibbon)
	{
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.PowerPoint.Shape shape;
		XmlDocument A;
		try
		{
			shape = sel.ShapeRange[1];
			A = Text.A(shape);
			using (List<TextLink>.Enumerator enumerator = Text.A(sel).GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					TextLink current = enumerator.Current;
					Text.A(ref A, LinkDetails(current));
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
					break;
				}
			}
			Text.A(A, shape);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		shape = null;
		A = null;
		Hyperlinks.BreakLinks(sel);
		if (!blnUpdateRibbon)
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
			clsRibbon.InvalidateLinkedItemControls();
			return;
		}
	}

	public static void BreakLink(TextLink tl)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		XmlDocument A;
		try
		{
			A = Text.A(tl.Shape);
			Text.A(ref A, LinkDetails(tl));
			Text.A(A, tl.Shape);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		A = null;
	}

	public static void BreakLink(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		try
		{
			shp.Tags.Delete(TAG_TEXT_LINK_XML);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static XmlDocument A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(Text.A(A));
		return xmlDocument;
	}

	private static void A(ref XmlDocument A, Link B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		XmlNode xmlNode = A.SelectSingleNode(XpathNodeById(B.RangeId));
		if (xmlNode != null)
		{
			xmlNode.ParentNode.RemoveChild(xmlNode);
			xmlNode = null;
		}
	}

	private static void A(XmlDocument A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		try
		{
			if (A.SelectNodes(XpathQuery()).Count > 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						SetXml(B, A.OuterXml);
						return;
					}
				}
			}
			BreakLink(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			BreakLink(B);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(TextRange2 A, bool B)
	{
		Font2 font = A.Font;
		float size = font.Size;
		string name = font.Name;
		int rGB = font.Fill.ForeColor.RGB;
		if (B)
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
			if (rGB != ColorTranslator.ToOle(Color.White))
			{
				font.UnderlineColor.RGB = Text.A();
			}
			else
			{
				font.UnderlineColor.RGB = Text.B();
			}
			font.UnderlineStyle = Text.A();
		}
		else
		{
			font.UnderlineColor.RGB = 0;
			font.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
		}
		font.Size = size;
		font.Name = name;
		font.Fill.ForeColor.RGB = rGB;
		font = null;
	}

	private static int A()
	{
		return ColorTranslator.ToOle(Color.FromArgb(36, 180, 126));
	}

	private static int B()
	{
		return ColorTranslator.ToOle(Color.FromArgb(255, 255, 153));
	}

	private static MsoTextUnderlineType A()
	{
		return MsoTextUnderlineType.msoUnderlineDoubleLine;
	}
}
