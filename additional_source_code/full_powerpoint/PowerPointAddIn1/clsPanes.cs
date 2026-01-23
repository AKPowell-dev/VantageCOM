using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

public sealed class clsPanes
{
	public static bool PaneAlreadyExists(ref Dictionary<int, CustomTaskPane> TaskPanes, Microsoft.Office.Interop.PowerPoint.Presentation pres, string strPaneTitle)
	{
		int hWND = pres.Windows[1].HWND;
		bool result = false;
		checked
		{
			CustomTaskPane customTaskPane;
			try
			{
				if (TaskPanes != null)
				{
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
						for (int i = TaskPanes.Count - 1; i >= 0; i += -1)
						{
							try
							{
								customTaskPane = TaskPanes.ElementAt(i).Value;
								try
								{
									if (Operators.CompareString(customTaskPane.Title, strPaneTitle, TextCompare: false) == 0)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											if (!customTaskPane.Visible)
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
												if (TaskPanes.ElementAt(i).Key != hWND)
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
													customTaskPane = null;
													result = true;
													break;
												}
												goto end_IL_0109;
											}
											break;
										}
									}
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									NG.A.CustomTaskPanes.Remove(customTaskPane);
									TaskPanes.Remove(TaskPanes.ElementAt(i).Key);
									ProjectData.ClearProjectError();
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							customTaskPane = null;
							continue;
							end_IL_0109:
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
			customTaskPane = null;
			return result;
		}
	}

	public static int CountVisiblePanes(string strPaneTitle)
	{
		int num = 0;
		IEnumerator<CustomTaskPane> enumerator = default(IEnumerator<CustomTaskPane>);
		try
		{
			enumerator = NG.A.CustomTaskPanes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomTaskPane current = enumerator.Current;
				try
				{
					if (Operators.CompareString(current.Title, strPaneTitle, TextCompare: false) != 0)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (current.Visible)
						{
							num = checked(num + 1);
						}
						break;
					}
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
				switch (5)
				{
				case 0:
					continue;
				}
				return num;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	public static int CountVisiblePresentations()
	{
		int num = 0;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = NG.A.Application.Presentations.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current;
				try
				{
					if (presentation.Windows.Count > 0)
					{
						num = checked(num + 1);
					}
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
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		int result = default(int);
		return result;
	}

	public static bool IsPowerPoint2013OrNewer()
	{
		return Conversion.Val(NG.A.Application.Version) >= 15.0;
	}

	public static void RemoveOrphanedPanes(ref Dictionary<int, CustomTaskPane> TaskPanes, string strPaneTitle)
	{
		if (TaskPanes == null)
		{
			return;
		}
		checked
		{
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
				for (int i = TaskPanes.Count - 1; i >= 0; i += -1)
				{
					CustomTaskPane value = TaskPanes.ElementAt(i).Value;
					try
					{
						Operators.CompareString(value.Title, strPaneTitle, TextCompare: false);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						NG.A.CustomTaskPanes.Remove(value);
						TaskPanes.Remove(TaskPanes.ElementAt(i).Key);
						ProjectData.ClearProjectError();
					}
					value = null;
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
	}

	public static bool PresentationCloseFinal(ref Dictionary<int, CustomTaskPane> TaskPanes, Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		CustomTaskPane value = null;
		bool result = false;
		try
		{
			if (TaskPanes != null)
			{
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
					int hWND = Pres.Windows[1].HWND;
					if (TaskPanes.TryGetValue(hWND, out value))
					{
						NG.A.CustomTaskPanes.Remove(value);
						TaskPanes.Remove(hWND);
						result = true;
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
		return result;
	}

	private static CustomTaskPane A(Dictionary<int, CustomTaskPane> A)
	{
		CustomTaskPane value = null;
		try
		{
			if (A != null)
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
					A.TryGetValue(NG.A.Application.ActiveWindow.HWND, out value);
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
		return value;
	}

	public static bool IsVisible(Dictionary<int, CustomTaskPane> panes)
	{
		bool result;
		try
		{
			result = A(panes).Visible;
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

	public static void RemoveTaskPanesByTitle(string strTitle)
	{
		checked
		{
			try
			{
				CustomTaskPaneCollection customTaskPanes = NG.A.CustomTaskPanes;
				for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
				{
					try
					{
						if (Operators.CompareString(customTaskPanes[i].Title, strTitle, TextCompare: false) != 0)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							customTaskPanes.RemoveAt(i);
							break;
						}
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
					switch (2)
					{
					case 0:
						continue;
					}
					customTaskPanes = null;
					return;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
	}

	internal static void A(CustomTaskPane A, clsDisplay B, int C = 0)
	{
		if (C == 0)
		{
			C = PB.Settings.TaskPaneWidth;
		}
		try
		{
			A.Width = checked((int)Math.Round((double)C * B.X));
		}
		catch (ArgumentException ex)
		{
			ProjectData.SetProjectError(ex);
			ArgumentException ex2 = ex;
			A.Width = 430;
			ProjectData.ClearProjectError();
		}
	}
}
