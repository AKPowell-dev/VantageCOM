using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VantagePackageHolder
{
    internal sealed class WorkbookAnalysisService
    {
        private const string EdgeDelimiter = "~~>";
        private readonly Excel.Application _app;
        private static readonly Regex SheetRefRegex = new Regex(@"('([^']+)'|[A-Za-z0-9_]+)!", RegexOptions.Compiled);

        public WorkbookAnalysisService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void DrawDependencyMap()
        {
            var workbook = _app.ActiveWorkbook;
            if (workbook == null)
            {
                return;
            }

            var previousCalc = _app.Calculation;
            var prevScreenUpdating = _app.ScreenUpdating;
            UiGuard guard = null;
            try
            {
                guard = new UiGuard(_app, hideStatusBar: true);
                _app.Calculation = Excel.XlCalculation.xlCalculationManual;
                _app.ScreenUpdating = false;

                var edges = GetSheetDependencies(workbook);
                var levels = BuildSheetLevels(workbook, edges);
                var mapSheet = CreateMapSheet(workbook);
                DrawNodesAndEdges(mapSheet, levels, edges);
                RangeHelpers.SafeActivateSheet(mapSheet);
            }
            catch
            {
                // ignore
            }
            finally
            {
                guard?.Dispose();
                _app.Calculation = previousCalc;
                _app.ScreenUpdating = prevScreenUpdating;
            }
        }

        private Dictionary<string, bool> GetSheetDependencies(Excel.Workbook workbook)
        {
            var edges = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                Excel.Range formulas = null;
                try
                {
                    formulas = sheet.UsedRange?.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                }
                catch
                {
                    formulas = null;
                }

                if (formulas == null)
                {
                    continue;
                }

                foreach (Excel.Range cell in formulas.Cells)
                {
                    try
                    {
                        var refs = ExtractSheetReferences(Convert.ToString(cell.Formula), sheet.Name, workbook);
                        foreach (var refName in refs)
                        {
                            var key = $"{refName}{EdgeDelimiter}{sheet.Name}";
                            if (!edges.ContainsKey(key))
                            {
                                edges[key] = true;
                            }
                        }
                    }
                    catch
                    {
                        // ignore formula parse issues
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell);
                    }
                }
            }

            return edges;
        }

        private IEnumerable<string> ExtractSheetReferences(string formula, string currentSheet, Excel.Workbook workbook)
        {
            if (string.IsNullOrEmpty(formula))
            {
                yield break;
            }

            foreach (Match match in SheetRefRegex.Matches(formula))
            {
                string candidate = match.Value;
                candidate = candidate.Substring(0, candidate.Length - 1);
                string sheetName;
                if (candidate.StartsWith("'", StringComparison.Ordinal))
                {
                    sheetName = candidate.Substring(1, candidate.Length - 2).Replace("''", "'");
                }
                else
                {
                    sheetName = candidate;
                }

                int bracket = sheetName.LastIndexOf(']');
                if (bracket >= 0)
                {
                    sheetName = sheetName.Substring(bracket + 1);
                }

                if (!string.Equals(sheetName, currentSheet, StringComparison.OrdinalIgnoreCase) && SheetExists(workbook, sheetName))
                {
                    yield return sheetName;
                }
            }
        }

        private bool SheetExists(Excel.Workbook workbook, string sheetName)
        {
            try
            {
                var ws = workbook.Worksheets[sheetName] as Excel.Worksheet;
                return ws != null;
            }
            catch
            {
                return false;
            }
        }

        private Dictionary<string, int> BuildSheetLevels(Excel.Workbook workbook, Dictionary<string, bool> edges)
        {
            var inbound = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var outbound = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                inbound[sheet.Name] = new List<string>();
                outbound[sheet.Name] = new List<string>();
            }

            foreach (var edge in edges.Keys)
            {
                var parts = edge.Split(new[] { EdgeDelimiter }, StringSplitOptions.None);
                if (parts.Length != 2)
                {
                    continue;
                }

                var src = parts[0];
                var tgt = parts[1];
                if (!outbound.TryGetValue(src, out var srcList))
                {
                    continue;
                }

                srcList.Add(tgt);
                if (!inbound.TryGetValue(tgt, out var tgtList))
                {
                    continue;
                }

                tgtList.Add(src);
            }

            var levels = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var queue = new Queue<string>();

            foreach (var kvp in inbound)
            {
                if (kvp.Value.Count == 0)
                {
                    levels[kvp.Key] = 0;
                    queue.Enqueue(kvp.Key);
                }
            }

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                int level = levels[current];
                foreach (var downstream in outbound[current])
                {
                    int newLevel = level + 1;
                    if (!levels.ContainsKey(downstream) || levels[downstream] < newLevel)
                    {
                        levels[downstream] = newLevel;
                        queue.Enqueue(downstream);
                    }
                }
            }

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (!levels.ContainsKey(sheet.Name))
                {
                    levels[sheet.Name] = 0;
                }
            }

            return levels;
        }

        private Excel.Worksheet CreateMapSheet(Excel.Workbook workbook)
        {
            const string sheetName = "Dependency Map";
            Excel.Worksheet mapSheet = null;
            Excel.Worksheet existing = null;
            try
            {
                existing = workbook.Worksheets[sheetName] as Excel.Worksheet;
                if (existing != null)
                {
                    existing.Delete();
                }
            }
            catch
            {
                // ignore if sheet missing
            }
            finally
            {
                if (existing != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(existing);
                }
            }

            mapSheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            mapSheet.Name = sheetName;
            mapSheet.Cells.Clear();
            try
            {
                mapSheet.Shapes.SelectAll();
                var shapeRange = _app.Selection as Excel.ShapeRange;
                shapeRange?.Delete();
            }
            catch
            {
                // ignore if nothing selected
            }
            return mapSheet;
        }

        private void DrawNodesAndEdges(Excel.Worksheet mapSheet, Dictionary<string, int> levels, Dictionary<string, bool> edges)
        {
            if (mapSheet == null)
            {
                return;
            }

            var levelBuckets = new Dictionary<int, List<string>>();
            foreach (var kvp in levels)
            {
                if (!levelBuckets.TryGetValue(kvp.Value, out var list))
                {
                    list = new List<string>();
                    levelBuckets[kvp.Value] = list;
                }

                list.Add(kvp.Key);
            }

            int maxCount = levelBuckets.Values.Select(l => l.Count).DefaultIfEmpty(1).Max();
            var nodeShapes = new Dictionary<string, Excel.Shape>(StringComparer.OrdinalIgnoreCase);

            const double nodeWidth = 140;
            const double nodeHeight = 60;
            const double horizontalSpacing = 110;
            const double verticalSpacing = 40;
            const double leftMargin = 60;
            const double topMargin = 60;
            double columnHeight = Math.Max(maxCount * (nodeHeight + verticalSpacing), nodeHeight + verticalSpacing);

            foreach (var level in levelBuckets.Keys.OrderBy(v => v))
            {
                var names = levelBuckets[level].OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToArray();
                double x = leftMargin + level * (nodeWidth + horizontalSpacing);
                double totalHeight = names.Length * nodeHeight + (names.Length - 1) * verticalSpacing;
                double y = topMargin + (columnHeight - totalHeight) / 2;

                foreach (var name in names)
                {
                    var shape = mapSheet.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, (float)x, (float)y, (float)nodeWidth, (float)nodeHeight);
                    shape.TextFrame2.TextRange.Text = name;
                    shape.TextFrame2.TextRange.Font.Size = 12;
                    shape.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 235, 247));
                    shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(79, 129, 189));
                    shape.Line.Weight = 1.5f;
                    shape.Name = $"node_{name}";
                    nodeShapes[name] = shape;
                    y += nodeHeight + verticalSpacing;
                }
            }

            int edgeColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(84, 130, 53));
            foreach (var edge in edges.Keys)
            {
                var parts = edge.Split(new[] { EdgeDelimiter }, StringSplitOptions.None);
                if (parts.Length != 2)
                {
                    continue;
                }

                if (!nodeShapes.TryGetValue(parts[0], out var srcShape) || !nodeShapes.TryGetValue(parts[1], out var tgtShape))
                {
                    continue;
                }

                var connector = mapSheet.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 0, 0, 0, 0);
                connector.Line.ForeColor.RGB = edgeColor;
                connector.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
                connector.Line.Weight = 1.5f;
                connector.ConnectorFormat.BeginConnect(srcShape, 2);
                connector.ConnectorFormat.EndConnect(tgtShape, 1);
                connector.RerouteConnections();
                connector.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }

            AddLegend(mapSheet, edgeColor);
            var title = mapSheet.Range["A1"];
            title.Value2 = "Dependency Map";
            title.Font.Bold = true;
            title.Font.Size = 16;
        }

        private void AddLegend(Excel.Worksheet mapSheet, int lineColor)
        {
            var box = mapSheet.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 20, 20, 220, 60);
            box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(242, 242, 242));
            box.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 191, 191));
            box.TextFrame2.TextRange.Text = "Legend" + Environment.NewLine + "Arrow: upstream sheet feeds downstream sheet";
            box.TextFrame2.TextRange.Font.Size = 10;
            box.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            box.TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

            var line = mapSheet.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 40, 60, 120, 60);
            line.Line.ForeColor.RGB = lineColor;
            line.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
            line.Line.Weight = 1.5f;
        }
    }
}
