using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class ChartNavigator
    {
        private readonly Excel.Application _app;

        public ChartNavigator(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public bool MoveSelectedLabels(double dx, double dy) => TryMoveSelectedLabel(dx, dy);

        public bool MoveSelectedChart(double dx, double dy) => TryMoveSelectedChart(dx, dy);

        public void SelectNearestChart()
        {
            var sheet = _app.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return;
            }

            if (!TryGetSelectionCenter(out var refX, out var refY))
            {
                RunStatusMessage("Unable to determine reference position.", 2000);
                return;
            }

            object bestTarget = null;
            double bestDist = -1;

            var chartObjects = sheet.ChartObjects() as Excel.ChartObjects;
            if (chartObjects != null)
            {
                for (int i = 1; i <= chartObjects.Count; i++)
                {
                    var cbo = chartObjects.Item(i) as Excel.ChartObject;
                    if (cbo == null || !ChartObjectIsVisible(cbo))
                    {
                        continue;
                    }

                    double dist = DistanceSquared(refX, refY, cbo.Left + cbo.Width / 2, cbo.Top + cbo.Height / 2);
                    if (bestDist < 0 || dist < bestDist)
                    {
                        bestDist = dist;
                        bestTarget = cbo;
                    }
                }
            }

            var shapes = sheet.Shapes;
            if (shapes != null)
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shp = shapes.Item(i) as Excel.Shape;
                    if (shp == null || !ShapeHasVisibleChart(shp))
                    {
                        continue;
                    }

                    double dist = DistanceSquared(refX, refY, shp.Left + shp.Width / 2, shp.Top + shp.Height / 2);
                    if (bestDist < 0 || dist < bestDist)
                    {
                        bestDist = dist;
                        bestTarget = shp;
                    }
                }
            }


            if (bestTarget == null)
            {
                RunStatusMessage("No charts found on this sheet.", 2000);
                return;
            }

            ActivateChartContainer(bestTarget);
            EnsureChartElementSelection(bestTarget);
        }

        private bool TryMoveSelectedLabel(double dx, double dy)
        {
            object selection = _app.Selection;
            if (selection == null)
            {
                return false;
            }

            string typeName;
            try
            {
                typeName = selection.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, selection, null)?.ToString();
            }
            catch
            {
                typeName = selection.GetType().Name;
            }

            try
            {
                switch (_app.Selection)
                {
                    case Excel.DataLabel lbl:
                        lbl.Position = Excel.XlDataLabelPosition.xlLabelPositionCustom;
                        lbl.Left += dx;
                        lbl.Top += dy;
                        return true;
                    case Excel.Point pt:
                        if (pt.HasDataLabel)
                        {
                            var dLabel = pt.DataLabel;
                            dLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionCustom;
                            dLabel.Left += dx;
                            dLabel.Top += dy;
                            return true;
                        }

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                // ignore
            }

            if (_app.Selection is Excel.Series series)
            {
                try
                {
                    var points = series.Points() as Excel.Points;
                    if (points != null)
                    {
                        for (int i = 1; i <= points.Count; i++)
                        {
                            var point = points.Item(i) as Excel.Point;
                            if (point == null || !point.HasDataLabel)
                            {
                                continue;
                            }

                            var lbl = point.DataLabel;
                            lbl.Position = Excel.XlDataLabelPosition.xlLabelPositionCustom;
                            lbl.Left += dx;
                            lbl.Top += dy;
                        }
                    }

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            if (_app.Selection is Excel.DataLabels labels)
            {
                try
                {
                    for (int i = 1; i <= labels.Count; i++)
                    {
                        var lbl = labels.Item(i) as Excel.DataLabel;
                        if (lbl == null)
                        {
                            continue;
                        }

                        lbl.Position = Excel.XlDataLabelPosition.xlLabelPositionCustom;
                        lbl.Left += dx;
                        lbl.Top += dy;
                    }

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        private bool TryMoveSelectedChart(double dx, double dy)
        {
            object selection = _app.Selection;
            if (selection == null)
            {
                return false;
            }

            var container = ResolveSelectedChartContainer(selection);
            if (container == null)
            {
                return false;
            }

            try
            {
                dynamic dyn = container;
                dyn.Left += dx;
                dyn.Top += dy;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private object ResolveSelectedChartContainer(object seed)
        {
            object current = seed;
            int hops = 0;

            while (current != null && hops < 20)
            {
                hops++;

                var typeName = current.GetType().Name;
                if (typeName == "ChartObject" || typeName == "Shape")
                {
                    return current;
                }

                if (typeName == "ShapeRange")
                {
                    dynamic range = current;
                    if (range.Count == 1)
                    {
                        current = range.Item(1);
                        continue;
                    }

                    break;
                }

                if (typeName == "Chart")
                {
                    dynamic chart = current;
                    var parent = chart.Parent;
                    if (parent is Excel.ChartObject || parent is Excel.Shape)
                    {
                        return parent;
                    }
                }

                try
                {
                    current = current.GetType().GetProperty("Parent")?.GetValue(current);
                }
                catch
                {
                    current = null;
                }
            }

            return null;
        }

        private bool TryGetSelectionCenter(out double centerX, out double centerY)
        {
            centerX = 0;
            centerY = 0;

            Excel.Range rng = null;
            try
            {
                if (_app.Selection is Excel.Range sel)
                {
                    rng = sel;
                }
            }
            catch
            {
                rng = null;
            }

            if (rng == null)
            {
                try
                {
                    rng = _app.ActiveCell;
                }
                catch
                {
                    rng = null;
                }
            }

            if (rng != null)
            {
                centerX = Convert.ToDouble(rng.Left) + Convert.ToDouble(rng.Width) / 2.0;
                centerY = Convert.ToDouble(rng.Top) + Convert.ToDouble(rng.Height) / 2.0;
                return true;
            }

            try
            {
                var window = _app.ActiveWindow;
                if (window != null)
                {
                    var vis = window.VisibleRange;
                    if (vis != null)
                    {
                        centerX = Convert.ToDouble(vis.Left) + Convert.ToDouble(vis.Width) / 2.0;
                        centerY = Convert.ToDouble(vis.Top) + Convert.ToDouble(vis.Height) / 2.0;
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private static double DistanceSquared(double x1, double y1, double x2, double y2)
        {
            double dx = x1 - x2;
            double dy = y1 - y2;
            return dx * dx + dy * dy;
        }

        private static bool ChartObjectIsVisible(Excel.ChartObject cbo)
        {
            if (cbo == null)
            {
                return false;
            }

            try
            {
                var visible = Convert.ToInt32(cbo.Visible) != 0;
                return visible && cbo.Width > 0 && cbo.Height > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool ShapeHasVisibleChart(Excel.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                return shape.Visible != Microsoft.Office.Core.MsoTriState.msoFalse && Convert.ToBoolean(shape.HasChart);
            }
            catch
            {
                return false;
            }
        }

        private void ActivateChartContainer(object target)
        {
            if (target == null)
            {
                return;
            }

            try
            {
                switch (target)
                {
                    case Excel.ChartObject chartObject:
                        RangeHelpers.SafeActivateSheet(chartObject.Parent as Excel.Worksheet);
                        chartObject.Activate();
                        break;
                    case Excel.Shape shape:
                        RangeHelpers.SafeActivateSheet(shape.Parent as Excel.Worksheet);
                        if (Convert.ToBoolean(shape.HasChart))
                        {
                            var chartParent = shape.Chart?.Parent;
                            ActivateComObject(chartParent);
                        }
                        else
                        {
                            shape.Select(Type.Missing);
                        }

                        break;
                    case Excel.Chart chart:
                        var parent = chart.Parent;
                        if (parent is Excel.ChartObject co)
                        {
                            RangeHelpers.SafeActivateSheet(co.Parent as Excel.Worksheet);
                            co.Activate();
                        }
                        else
                        {
                            ActivateComObject(chart.Parent);
                        }

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                // ignore
            }
        }

        private void EnsureChartElementSelection(object container)
        {
            if (container == null)
            {
                return;
            }

            Excel.Chart chart = null;
            switch (container)
            {
                case Excel.ChartObject co:
                    chart = co.Chart;
                    break;
                case Excel.Shape shape when Convert.ToBoolean(shape.HasChart):
                    chart = shape.Chart;
                    break;
                case Excel.Chart c:
                    chart = c;
                    break;
            }

            if (chart == null)
            {
                return;
            }

                try
                {
                    chart.ChartArea.Select();
                }
                catch
                {
                    ActivateComObject(chart.Parent);
                }
        }

        private static void ActivateComObject(object target)
        {
            if (target == null)
            {
                return;
            }

            try
            {
                dynamic dyn = target;
                dyn.Activate();
            }
            catch
            {
                // ignore
            }
        }

        private void RunStatusMessage(string message, int milliseconds)
        {
            try
            {
                _app.Run("SetStatusBarTemporarily", message, milliseconds);
            }
            catch
            {
                // ignore
            }
        }
    }
}
