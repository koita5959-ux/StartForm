using System.Runtime.InteropServices;

namespace StartForm.Helpers
{
    internal static class WindowHelper
    {
        private const int BoundsTolerance = 12;

        /// <summary>
        /// シャドウを除いた実際の描画領域を取得
        /// </summary>
        public static NativeMethods.RECT GetExtendedFrameBounds(IntPtr hWnd)
        {
            NativeMethods.DwmGetWindowAttribute(
                hWnd,
                NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
                out NativeMethods.RECT rect,
                Marshal.SizeOf<NativeMethods.RECT>());
            return rect;
        }

        public static bool TryGetExtendedFrameBounds(IntPtr hWnd, out Rectangle bounds)
        {
            int hr = NativeMethods.DwmGetWindowAttribute(
                hWnd,
                NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
                out NativeMethods.RECT rect,
                Marshal.SizeOf<NativeMethods.RECT>());

            if (hr == 0 && rect.Right > rect.Left && rect.Bottom > rect.Top)
            {
                bounds = Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
                return true;
            }

            bounds = Rectangle.Empty;
            return false;
        }

        public static bool IsCloaked(IntPtr hWnd)
        {
            int cloaked;
            int hr = NativeMethods.DwmGetWindowAttribute(
                hWnd,
                NativeMethods.DWMWA_CLOAKED,
                out cloaked,
                Marshal.SizeOf<int>());

            return hr == 0 && cloaked != 0;
        }

        /// <summary>
        /// GetWindowRect と ExtendedFrameBounds の差分からシャドウ幅を算出
        /// </summary>
        public static (int left, int right, int top, int bottom) GetShadowMargins(IntPtr hWnd)
        {
            NativeMethods.GetWindowRect(hWnd, out NativeMethods.RECT windowRect);
            var frameRect = GetExtendedFrameBounds(hWnd);

            int left = frameRect.Left - windowRect.Left;
            int right = windowRect.Right - frameRect.Right;
            int top = frameRect.Top - windowRect.Top;
            int bottom = windowRect.Bottom - frameRect.Bottom;

            return (left, right, top, bottom);
        }

        /// <summary>
        /// シャドウ分を補正したRECTを返す
        /// </summary>
        public static Rectangle AdjustBoundsForShadow(IntPtr hWnd, Rectangle bounds)
        {
            var shadow = GetShadowMargins(hWnd);

            return new Rectangle(
                bounds.X - shadow.left,
                bounds.Y - shadow.top,
                bounds.Width + shadow.left + shadow.right,
                bounds.Height + shadow.top + shadow.bottom);
        }

        /// <summary>
        /// シャドウ補正込みで配置し、必要なら最終状態も適用する
        /// </summary>
        public static bool ApplyWindowPlacement(IntPtr hWnd, Rectangle bounds, string windowState)
        {
            if (hWnd == IntPtr.Zero || !NativeMethods.IsWindow(hWnd))
            {
                ExecutionLogger.Warn($"ApplyWindowPlacement skipped: invalid window handle={hWnd}");
                return false;
            }

            bool maximizeTarget = string.Equals(windowState, "Maximized", StringComparison.OrdinalIgnoreCase);
            bool minimizeTarget = string.Equals(windowState, "Minimized", StringComparison.OrdinalIgnoreCase);
            var adjustedBounds = maximizeTarget
                ? bounds
                : AdjustBoundsForShadow(hWnd, bounds);
            var placement = new NativeMethods.WINDOWPLACEMENT
            {
                length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>(),
                showCmd = NativeMethods.SW_SHOWNORMAL,
                rcNormalPosition = new NativeMethods.RECT
                {
                    Left = adjustedBounds.Left,
                    Top = adjustedBounds.Top,
                    Right = adjustedBounds.Right,
                    Bottom = adjustedBounds.Bottom
                }
            };

            NativeMethods.GetWindowPlacement(hWnd, ref placement);
            placement.length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>();
            placement.flags = 0;
            placement.showCmd = NativeMethods.SW_SHOWNORMAL;
            placement.rcNormalPosition = new NativeMethods.RECT
            {
                Left = adjustedBounds.Left,
                Top = adjustedBounds.Top,
                Right = adjustedBounds.Right,
                Bottom = adjustedBounds.Bottom
            };

            bool placementResult = NativeMethods.SetWindowPlacement(hWnd, ref placement);
            bool showResult = minimizeTarget || maximizeTarget
                ? true
                : NativeMethods.ShowWindowAsync(hWnd, NativeMethods.SW_SHOWNORMAL);
            bool posResult = maximizeTarget
                ? true
                : NativeMethods.SetWindowPos(
                    hWnd,
                    IntPtr.Zero,
                    adjustedBounds.X,
                    adjustedBounds.Y,
                    adjustedBounds.Width,
                    adjustedBounds.Height,
                    NativeMethods.SWP_NOZORDER |
                    NativeMethods.SWP_NOOWNERZORDER |
                    NativeMethods.SWP_NOACTIVATE |
                    NativeMethods.SWP_ASYNCWINDOWPOS |
                    (minimizeTarget ? 0u : NativeMethods.SWP_SHOWWINDOW));

            ApplyWindowState(hWnd, windowState);
            ExecutionLogger.Info(
                $"ApplyWindowPlacement hwnd=0x{hWnd.ToInt64():X} target={bounds} adjusted={adjustedBounds} state={windowState} " +
                $"setPlacement={placementResult} show={showResult} setPos={posResult}");
            return true;
        }

        public static Rectangle GetStableWindowBounds(IntPtr hWnd)
        {
            var placement = new NativeMethods.WINDOWPLACEMENT
            {
                length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>()
            };

            if (NativeMethods.GetWindowPlacement(hWnd, ref placement) &&
                placement.showCmd == NativeMethods.SW_SHOWMINIMIZED)
            {
                var rect = placement.rcNormalPosition;
                return Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
            }

            if (TryGetExtendedFrameBounds(hWnd, out Rectangle frameBounds))
                return frameBounds;

            NativeMethods.GetWindowRect(hWnd, out NativeMethods.RECT windowRect);
            return Rectangle.FromLTRB(windowRect.Left, windowRect.Top, windowRect.Right, windowRect.Bottom);
        }

        public static bool IsPlacementClose(IntPtr hWnd, Rectangle targetBounds, int tolerance = BoundsTolerance)
        {
            if (hWnd == IntPtr.Zero || !NativeMethods.IsWindow(hWnd))
                return false;

            var current = GetStableWindowBounds(hWnd);
            return Math.Abs(current.X - targetBounds.X) <= tolerance &&
                   Math.Abs(current.Y - targetBounds.Y) <= tolerance &&
                   Math.Abs(current.Width - targetBounds.Width) <= tolerance &&
                   Math.Abs(current.Height - targetBounds.Height) <= tolerance;
        }

        /// <summary>
        /// ウィンドウの表示状態を文字列で返す
        /// </summary>
        public static string GetWindowState(IntPtr hWnd)
        {
            var placement = new NativeMethods.WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf<NativeMethods.WINDOWPLACEMENT>();
            NativeMethods.GetWindowPlacement(hWnd, ref placement);

            return placement.showCmd switch
            {
                1 => "Normal",
                2 => "Minimized",
                3 => "Maximized",
                _ => "Normal"
            };
        }

        public static void ApplyWindowState(IntPtr hWnd, string windowState)
        {
            bool result;
            switch (windowState)
            {
                case "Maximized":
                    result = NativeMethods.ShowWindowAsync(hWnd, NativeMethods.SW_MAXIMIZE);
                    break;
                case "Minimized":
                    result = NativeMethods.ShowWindowAsync(hWnd, NativeMethods.SW_MINIMIZE);
                    break;
                default:
                    result = NativeMethods.ShowWindowAsync(hWnd, NativeMethods.SW_RESTORE);
                    break;
            }

            ExecutionLogger.Info(
                $"ApplyWindowState hwnd=0x{hWnd.ToInt64():X} state={windowState} result={result}");
        }
    }
}
