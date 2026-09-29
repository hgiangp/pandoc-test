using System;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Threading;

public static class ShapeRaster
{
    [DllImport("user32.dll", SetLastError = true)] static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll", SetLastError = true)] static extern bool CloseClipboard();
    [DllImport("user32.dll", SetLastError = true)] static extern bool EmptyClipboard();
    [DllImport("user32.dll")] static extern bool IsClipboardFormatAvailable(uint format);
    [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint uFormat);
    [DllImport("gdi32.dll")] static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[] lpbBuffer);
    const uint CF_ENHMETAFILE = 14;

    static bool Open()
    {
        for (int i = 0; i < 30; i++)
        {
            if (OpenClipboard(IntPtr.Zero)) return true;
            Thread.Sleep(100);
        }
        return false;
    }

    public static void ClearClipboard()
    {
        if (!Open()) return;
        try { EmptyClipboard(); } finally { CloseClipboard(); }
    }

    // Returns the Enhanced Metafile currently on the clipboard, or null.
    public static byte[] GetClipboardEmf()
    {
        if (!Open()) return null;
        try
        {
            if (!IsClipboardFormatAvailable(CF_ENHMETAFILE)) return null;
            IntPtr h = GetClipboardData(CF_ENHMETAFILE);
            if (h == IntPtr.Zero) return null;
            uint size = GetEnhMetaFileBits(h, 0, null);
            if (size == 0) return null;
            byte[] buf = new byte[size];
            GetEnhMetaFileBits(h, size, buf);
            return buf;
        }
        finally { CloseClipboard(); }
    }

    // Renders an EMF to PNG at the given DPI (physical size is preserved through the PNG DPI).
    // Returns { widthPx, heightPx }.
    public static int[] RenderEmfToPng(byte[] emf, string pngPath, int dpi, bool trim, int pad, int maxSide)
    {
        using (var ms = new MemoryStream(emf))
        using (var mf = new Metafile(ms))
        {
            MetafileHeader h = mf.GetMetafileHeader();
            double wIn = h.Bounds.Width / (double)h.DpiX;
            double hIn = h.Bounds.Height / (double)h.DpiY;
            float effDpi = dpi;
            int w = (int)Math.Ceiling(wIn * dpi);
            int ht = (int)Math.Ceiling(hIn * dpi);
            if (w < 1 || ht < 1) throw new InvalidOperationException("Empty metafile");
            if (Math.Max(w, ht) > maxSide)
            {
                double f = maxSide / (double)Math.Max(w, ht);
                w = Math.Max(1, (int)(w * f));
                ht = Math.Max(1, (int)(ht * f));
                effDpi = (float)(dpi * f);
            }

            using (var bmp = new Bitmap(w, ht, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    g.DrawImage(mf, new Rectangle(0, 0, w, ht));
                }
                Rectangle crop = trim ? ContentBounds(bmp, pad) : new Rectangle(0, 0, w, ht);
                using (var outBmp = bmp.Clone(crop, PixelFormat.Format24bppRgb))
                {
                    outBmp.SetResolution(effDpi, effDpi);
                    outBmp.Save(pngPath, ImageFormat.Png);
                    return new int[] { outBmp.Width, outBmp.Height };
                }
            }
        }
    }

    static Rectangle ContentBounds(Bitmap bmp, int pad)
    {
        var full = new Rectangle(0, 0, bmp.Width, bmp.Height);
        BitmapData data = bmp.LockBits(full, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        int stride = data.Stride;
        byte[] px = new byte[stride * bmp.Height];
        Marshal.Copy(data.Scan0, px, 0, px.Length);
        bmp.UnlockBits(data);

        int minX = bmp.Width, minY = bmp.Height, maxX = -1, maxY = -1;
        for (int y = 0; y < bmp.Height; y++)
        {
            int row = y * stride;
            for (int x = 0; x < bmp.Width; x++)
            {
                int i = row + x * 4;
                if (px[i] < 245 || px[i + 1] < 245 || px[i + 2] < 245)
                {
                    if (x < minX) minX = x;
                    if (x > maxX) maxX = x;
                    if (y < minY) minY = y;
                    if (y > maxY) maxY = y;
                }
            }
        }
        if (maxX < 0) return full;
        minX = Math.Max(0, minX - pad);
        minY = Math.Max(0, minY - pad);
        maxX = Math.Min(bmp.Width - 1, maxX + pad);
        maxY = Math.Min(bmp.Height - 1, maxY + pad);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }
}
