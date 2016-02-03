using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace ParallelLinq
{
    public abstract class MandelbrotGenerator
    {
        private readonly ImageOptions options;
        protected ImageOptions Options { get { return options; } }

        // These properties just make the printed code shorter
        // than Options.ImageWidth and Options.ImageHeight
        protected int Height { get { return options.ImageHeight; } }
        protected int Width { get { return options.ImageWidth; } }

        protected MandelbrotGenerator(ImageOptions options)
        {
            this.options = options;
        }

        protected abstract byte[] GeneratePixels();

        public void Display()
        {
            Stopwatch timer = Stopwatch.StartNew();
            byte[] imageData = GeneratePixels();
            timer.Stop();

            Console.WriteLine("Generation took {0}ms", timer.ElapsedMilliseconds);

            using (Image image = CreateBitmap(imageData))
            {
                Form form = new Form
                {
                    Controls = { new PictureBox { Image = image, Dock = DockStyle.Fill } },
                    FormBorderStyle = FormBorderStyle.FixedSingle,
                    ClientSize = image.Size,
                    Text = "Mandelbrot"
                };
                // Blocks until closed
                Application.Run(form);
            }
        }

        private Image CreateBitmap(byte[] imageData)
        {
            unsafe
            {
                fixed (byte* ptr = imageData)
                {
                    IntPtr scan0 = new IntPtr(ptr);
                    Bitmap bitmap = new Bitmap(Width, Height, Width, PixelFormat.Format8bppIndexed, scan0);
                    ColorPalette palette = bitmap.Palette;
                    palette.Entries[0] = Color.White;
                    for (int i = 1; i < 256; i++)
                    {
                        palette.Entries[i] = Color.FromArgb((i * 7) % 256, (i * 5) % 256, 255);
                    }
                    bitmap.Palette = palette;
                    return (Image)bitmap.Clone();
                }
            }
        }

        protected byte ComputeIndex(int row, int col)
        {
            double x = (col * options.SampleWidth) / Width + options.OffsetX;
            double y = (row * options.SampleHeight) / Height + options.OffsetY;

            double y0 = y;
            double x0 = x;

            for (int i = 0; i < options.MaxIterations; i++)
            {
                if (x * x + y * y >= 4)
                {
                    return (byte)((i % 255) + 1);
                }
                double xtemp = x * x - y * y + x0;
                y = 2 * x * y + y0;
                x = xtemp;
            }
            return 0;
        }


    }
}
