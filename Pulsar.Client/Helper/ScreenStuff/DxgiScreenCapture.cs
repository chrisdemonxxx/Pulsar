using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using Device = SharpDX.Direct3D11.Device;

namespace Pulsar.Client.Helper
{
    /// <summary>
    /// Provides high performance screen capture using the DXGI Desktop Duplication API.
    /// </summary>
    public class DxgiScreenCapture : IDisposable
    {
        private Device _device;
        private OutputDuplication _duplication;
        private Texture2D _stagingTexture;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of <see cref="DxgiScreenCapture"/>.
        /// </summary>
        public DxgiScreenCapture()
        {
            Initialize();
        }

        private void Initialize()
        {
            _device = new Device(SharpDX.Direct3D.DriverType.Hardware,
                DeviceCreationFlags.BgraSupport | DeviceCreationFlags.SingleThreaded);

            using (var adapter = _device.QueryInterface<SharpDX.DXGI.Device>()
                                        .GetParent<Adapter>())
            {
                using (var output = adapter.GetOutput(0))
                using (var output1 = output.QueryInterface<Output1>())
                {
                    _duplication = output1.DuplicateOutput(_device);
                }
            }
        }

        /// <summary>
        /// Captures the latest available frame from the primary monitor.
        /// </summary>
        /// <returns>A bitmap containing the captured frame, or null if no frame is available.</returns>
        public Bitmap GetLatestFrame()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(DxgiScreenCapture));

            try
            {
                _duplication.AcquireNextFrame(100, out var frameInfo, out var screenResource);
                using (var texture = screenResource.QueryInterface<Texture2D>())
                {
                    EnsureStagingTexture(texture);
                    _device.ImmediateContext.CopyResource(texture, _stagingTexture);
                }
                _duplication.ReleaseFrame();

                return CopyStagingTextureToBitmap();
            }
            catch (SharpDXException ex) when (ex.ResultCode.Code == ResultCode.WaitTimeout.Result.Code)
            {
                // No new frame available
                return null;
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        private void EnsureStagingTexture(Texture2D texture)
        {
            if (_stagingTexture != null &&
                _stagingTexture.Description.Width == texture.Description.Width &&
                _stagingTexture.Description.Height == texture.Description.Height)
                return;

            _stagingTexture?.Dispose();
            var desc = texture.Description;
            desc.CpuAccessFlags = CpuAccessFlags.Read;
            desc.BindFlags = BindFlags.None;
            desc.OptionFlags = ResourceOptionFlags.None;
            desc.Usage = ResourceUsage.Staging;
            desc.SampleDescription.Count = 1;
            desc.SampleDescription.Quality = 0;
            _stagingTexture = new Texture2D(_device, desc);
        }

        private Bitmap CopyStagingTextureToBitmap()
        {
            var width = _stagingTexture.Description.Width;
            var height = _stagingTexture.Description.Height;
            var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            var sourceBox = _device.ImmediateContext.MapSubresource(_stagingTexture, 0, MapMode.Read, MapFlags.None);
            var destData = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, bitmap.PixelFormat);

            for (int y = 0; y < height; y++)
            {
                Utilities.CopyMemory(destData.Scan0 + y * destData.Stride, sourceBox.DataPointer + y * sourceBox.RowPitch, width * 4);
            }

            bitmap.UnlockBits(destData);
            _device.ImmediateContext.UnmapSubresource(_stagingTexture, 0);
            return bitmap;
        }

        /// <summary>
        /// Releases all resources used by this instance.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
                return;

            _stagingTexture?.Dispose();
            _duplication?.Dispose();
            _device?.Dispose();
            _disposed = true;
        }
    }
}
