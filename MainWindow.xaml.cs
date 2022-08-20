using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ScreenShot
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool isMouseDown = false;
        private System.Windows.Point startPoint;
        private System.Windows.Shapes.Rectangle rectangle;
        private System.Windows.Controls.Image bakupImage = new System.Windows.Controls.Image();
        private System.Windows.Controls.Image selectedImage;
        private double ratio;

        public MainWindow()
        {
            InitializeComponent();
            ratio = System.Windows.Forms.Screen.AllScreens[0].Bounds.Width / SystemParameters.PrimaryScreenWidth;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GetScreenshot();
        }

        // 螢幕快照
        public void GetScreenSnapshot()
        {
            double screenLeft = SystemParameters.VirtualScreenLeft;
            double screenTop = SystemParameters.VirtualScreenTop;
            double screenWidth = SystemParameters.VirtualScreenWidth;
            double screenHeight = SystemParameters.VirtualScreenHeight;

            using (Bitmap bmp = new Bitmap((int)screenWidth, (int)screenHeight))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    Visibility = Visibility.Collapsed;

                    string path = @"C:\Users\ADMIN\Desktop\";
                    string filename = "ScreenCapture-" + DateTime.Now.ToString("ddMMyyyy-hhmmss") + ".png";
                    g.CopyFromScreen((int)screenLeft, (int)screenTop, 0, 0, bmp.Size);
                    bmp.Save(path + filename);

                    Visibility = Visibility.Visible;
                }
            }

        }

        // 螢幕截圖
        public void GetScreenshot()
        {
            double bakTop = Top;
            double bakLeft = Left;
            double bakWidth = Width;
            double bakHeight = Height;

            double screenLeft = SystemParameters.VirtualScreenLeft;
            double screenTop = SystemParameters.VirtualScreenTop;
            double screenWidth = SystemParameters.VirtualScreenWidth * ratio;
            double screenHeight = SystemParameters.VirtualScreenHeight * ratio;

            Width = 0;
            Height = 0;
            Top = int.MinValue;
            Left = int.MinValue;

            using (Bitmap bmp = new Bitmap((int)screenWidth, (int)screenHeight))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen((int)screenLeft, (int)screenTop, 0, 0, bmp.Size);
                    System.Windows.Controls.Image image = new System.Windows.Controls.Image();
                    image.Source = BitmapToImageSource(bmp);
                    bakupImage.Source = BitmapToImageSource(bmp);

                    Window window = new Window();

                    window.Width = screenWidth / ratio;
                    window.Height = screenHeight / ratio;
                    window.Left = screenLeft / ratio;
                    window.Top = screenTop;
                    window.Owner = this;
                    window.WindowStyle = WindowStyle.None;
                    window.ResizeMode = ResizeMode.NoResize;
                    window.ShowInTaskbar = false;
                    window.Cursor = Cursors.Cross;

                    window.PreviewMouseRightButtonUp += Window_PreviewMouseRightButtonUp;
                    window.PreviewMouseLeftButtonUp += Window_PreviewMouseLeftButtonUp;
                    window.PreviewMouseLeftButtonDown += Window_PreviewMouseLeftButtonDown;
                    window.MouseMove += Window_MouseMove;

                    Grid grid = new Grid();
                    Grid shadowGrid = new Grid();
                    Canvas canvas = new Canvas();
                    rectangle = new System.Windows.Shapes.Rectangle();
                    selectedImage = new System.Windows.Controls.Image();
                    System.Windows.Media.Brush brush = new SolidColorBrush(Colors.Black);
                    brush.Opacity = 0.5;

                    grid.Margin = new Thickness(0);
                    grid.Children.Add(image);

                    canvas.Children.Add(selectedImage);
                    canvas.Children.Add(rectangle);
                    shadowGrid.Margin = new Thickness(0);
                    shadowGrid.Background = brush;
                    grid.Children.Add(shadowGrid);
                    grid.Children.Add(canvas);

                    window.Content = grid;
                    window.Show();
                }
            }

            Width = bakWidth;
            Height = bakHeight;
            Top = bakTop;
            Left = bakLeft;
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isMouseDown)
                return;

            var point = e.GetPosition(sender as Window);
            Rect rect = new Rect(startPoint, point);
            rectangle.Width = rect.Width;
            rectangle.Height = rect.Height;
            rectangle.StrokeThickness = 2;
            rectangle.Stroke = System.Windows.Media.Brushes.Aqua;

            if ((int)rect.Width > 0 && (int)rect.Height > 0)
            {
                var selectedRect = new Int32Rect((int)(rect.X * ratio), (int)(rect.Y * ratio), (int)(rect.Width * ratio), (int)(rect.Height * ratio));
                selectedImage.Source = new CroppedBitmap((BitmapSource)bakupImage.Source, selectedRect);
            }

            Canvas.SetLeft(selectedImage, rect.X);
            Canvas.SetTop(selectedImage, rect.Y);
            Canvas.SetLeft(rectangle, rect.X);
            Canvas.SetTop(rectangle, rect.Y);
        }

        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            startPoint = e.GetPosition(sender as Window);
            isMouseDown = true;
        }

        private void Window_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            isMouseDown = false;
            if (SaveImage(selectedImage.Source))
            {
                Clipboard.SetImage((BitmapSource)selectedImage.Source);
                Window window = sender as Window;
                window.Close();
            }
            selectedImage.Source = null;
        }

        private void Window_PreviewMouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            isMouseDown = false;
            Window window = sender as Window;
            window.Close();
        }

        private ImageSource BitmapToImageSource(Bitmap bitmap)
        {
            MemoryStream stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Bmp);
            stream.Position = 0;
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = stream;
            bitmapImage.EndInit();
            return bitmapImage;
        }

        private string OpenFolderPicker()
        {
            CommonOpenFileDialog dialog = new();
            dialog.Title = "Folder Picker";
            dialog.Multiselect = false;
            dialog.IsFolderPicker = true;

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                if (dialog.FileNames != null)
                {
                    foreach (string fileName in dialog.FileNames)
                    {
                        return fileName;
                    }
                }
            }

            return "";
        }

        private bool SaveImage(ImageSource imageSource)
        {
            try
            {
                if (imageSource == null)
                    return false;

                string path = OpenFolderPicker();

                if (string.IsNullOrEmpty(path))
                    return false;

                string filename = "\\ScreenCapture-" + DateTime.Now.ToString("ddMMyyyy-hhmmss") + ".png";

                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create((BitmapSource)imageSource));
                using (FileStream stream = new FileStream(path + filename, FileMode.Create))
                {
                    encoder.Save(stream);
                }

                return true;
            }
            catch { }

            return false;
        }
    }
}
