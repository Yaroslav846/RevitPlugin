using System.Windows;
using System.Windows.Input;

namespace RevitPlugin
{
    public partial class SettingsWindow : Window
    {
        private PluginSettings _settings;

        public SettingsWindow()
        {
            InitializeComponent();
            _settings = PluginSettings.Load();

            txtWall.Text = _settings.WallAreaParam;
            txtSkirt.Text = _settings.SkirtingParam;
            txtHeight.Text = _settings.HeightParam;
        }

        // Позволяет перетаскивать окно за любую область
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try { this.DragMove(); } catch { }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            _settings.WallAreaParam = txtWall.Text;
            _settings.SkirtingParam = txtSkirt.Text;
            _settings.HeightParam = txtHeight.Text;
            _settings.Save();

            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}