using System.ComponentModel;
using System.Windows;

namespace LubanExcelMerge.Gui;

public partial class SaveProgressWindow : Window
{
    private bool _canClose;

    public SaveProgressWindow()
    {
        InitializeComponent();
    }

    public void Complete()
    {
        _canClose = true;
        Close();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!_canClose)
            e.Cancel = true;
        base.OnClosing(e);
    }
}
