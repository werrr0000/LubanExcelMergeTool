using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.VisualTree;

namespace LubanExcelMerge.Desktop;

public sealed class SynchronizedDataGrid : DataGrid
{
    protected override Type StyleKeyOverride => typeof(DataGrid);
    private ScrollBar? _horizontal;
    private ScrollBar? _vertical;
    private bool _applying;
    public event Action<Vector>? ViewportChanged;
    public Vector Viewport => new(_horizontal?.Value ?? 0, _vertical?.Value ?? 0);

    protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
    {
        if (_horizontal is not null) _horizontal.PropertyChanged -= OnBarChanged;
        if (_vertical is not null) _vertical.PropertyChanged -= OnBarChanged;
        base.OnApplyTemplate(e);
        _horizontal = e.NameScope.Find<ScrollBar>("PART_HorizontalScrollbar");
        _vertical = e.NameScope.Find<ScrollBar>("PART_VerticalScrollbar");
        if (_horizontal is not null) _horizontal.PropertyChanged += OnBarChanged;
        if (_vertical is not null) _vertical.PropertyChanged += OnBarChanged;
    }

    private void OnBarChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (!_applying && e.Property == ScrollBar.ValueProperty)
            ViewportChanged?.Invoke(Viewport);
    }

    public void SetViewport(Vector offset)
    {
        var delta = Viewport - offset;
        if (Math.Abs(delta.X) < .01 && Math.Abs(delta.Y) < .01) return;
        _applying = true;
        try
        {
            // DataGrid 11.3.13 自行管理滚动条，没有 ScrollViewer。
            // 使用其滚轮入口更新实际行列布局；该版本每个滚轮单位对应 50 像素。
            using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, true);
            OnPointerWheelChanged(new PointerWheelEventArgs(this, pointer,
                this.GetVisualRoot() as Visual ?? this, default, 0, default, KeyModifiers.None, delta / 50));
            UpdateLayout();
        }
        finally { _applying = false; }
    }
}
