using System.ComponentModel;
using System.Runtime.CompilerServices;
using LubanExcelMerge.Cli;
using LubanExcelMerge.Core;

namespace LubanExcelMerge.Gui;

public sealed class ConflictItemViewModel : INotifyPropertyChanged
{
    public ConflictItemViewModel(ResolvableMergeConflict model) => Model = model;

    internal ResolvableMergeConflict Model { get; }
    public string Id => Model.Id;
    public string RecordKey => Model.Conflict.RecordKey;
    public string FieldName => Model.Conflict.FieldName ?? "整条记录";
    public string KindText => Model.Conflict.Kind switch
    {
        MergeConflictKind.CellChangedDifferently => "内容冲突",
        MergeConflictKind.AddAdd => "同键新增",
        MergeConflictKind.DeleteModify => "删除/修改",
        MergeConflictKind.MetadataChanged => "元数据复核",
        _ => Model.Conflict.Kind.ToString()
    };
    public bool IsMetadataChange => Model.Conflict.Kind == MergeConflictKind.MetadataChanged;
    public string RowText => Model.RowNumber?.ToString() ?? "-";
    public string BaseValue => Model.BaseValue;
    public string LocalValue => Model.LocalValue;
    public string RemoteValue => Model.RemoteValue;
    public string Message => Model.Conflict.Message;
    public MergeChoice? SelectedChoice => Model.SelectedChoice;
    public bool IsResolved => Model.IsResolved;
    public string StatusText => IsResolved ? "已解决" : "未解决";
    public string ResultText => SelectedChoice switch
    {
        MergeChoice.Base => BaseValue,
        MergeChoice.Local => LocalValue,
        MergeChoice.Remote => RemoteValue,
        _ => "待选择"
    };

    internal void Refresh()
    {
        OnPropertyChanged(nameof(SelectedChoice));
        OnPropertyChanged(nameof(IsResolved));
        OnPropertyChanged(nameof(StatusText));
        OnPropertyChanged(nameof(ResultText));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
