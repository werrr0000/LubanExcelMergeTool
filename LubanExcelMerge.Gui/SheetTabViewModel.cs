using System.ComponentModel;
using System.Runtime.CompilerServices;
using LubanExcelMerge.Cli;

namespace LubanExcelMerge.Gui;

public sealed class SheetTabViewModel : INotifyPropertyChanged
{
    public SheetTabViewModel(PreparedSheetMerge model)
    {
        Model = model;
        Conflicts = model.Conflicts.Select(conflict => new ConflictItemViewModel(conflict)).ToArray();
    }

    internal PreparedSheetMerge Model { get; }
    public string SheetName => Model.SheetName;
    public string KeyName => Model.KeyName;
    public IReadOnlyList<ConflictItemViewModel> Conflicts { get; }
    public int ConflictCount => Model.Conflicts.Count;
    public int RemainingCount => Model.RemainingConflicts;
    public int ResolvedCount => ConflictCount - RemainingCount;
    public int AutomaticMergeCount => Model.AutomaticMergeCount;
    public int ProcessedMergeCount => Model.ProcessedMergeCount;
    public int MetadataChangeCount => Model.MetadataChangeCount;
    public bool HasUnresolvedConflicts => RemainingCount > 0;
    public bool HasUnresolvedMetadataChanges => Conflicts.Any(conflict =>
        conflict.IsMetadataChange && !conflict.IsResolved);

    internal void Refresh()
    {
        OnPropertyChanged(nameof(RemainingCount));
        OnPropertyChanged(nameof(ResolvedCount));
        OnPropertyChanged(nameof(HasUnresolvedConflicts));
        OnPropertyChanged(nameof(HasUnresolvedMetadataChanges));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
