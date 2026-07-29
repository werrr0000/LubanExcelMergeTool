using LubanExcelMerge.Core;

var tests = new (string Name, Action Run)[]
{
    ("REMOTE-only cell change", RemoteOnlyCellChange),
    ("LOCAL-only cell change", LocalOnlyCellChange),
    ("identical cell change", IdenticalCellChange),
    ("different cell change conflicts", DifferentCellChangeConflicts),
    ("numeric representations compare equal", NumericRepresentationsCompareEqual),
    ("formula text takes precedence over cache", FormulaTextTakesPrecedence),
    ("different fields merge automatically", DifferentFieldsMerge),
    ("LOCAL-only record addition", LocalOnlyAddition),
    ("REMOTE-only record addition", RemoteOnlyAddition),
    ("identical add/add merges once", IdenticalAddition),
    ("different add/add conflicts", DifferentAdditionConflicts),
    ("delete unchanged record", DeleteUnchangedRecord),
    ("delete/modify conflicts", DeleteModifyConflicts),
    ("both delete record", BothDeleteRecord)
};

var failures = new List<string>();
foreach (var test in tests)
{
    try
    {
        test.Run();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception exception)
    {
        failures.Add($"FAIL {test.Name}: {exception.Message}");
    }
}

foreach (var failure in failures)
    Console.Error.WriteLine(failure);

Console.WriteLine($"Executed {tests.Length} tests: {tests.Length - failures.Count} passed, {failures.Count} failed.");
return failures.Count == 0 ? 0 : 1;

static CellPayload Text(string value) => new(CellValueKind.String, value);
static CellPayload Number(string value) => new(CellValueKind.Number, value);
static CellPayload Formula(string formula, string cache) => new(CellValueKind.Formula, FormulaText: formula, CachedValue: cache);
static LubanRecord Record(string key, params (string Field, CellPayload Value)[] fields) =>
    new(key, fields.Select(field => new KeyValuePair<string, CellPayload>(field.Field, field.Value)));

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"Expected {expected}, got {actual}.");
}

static void True(bool value)
{
    if (!value)
        throw new InvalidOperationException("Expected true, got false.");
}

static void RemoteOnlyCellChange()
{
    var result = CellThreeWayMerger.Merge(Text("A"), Text("A"), Text("B"), "1", "Name");
    Equal(MergeDecisionKind.TakeRemote, result.Kind);
    Equal("B", result.Result!.RawValue);
}

static void LocalOnlyCellChange()
{
    var result = CellThreeWayMerger.Merge(Text("A"), Text("B"), Text("A"), "1", "Name");
    Equal(MergeDecisionKind.TakeLocal, result.Kind);
}

static void IdenticalCellChange()
{
    var result = CellThreeWayMerger.Merge(Text("A"), Text("B"), Text("B"), "1", "Name");
    Equal(MergeDecisionKind.BothChangedIdentically, result.Kind);
}

static void DifferentCellChangeConflicts()
{
    var result = CellThreeWayMerger.Merge(Text("A"), Text("B"), Text("C"), "1", "Name");
    Equal(MergeConflictKind.CellChangedDifferently, result.Conflict!.Kind);
}

static void NumericRepresentationsCompareEqual()
{
    True(Number("1").ContentEquals(Number("1.0")));
    True(!Text("1").ContentEquals(Number("1")));
}

static void FormulaTextTakesPrecedence()
{
    True(!Formula("A1+B1", "10").ContentEquals(Formula("A1+C1", "10")));
    True(Formula("A1+B1", "10").ContentEquals(Formula("A1+B1", "11")));
}

static void DifferentFieldsMerge()
{
    var @base = Record("1", ("A", Text("old-a")), ("B", Text("old-b")));
    var local = Record("1", ("A", Text("local-a")), ("B", Text("old-b")));
    var remote = Record("1", ("A", Text("old-a")), ("B", Text("remote-b")));
    var result = RecordThreeWayMerger.Merge(@base, local, remote);
    Equal("local-a", result.Record!.Fields["A"].RawValue);
    Equal("remote-b", result.Record.Fields["B"].RawValue);
}

static void LocalOnlyAddition()
{
    var result = RecordThreeWayMerger.Merge(null, Record("1", ("A", Text("x"))), null);
    Equal(MergeDecisionKind.AddedLocal, result.Kind);
}

static void RemoteOnlyAddition()
{
    var result = RecordThreeWayMerger.Merge(null, null, Record("1", ("A", Text("x"))));
    Equal(MergeDecisionKind.AddedRemote, result.Kind);
}

static void IdenticalAddition()
{
    var record = Record("1", ("A", Text("x")));
    var result = RecordThreeWayMerger.Merge(null, record, Record("1", ("A", Text("x"))));
    Equal(MergeDecisionKind.AddedIdentically, result.Kind);
}

static void DifferentAdditionConflicts()
{
    var result = RecordThreeWayMerger.Merge(null, Record("1", ("A", Text("x"))), Record("1", ("A", Text("y"))));
    Equal(MergeConflictKind.AddAdd, result.Conflicts.Single().Kind);
}

static void DeleteUnchangedRecord()
{
    var record = Record("1", ("A", Text("x")));
    var result = RecordThreeWayMerger.Merge(record, null, Record("1", ("A", Text("x"))));
    Equal(MergeDecisionKind.Deleted, result.Kind);
}

static void DeleteModifyConflicts()
{
    var result = RecordThreeWayMerger.Merge(
        Record("1", ("A", Text("x"))),
        null,
        Record("1", ("A", Text("y"))));
    Equal(MergeConflictKind.DeleteModify, result.Conflicts.Single().Kind);
}

static void BothDeleteRecord()
{
    var result = RecordThreeWayMerger.Merge(Record("1", ("A", Text("x"))), null, null);
    Equal(MergeDecisionKind.Deleted, result.Kind);
}
