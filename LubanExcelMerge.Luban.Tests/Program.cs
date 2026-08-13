using LubanExcelMerge.Luban;

var tests = new List<(string Name, Action Run)>
{
    ("CSV quoted commas, escaped quotes and newlines", CsvQuotedFields),
    ("CSV bare quotes in Luban metadata are preserved", CsvBareQuotesInMetadata),
    ("logical table metadata rows are ignored", CatalogIgnoresMetadata),
    ("catalog matches normalized input path", CatalogMatchesPath),
    ("compound index declaration", CompoundIndex),
    ("multiple index declaration", MultipleIndexes),
    ("dynamic metadata area and empty columns", DynamicMetadataArea),
    ("field names remain case-sensitive", FieldNamesAreCaseSensitive),
    ("duplicate keys report every source row", DuplicateKeys),
    ("empty composite key is rejected", EmptyCompositeKey),
    ("numeric key representations and defaults normalize", NumericKeysNormalize),
    ("selector falls back to next unique index", SelectorFallsBack),
    ("first exported field is inferred", FirstFieldIsInferred),
    ("record key encoding cannot collide", RecordKeyEncodingCannotCollide)
};

if (args is ["--tables", var tablesPath])
    tests.Add(("real project __tables__.csv smoke test", () => RealTablesSmokeTest(tablesPath)));

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

Console.WriteLine($"Executed {tests.Count} tests: {tests.Count - failures.Count} passed, {failures.Count} failed.");
return failures.Count == 0 ? 0 : 1;

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

static LubanRawRow Row(int number, params string?[] cells) => new(number, cells);
static LubanRecordSource Source(int row, params (string Name, string? Value)[] values) =>
    new("book.xlsx", "Sheet1", row, values.ToDictionary(value => value.Name, value => value.Value, StringComparer.Ordinal));

static LubanSchema Schema() => LubanSchemaParser.Parse(new[]
{
    Row(1, "##var", "Id", "SubId", "Name"),
    Row(2, "##", "编号", "子编号", "名称"),
    Row(3, "##type", "int", "string", "string"),
    Row(4, null, "1", "1", "first")
});

static void CsvQuotedFields()
{
    var rows = CsvReader.ReadAll(new StringReader("a,\"b,c\",\"line1\nline2\",\"say \"\"hi\"\"\"\r\n"));
    Equal("b,c", rows[0][1]);
    Equal("line1\nline2", rows[0][2]);
    Equal("say \"hi\"", rows[0][3]);
}

static void CsvBareQuotesInMetadata()
{
    const string csv = "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
                       "##,TbLubanTest,LubanTest,TRUE,LubanTest.xlsx,Id\",key2,key3\n" +
                       ",TbAccountLv,AccountLv,TRUE,AccountLv.xlsx,,,c\n";
    var rows = CsvReader.ReadAll(new StringReader(csv));
    Equal("Id\"", rows[1][5]);

    var catalog = LogicalTableCatalog.Parse(new StringReader(csv));
    Equal(1, catalog.Tables.Count);
    Equal("TbAccountLv", catalog.Tables[0].FullName);
}

static void CatalogIgnoresMetadata()
{
    const string csv = "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
                       "##,说明,,,,,,\n" +
                       ",TbAccountLv,AccountLv,TRUE,AccountLv.xlsx,,,c\n";
    var catalog = LogicalTableCatalog.Parse(new StringReader(csv));
    Equal(1, catalog.Tables.Count);
    Equal("TbAccountLv", catalog.Tables[0].FullName);
}

static void CatalogMatchesPath()
{
    const string csv = "##var,full_name,value_type,read_schema_from_file,input,index,mode,group\n" +
                       ",TbModel,Model,TRUE,\"Model/Model.xlsx,Model/Other.xlsx\",Id,map,c\n";
    var catalog = LogicalTableCatalog.Parse(new StringReader(csv));
    Equal("TbModel", catalog.MatchInput("Model\\Other.xlsx").Single().FullName);
}

static void CompoundIndex()
{
    var index = RecordKeyDefinition.ParseDeclaration(" Id + SmallGameType ").Single();
    Equal("Id+SmallGameType", index.DisplayName);
}

static void MultipleIndexes()
{
    var indexes = RecordKeyDefinition.ParseDeclaration("Id+Type, Code");
    Equal(2, indexes.Count);
    Equal("Code", indexes[1].DisplayName);
}

static void DynamicMetadataArea()
{
    var schema = LubanSchemaParser.Parse(new[]
    {
        Row(1, "##var", null, "Id", null, "Value"),
        Row(2, "##", "说明", "编号", "布局", "值"),
        Row(3, "##type", null, "int", null, "string"),
        Row(4, "##var", null, null, null, "str"),
        Row(5, null, null, "1", null, "x")
    });
    Equal(5, schema.DataStartRowNumber);
    Equal(2, schema.Fields.Count);
    Equal(2, schema.Fields[0].ColumnIndex);
    Equal("int", schema.Fields[0].TypeName);
}

static void FieldNamesAreCaseSensitive()
{
    var result = RecordKeyValidator.Validate(Schema(), new RecordKeyDefinition(new[] { "id" }), new[] { Source(4, ("Id", "1")) });
    Equal(KeyValidationIssueKind.MissingField, result.Issues.Single().Kind);
}

static void DuplicateKeys()
{
    var result = RecordKeyValidator.Validate(
        Schema(),
        new RecordKeyDefinition(new[] { "Id" }),
        new[] { Source(4, ("Id", "1")), Source(9, ("Id", "1")) });
    Equal(2, result.Issues.Count(issue => issue.Kind == KeyValidationIssueKind.DuplicateKey));
    True(result.Issues.Any(issue => issue.RowNumber == 4));
    True(result.Issues.Any(issue => issue.RowNumber == 9));
}

static void EmptyCompositeKey()
{
    var result = RecordKeyValidator.Validate(
        Schema(),
        new RecordKeyDefinition(new[] { "Id", "SubId" }),
        new[] { Source(4, ("Id", "1"), ("SubId", "")) });
    Equal(KeyValidationIssueKind.EmptyKey, result.Issues.Single().Kind);
}

static void NumericKeysNormalize()
{
    var schema = LubanSchemaParser.Parse(new[]
    {
        Row(1, "##var", "Id", "Type"),
        Row(2, "##type", "int", "int"),
        Row(3, null, "1", null)
    });
    var result = RecordKeyValidator.Validate(
        schema,
        new RecordKeyDefinition(new[] { "Id", "Type" }),
        new[]
        {
            Source(3, ("Id", "1.0"), ("Type", null)),
            Source(4, ("Id", "1"), ("Type", "0"))
        });
    Equal(2, result.Issues.Count(issue => issue.Kind == KeyValidationIssueKind.DuplicateKey));
}

static void SelectorFallsBack()
{
    var records = new[]
    {
        Source(4, ("Id", "1"), ("SubId", "a")),
        Source(5, ("Id", "1"), ("SubId", "b"))
    };
    var selection = PrimaryKeySelector.Select(
        Schema(),
        RecordKeyDefinition.ParseDeclaration("Id,SubId"),
        records,
        records,
        records);
    Equal("SubId", selection.Selected!.DisplayName);
}

static void FirstFieldIsInferred()
{
    var records = new[] { Source(4, ("Id", "1")) };
    var selection = PrimaryKeySelector.Select(Schema(), Array.Empty<RecordKeyDefinition>(), records, records, records);
    Equal("Id", selection.Selected!.DisplayName);
}

static void RecordKeyEncodingCannotCollide()
{
    var first = new LubanRecordKey(new[] { "ab", "c" });
    var second = new LubanRecordKey(new[] { "a", "bc" });
    True(first.StableValue != second.StableValue);
}

static void RealTablesSmokeTest(string tablesPath)
{
    using var reader = new StreamReader(tablesPath, detectEncodingFromByteOrderMarks: true);
    var catalog = LogicalTableCatalog.Parse(reader);
    True(catalog.Tables.Count > 300);
    Equal("TbAccountLv", catalog.MatchInput("AccountLv.xlsx").Single().FullName);
    Equal("Id+SmallGameType", catalog.MatchInput("BattleActionConstConfig.xlsx").Single().DeclaredIndexes.Single().DisplayName);
    Equal("GroupId+DialogId", catalog.MatchInput("DialogConfig_Chapter01.xlsx").Single().DeclaredIndexes.Single().DisplayName);
}
