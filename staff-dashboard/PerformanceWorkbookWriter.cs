using System.IO.Compression;
using System.Security;
using System.Text;

namespace YouthSpring.StaffDashboard;

internal static class PerformanceWorkbookWriter
{
    private sealed record SheetData(string Name, List<List<CellValue>> Rows, double[] Widths);
    private sealed record CellValue(string Value, bool Numeric = false, int Style = 0);

    public static void Write(string path, MonthlyPerformanceReport report)
    {
        var sheets = BuildSheets(report);
        var tempPath = path + ".tmp";
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        if (File.Exists(tempPath)) File.Delete(tempPath);

        try
        {
            using (var archive = ZipFile.Open(tempPath, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml", ContentTypesXml(sheets.Count));
                WriteEntry(archive, "_rels/.rels", RootRelationshipsXml());
                WriteEntry(archive, "xl/workbook.xml", WorkbookXml(sheets));
                WriteEntry(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml(sheets.Count));
                WriteEntry(archive, "xl/styles.xml", StylesXml());
                for (var index = 0; index < sheets.Count; index++)
                    WriteEntry(archive, $"xl/worksheets/sheet{index + 1}.xml", WorksheetXml(sheets[index]));
            }

            File.Move(tempPath, path, true);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private static List<SheetData> BuildSheets(MonthlyPerformanceReport report)
    {
        var summaryRows = new List<List<CellValue>>
        {
            new() { Text($"청년이봄 {report.Month} 월간 실적", 1) },
            new() { Text("생성 시각"), Text(report.GeneratedAt) },
            new(),
            new() { Text("구분", 2), Text("실인원", 2), Text("연인원", 2) }
        };
        foreach (var item in report.Summary)
        {
            summaryRows.Add(new()
            {
                Text(item.Label), Number(item.UniqueUsers), Number(item.TotalUsers)
            });
        }
        summaryRows.Add(new()
        {
            Text("전체", 3), Number(report.TotalUniqueUsers, 3), Number(report.TotalUsers, 3)
        });
        summaryRows.Add(new());
        summaryRows.Add(new() { Text("총 이용자"), Number(report.TotalUsers), Text("4개 부문 연인원 합계") });
        summaryRows.Add(new() { Text("총 실인원"), Number(report.TotalUniqueUsers), Text("해당 월 전체 부문 중복 제거") });
        summaryRows.Add(new());
        summaryRows.Add(new() { Text("실인원 기준"), Text(report.UniqueBasis) });

        return new()
        {
            new SheetData("월간 요약", summaryRows, new double[] { 20, 16, 28 }),
            BuildDailySheet("방문자", report.Daily.Visit),
            BuildDailySheet("커넥트룸", report.Daily.Connect),
            BuildDailySheet("커리어존", report.Daily.Career),
            BuildDailySheet("프린터", report.Daily.Printer),
            BuildVisitSimplifiedSheet(report.Simplified.Visit),
            BuildPrinterSimplifiedSheet(report.Simplified.Printer),
            BuildRoomSimplifiedSheet("커리어존 간소화", report.Simplified.Career, true),
            BuildRoomSimplifiedSheet("커넥트룸 간소화", report.Simplified.Connect, false)
        };
    }

    private static SheetData BuildVisitSimplifiedSheet(List<PerformanceVisitSimplifiedRow> rows)
    {
        var sheetRows = new List<List<CellValue>>
        {
            new() { Text("연번", 2), Text("날짜", 2), Text("성별", 2), Text("나이", 2) }
        };
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            sheetRows.Add(new()
            {
                Number(index + 1), Text(row.Date), Text(row.Gender), Number(row.Age)
            });
        }
        return new SheetData("방문일지 간소화", sheetRows, new double[] { 9, 14, 12, 10 });
    }

    private static SheetData BuildPrinterSimplifiedSheet(List<PerformancePrinterSimplifiedRow> rows)
    {
        var sheetRows = new List<List<CellValue>>
        {
            new()
            {
                Text("연번", 2), Text("날짜", 2), Text("성별", 2), Text("나이", 2), Text("사용 매수", 2)
            }
        };
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            sheetRows.Add(new()
            {
                Number(index + 1), Text(row.Date), Text(row.Gender), Number(row.Age), Number(row.Count)
            });
        }
        return new SheetData("프린터 간소화", sheetRows, new double[] { 9, 14, 12, 10, 13 });
    }

    private static SheetData BuildRoomSimplifiedSheet(
        string name,
        List<PerformanceRoomSimplifiedRow> rows,
        bool includeSpace)
    {
        var header = new List<CellValue> { Text("연번", 2), Text("날짜", 2) };
        if (includeSpace) header.Add(Text("사용공간", 2));
        header.AddRange(
        [
            Text("사용목적", 2), Text("인원수", 2), Text("비고", 2),
            Text("시작시간", 2), Text("종료시간", 2),
            Text("20~29세(남)", 2), Text("20~29세(여)", 2),
            Text("30~39세(남)", 2), Text("30~39세(여)", 2),
            Text("~19세(남)", 2), Text("~19세(여)", 2),
            Text("40세 이상(남)", 2), Text("40세 이상(여)", 2)
        ]);
        var sheetRows = new List<List<CellValue>> { header };
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            var data = new List<CellValue> { Number(index + 1), Text(row.Date) };
            if (includeSpace) data.Add(Text(row.Space));
            data.AddRange(
            [
                Text(row.Purpose), Number(row.Headcount), Text(row.Note),
                Text(row.StartTime), Text(row.EndTime),
                Number(row.Male20s), Number(row.Female20s),
                Number(row.Male30s), Number(row.Female30s),
                Number(row.MaleUnder19), Number(row.FemaleUnder19),
                Number(row.Male40Plus), Number(row.Female40Plus)
            ]);
            sheetRows.Add(data);
        }

        var widths = new List<double> { 9, 14 };
        if (includeSpace) widths.Add(13);
        widths.AddRange(new double[] { 24, 10, 28, 12, 12, 14, 14, 14, 14, 13, 13, 16, 16 });
        return new SheetData(name, sheetRows, widths.ToArray());
    }

    private static SheetData BuildDailySheet(string name, List<PerformanceDailyRow> rows)
    {
        var sheetRows = new List<List<CellValue>>
        {
            new()
            {
                Text("날짜", 2), Text("19세 이하 남성", 2), Text("19세 이하 여성", 2),
                Text("20~24세 남성", 2), Text("20~24세 여성", 2),
                Text("25~29세 남성", 2), Text("25~29세 여성", 2),
                Text("30~34세 남성", 2), Text("30~34세 여성", 2),
                Text("35~39세 남성", 2), Text("35~39세 여성", 2),
                Text("40세 이상 남성", 2), Text("40세 이상 여성", 2),
                Text("미분류", 2), Text("합계", 2)
            }
        };
        foreach (var row in rows)
        {
            sheetRows.Add(new()
            {
                Text(row.Date), Number(row.MaleUnder19), Number(row.FemaleUnder19),
                Number(row.Male20To24), Number(row.Female20To24),
                Number(row.Male25To29), Number(row.Female25To29),
                Number(row.Male30To34), Number(row.Female30To34),
                Number(row.Male35To39), Number(row.Female35To39),
                Number(row.Male40Plus), Number(row.Female40Plus),
                Number(row.Unknown), Number(row.Total, 3)
            });
        }
        sheetRows.Add(new()
        {
            Text("합계", 3),
            Number(rows.Sum(row => row.MaleUnder19), 3),
            Number(rows.Sum(row => row.FemaleUnder19), 3),
            Number(rows.Sum(row => row.Male20To24), 3),
            Number(rows.Sum(row => row.Female20To24), 3),
            Number(rows.Sum(row => row.Male25To29), 3),
            Number(rows.Sum(row => row.Female25To29), 3),
            Number(rows.Sum(row => row.Male30To34), 3),
            Number(rows.Sum(row => row.Female30To34), 3),
            Number(rows.Sum(row => row.Male35To39), 3),
            Number(rows.Sum(row => row.Female35To39), 3),
            Number(rows.Sum(row => row.Male40Plus), 3),
            Number(rows.Sum(row => row.Female40Plus), 3),
            Number(rows.Sum(row => row.Unknown), 3),
            Number(rows.Sum(row => row.Total), 3)
        });
        return new SheetData(name, sheetRows, new double[]
        {
            14, 16, 16, 15, 15, 15, 15, 15, 15, 15, 15, 16, 16, 12, 12
        });
    }

    private static CellValue Text(string value, int style = 0) => new(value ?? "", false, style);
    private static CellValue Number(int value, int style = 0) => new(value.ToString(), true, style);

    private static string WorksheetXml(SheetData sheet)
    {
        var builder = new StringBuilder();
        builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        builder.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
        builder.Append("<cols>");
        for (var index = 0; index < sheet.Widths.Length; index++)
            builder.Append($"<col min=\"{index + 1}\" max=\"{index + 1}\" width=\"{sheet.Widths[index]}\" customWidth=\"1\"/>");
        builder.Append("</cols><sheetData>");
        for (var rowIndex = 0; rowIndex < sheet.Rows.Count; rowIndex++)
        {
            var rowNumber = rowIndex + 1;
            builder.Append($"<row r=\"{rowNumber}\">");
            var row = sheet.Rows[rowIndex];
            for (var columnIndex = 0; columnIndex < row.Count; columnIndex++)
            {
                var cell = row[columnIndex];
                var reference = ColumnName(columnIndex + 1) + rowNumber;
                if (cell.Numeric)
                    builder.Append($"<c r=\"{reference}\" s=\"{cell.Style}\"><v>{cell.Value}</v></c>");
                else
                    builder.Append($"<c r=\"{reference}\" s=\"{cell.Style}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{Escape(cell.Value)}</t></is></c>");
            }
            builder.Append("</row>");
        }
        builder.Append("</sheetData><autoFilter ref=\"A1:");
        builder.Append(ColumnName(Math.Max(1, sheet.Rows.Max(row => row.Count))));
        builder.Append("1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");
        return builder.ToString();
    }

    private static string ContentTypesXml(int sheetCount)
    {
        var overrides = string.Concat(Enumerable.Range(1, sheetCount).Select(index =>
            $"<Override PartName=\"/xl/worksheets/sheet{index}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"));
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" + overrides + "</Types>";
    }

    private static string RootRelationshipsXml() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    private static string WorkbookXml(List<SheetData> sheets)
    {
        var sheetXml = string.Concat(sheets.Select((sheet, index) =>
            $"<sheet name=\"{Escape(sheet.Name)}\" sheetId=\"{index + 1}\" r:id=\"rId{index + 1}\"/>"));
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets>" + sheetXml + "</sheets></workbook>";
    }

    private static string WorkbookRelationshipsXml(int sheetCount)
    {
        var relationships = string.Concat(Enumerable.Range(1, sheetCount).Select(index =>
            $"<Relationship Id=\"rId{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{index}.xml\"/>"));
        relationships += $"<Relationship Id=\"rId{sheetCount + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationships + "</Relationships>";
    }

    private static string StylesXml() =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<fonts count=\"3\"><font><sz val=\"10\"/><name val=\"Malgun Gothic\"/></font><font><b/><sz val=\"16\"/><name val=\"Malgun Gothic\"/></font><font><b/><sz val=\"10\"/><color rgb=\"FFFFFFFF\"/><name val=\"Malgun Gothic\"/></font></fonts>" +
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF376C8E\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>" +
        "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=\"thin\"><color rgb=\"FFD9D9D9\"/></left><right style=\"thin\"><color rgb=\"FFD9D9D9\"/></right><top style=\"thin\"><color rgb=\"FFD9D9D9\"/></top><bottom style=\"thin\"><color rgb=\"FFD9D9D9\"/></bottom><diagonal/></border></borders>" +
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
        "<cellXfs count=\"4\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"2\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\"><alignment horizontal=\"center\"/></xf></cellXfs>" +
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";

    private static string ColumnName(int index)
    {
        var name = "";
        while (index > 0)
        {
            index--;
            name = (char)('A' + index % 26) + name;
            index /= 26;
        }
        return name;
    }

    private static string Escape(string value) => SecurityElement.Escape(value) ?? "";

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
