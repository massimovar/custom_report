#region Using directives
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using UAManagedCore;
using OpcUa = UAManagedCore.OpcUa;
using FTOptix.UI;
using FTOptix.HMIProject;
using FTOptix.NativeUI;
using FTOptix.Retentivity;
using FTOptix.CoreBase;
using FTOptix.Core;
using FTOptix.NetLogic;
using FTOptix.Alarm;
using FTOptix.EventLogger;
using FTOptix.Store;
using FTOptix.SQLiteStore;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using FTOptix.DataLogger;
// Aliases to resolve namespace conflicts
using MigraDocColor = MigraDoc.DocumentObjectModel.Color;
using MigraDocColors = MigraDoc.DocumentObjectModel.Colors;
using MigraDocTable = MigraDoc.DocumentObjectModel.Tables.Table;
using MigraDocColumn = MigraDoc.DocumentObjectModel.Tables.Column;
using MigraDocRow = MigraDoc.DocumentObjectModel.Tables.Row;
using MigraDocCell = MigraDoc.DocumentObjectModel.Tables.Cell;
using MigraDocVerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment;
using MigraDoc.DocumentObjectModel.Shapes;
using FTOptix.SerialPort;
using FTOptix.CommunicationDriver;
using FTOptix.ODBCStore;
#endregion

/// <summary>
/// Generates PDF alarm reports from EmbeddedDatabase1.AlarmsEventLogger1.
/// Uses MigraDoc/PDFsharp for PDF generation with locale-aware message columns.
/// </summary>
public class report_manager : BaseNetLogic
{
    #region Configuration - Customizable Colors

    // Header box colors (RGB format)
    private readonly MigraDocColor HeaderBackgroundColor = new MigraDocColor(0, 51, 102);      // Dark blue
    private readonly MigraDocColor HeaderTextColor = MigraDocColors.White;

    // Table header colors
    private readonly MigraDocColor TableHeaderBackgroundColor = new MigraDocColor(0, 102, 153); // Medium blue
    private readonly MigraDocColor TableHeaderTextColor = MigraDocColors.White;

    // Table row colors (alternating)
    private readonly MigraDocColor TableRowEvenColor = new MigraDocColor(240, 248, 255);        // Alice blue
    private readonly MigraDocColor TableRowOddColor = MigraDocColors.White;

    // Table border color
    private readonly MigraDocColor TableBorderColor = new MigraDocColor(180, 180, 180);         // Light gray

    #endregion

    #region Private Fields

    private Store _store;
    private string _customerLogoPath;
    private string _lastLogoPath;
    private string _machineName;

    #endregion

    public override void Start()
    {
        // Required for PDFsharp on .NET Core/.NET 5+
        AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Initialize store and configuration
        _store = Project.Current.Get<Store>("DataStores/EmbeddedDatabase1");
        _machineName = GetVariableValueOrDefault("MachineName", "Machine Name");

        // Logo paths - NOTE: PDFsharp only supports PNG, JPEG, BMP, GIF (NOT SVG)
        _customerLogoPath = GetVariableValueOrDefault("CustomerLogoPath", ResourceUri.FromProjectRelativePath("logos/logo1.png").Uri);
        _lastLogoPath = GetVariableValueOrDefault("LastLogoPath", ResourceUri.FromProjectRelativePath("logos/logo2.png").Uri);

        // Log resolved paths for debugging
        Log.Info("AlarmReport.Init", $"Customer logo: {GetAbsolutePath(_customerLogoPath)}");
        Log.Info("AlarmReport.Init", $"Last logo: {GetAbsolutePath(_lastLogoPath)}");

        Log.Info("AlarmReport.Init", "Alarm Report Manager initialized");
    }

    public override void Stop()
    {
        // Cleanup if needed
    }

    #region Public Methods (Exported)

    /// <summary>
    /// Generates PDF alarm report for the specified date range.
    /// </summary>
    [ExportMethod]
    public void GenerateAlarmsReportWithDates(DateTime startDate, DateTime endDate)
    {
        var longRunningTask = new LongRunningTask(() => GenerateReport(startDate, endDate), LogicObject);
        longRunningTask.Start();
    }

    /// <summary>Sets the customer logo path (left side of header).</summary>
    [ExportMethod]
    public void SetCustomerLogo(string logoPath)
    {
        _customerLogoPath = logoPath;
        Log.Info("AlarmReport.Config", $"Customer logo path set to: {logoPath}");
    }

    /// <summary>Sets the Last logo path (right side of header).</summary>
    [ExportMethod]
    public void SetLastLogo(string logoPath)
    {
        _lastLogoPath = logoPath;
        Log.Info("AlarmReport.Config", $"Last logo path set to: {logoPath}");
    }

    /// <summary>Sets the machine name displayed in the report header.</summary>
    [ExportMethod]
    public void SetMachineName(string machineName)
    {
        _machineName = machineName;
        Log.Info("AlarmReport.Config", $"Machine name set to: {machineName}");
    }

    /// <summary>
    /// Generates PDF datalogger report with trend chart for the specified date range.
    /// </summary>
    [ExportMethod]
    public void GenerateDataloggerReportWithDates(DateTime startDate, DateTime endDate)
    {
        var longRunningTask = new LongRunningTask(() => GenerateDataloggerReport(startDate, endDate), LogicObject);
        longRunningTask.Start();
    }

    /// <summary>
    /// Generates a recipe PDF label with dynamic height for KFI thermal printer (100mm width).
    /// Uses placeholder data for an industrial machine recipe.
    /// </summary>
    [ExportMethod]
    public void GenerateRecipePdf()
    {
        var longRunningTask = new LongRunningTask(() => GenerateRecipeLabel(), LogicObject);
        longRunningTask.Start();
    }

    #endregion

    #region Private Methods

    /// <summary>Orchestrates report generation: fetch data, create document, save PDF.</summary>
    private void GenerateReport(DateTime startDate, DateTime endDate)
    {
        try
        {
            Log.Info("AlarmReport.Generate", $"Generating report from {startDate} to {endDate}");

            var alarmData = FetchAlarmData(startDate, endDate);
            if (alarmData == null || alarmData.Count == 0)
            {
                Log.Warning("AlarmReport.Generate", "No alarm data found for the specified date range");
            }

            Document document = CreateDocument(startDate, endDate, alarmData);
            string outputPath = GetOutputPath(startDate, endDate);
            SaveDocument(document, outputPath);

            Log.Info("AlarmReport.Generate", $"Report saved: {outputPath}");
        }
        catch (Exception ex)
        {
            Log.Error("AlarmReport.Generate", $"Error generating report: {ex.Message}");
        }
    }

    /// <summary>Fetches alarm records from AlarmsEventLogger1 within the date range.</summary>
    private List<AlarmRecord> FetchAlarmData(DateTime startDate, DateTime endDate)
    {
        var alarmRecords = new List<AlarmRecord>();

        try
        {
            if (_store == null)
            {
                _store = Project.Current.Get<Store>("DataStores/EmbeddedDatabase1");
            }

            if (_store == null)
            {
                Log.Error("AlarmReport.Query", "Cannot find EmbeddedDatabase1 store");
                return alarmRecords;
            }

            // Format dates as ISO 8601 with T separator (database format)
            string startDateStr = startDate.ToString("yyyy-MM-ddTHH:mm:ss");
            string endDateStr = endDate.ToString("yyyy-MM-ddTHH:mm:ss");

            string query = $@"SELECT * FROM AlarmsEventLogger1 
                              WHERE LocalTime >= ""{startDateStr}"" AND LocalTime <= ""{endDateStr}""
                              ORDER BY LocalTime DESC";

            Object[,] resultSet;
            String[] header;
            _store.Query(query, out header, out resultSet);

            if (resultSet == null || resultSet.GetLength(0) == 0)
            {
                return alarmRecords;
            }

            Log.Info("AlarmReport.Query", $"Found {resultSet.GetLength(0)} alarms");

            // Determine locale-specific Message column (e.g., Message_it-IT, Message_en-US)
            string currentLocale = GetCurrentLocale();
            string localizedMessageColumn = $"Message_{currentLocale}";

            // Find column indices with fallback options
            int localTimeIdx = FindColumnIndex(header, "LocalTime", "Time", "Timestamp", "EventTime");
            int messageIdx = FindColumnIndex(header, localizedMessageColumn);
            if (messageIdx < 0)
            {
                messageIdx = FindColumnIndex(header, "Message", "Text", "Description", "AlarmMessage");
            }
            int deviceIdx = FindColumnIndex(header, "Device", "SourceName", "Source", "DeviceName", "AlarmName");
            int sourcePathIdx = FindColumnIndex(header, "SourcePath", "Path", "ObjectPath");

            // Parse results into AlarmRecord objects
            int rowCount = resultSet.GetLength(0);
            for (int i = 0; i < rowCount; i++)
            {
                var record = new AlarmRecord
                {
                    ActivationTime = localTimeIdx >= 0 && resultSet[i, localTimeIdx] != null
                        ? Convert.ToDateTime(resultSet[i, localTimeIdx])
                        : DateTime.MinValue,
                    AlarmText = messageIdx >= 0 && resultSet[i, messageIdx] != null
                        ? resultSet[i, messageIdx].ToString()
                        : "",
                    AssociatedDevice = deviceIdx >= 0 && resultSet[i, deviceIdx] != null
                        ? resultSet[i, deviceIdx].ToString()
                        : (sourcePathIdx >= 0 && resultSet[i, sourcePathIdx] != null
                            ? resultSet[i, sourcePathIdx].ToString()
                            : "")
                };

                alarmRecords.Add(record);
            }

            Log.Info("AlarmReport.Query", $"Fetched {alarmRecords.Count} alarm records");
        }
        catch (Exception ex)
        {
            Log.Error("AlarmReport.Query", $"Error fetching alarm data: {ex.Message}");
        }

        return alarmRecords;
    }

    /// <summary>Creates the MigraDoc document structure.</summary>
    private Document CreateDocument(DateTime startDate, DateTime endDate, List<AlarmRecord> alarmData)
    {
        Document document = new Document();
        document.Info.Title = "Alarm Report";
        document.Info.Author = "FTOptix";
        document.Info.Subject = $"Alarm Report from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}";

        DefineStyles(document);

        Section section = document.AddSection();
        section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);

        AddReportHeader(section, startDate, endDate);
        AddReportBody(section, alarmData);
        AddReportFooter(section);

        return document;
    }

    /// <summary>Defines font and paragraph styles for the document.</summary>
    private void DefineStyles(Document document)
    {
        // Modify the Normal style
        Style style = document.Styles["Normal"];
        style.Font.Name = "Arial";
        style.Font.Size = 10;

        // Create Title style
        style = document.Styles.AddStyle("Title", "Normal");
        style.Font.Size = 24;
        style.Font.Bold = true;
        style.Font.Color = HeaderTextColor;
        style.ParagraphFormat.SpaceAfter = 6;

        // Create Header style
        style = document.Styles.AddStyle("Header", "Normal");
        style.Font.Size = 11;
        style.Font.Bold = true;
        style.Font.Color = HeaderTextColor;

        // Create TableHeader style
        style = document.Styles.AddStyle("TableHeader", "Normal");
        style.Font.Size = 10;
        style.Font.Bold = true;
        style.Font.Color = TableHeaderTextColor;

        // Create TableCell style
        style = document.Styles.AddStyle("TableCell", "Normal");
        style.Font.Size = 9;
    }

    /// <summary>Adds header: logos, title, machine name, and filter dates.</summary>
    private void AddReportHeader(Section section, DateTime startDate, DateTime endDate)
    {
        // Create a table for the header layout
        MigraDocTable headerTable = section.AddTable();
        headerTable.Borders.Visible = false;
        headerTable.Format.SpaceAfter = Unit.FromCentimeter(0.5);

        // Define columns: Logo Left | Title/Info Center | Logo Right
        MigraDocColumn col1 = headerTable.AddColumn(Unit.FromCentimeter(4));
        MigraDocColumn col2 = headerTable.AddColumn(Unit.FromCentimeter(10));
        MigraDocColumn col3 = headerTable.AddColumn(Unit.FromCentimeter(4));

        // First row with logos and title
        MigraDocRow row1 = headerTable.AddRow();
        row1.Height = Unit.FromCentimeter(2.5);
        row1.VerticalAlignment = MigraDocVerticalAlignment.Center;

        // Customer logo (left)
        MigraDocCell logoLeftCell = row1.Cells[0];
        string customerLogoFullPath = GetAbsolutePath(_customerLogoPath);
        Log.Info("AlarmReport.Header", $"Customer logo path: {customerLogoFullPath}, Exists: {File.Exists(customerLogoFullPath)}");

        if (!string.IsNullOrEmpty(_customerLogoPath) && File.Exists(customerLogoFullPath))
        {
            try
            {
                var image = logoLeftCell.AddImage(customerLogoFullPath);
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("AlarmReport.Header", $"Could not load customer logo: {ex.Message}");
            }
        }
        else
        {
            Log.Warning("AlarmReport.Header", $"Customer logo not found at: {customerLogoFullPath}");
        }

        // Title and info (center)
        MigraDocCell centerCell = row1.Cells[1];
        centerCell.Format.Alignment = ParagraphAlignment.Center;
        centerCell.Shading.Color = HeaderBackgroundColor;
        centerCell.VerticalAlignment = MigraDocVerticalAlignment.Center;

        Paragraph titleParagraph = centerCell.AddParagraph("Alarm Report");
        titleParagraph.Style = "Title";
        titleParagraph.Format.Alignment = ParagraphAlignment.Center;

        // Last logo (right)
        MigraDocCell logoRightCell = row1.Cells[2];
        logoRightCell.Format.Alignment = ParagraphAlignment.Right;
        string lastLogoFullPath = GetAbsolutePath(_lastLogoPath);
        Log.Info("AlarmReport.Header", $"Last logo path: {lastLogoFullPath}, Exists: {File.Exists(lastLogoFullPath)}");

        if (!string.IsNullOrEmpty(_lastLogoPath) && File.Exists(lastLogoFullPath))
        {
            try
            {
                var image = logoRightCell.AddImage(lastLogoFullPath);
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("AlarmReport.Header", $"Could not load Last logo: {ex.Message}");
            }
        }
        else
        {
            Log.Warning("AlarmReport.Header", $"Last logo not found at: {lastLogoFullPath}");
        }

        // Add info box below header
        MigraDocTable infoTable = section.AddTable();
        infoTable.Borders.Width = 0.5;
        infoTable.Borders.Color = TableBorderColor;
        infoTable.Format.SpaceAfter = Unit.FromCentimeter(0.5);

        MigraDocColumn infoCol1 = infoTable.AddColumn(Unit.FromCentimeter(4));
        MigraDocColumn infoCol2 = infoTable.AddColumn(Unit.FromCentimeter(14));

        // Machine Name row
        MigraDocRow machineRow = infoTable.AddRow();
        machineRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        machineRow.Cells[0].AddParagraph("Machine Name:").Style = "TableHeader";
        machineRow.Cells[1].AddParagraph(_machineName);

        // Start Date row
        MigraDocRow startRow = infoTable.AddRow();
        startRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        startRow.Cells[0].AddParagraph("Filter Start Date:").Style = "TableHeader";
        startRow.Cells[1].AddParagraph(startDate.ToString("yyyy-MM-dd HH:mm:ss"));

        // End Date row
        MigraDocRow endRow = infoTable.AddRow();
        endRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        endRow.Cells[0].AddParagraph("Filter End Date:").Style = "TableHeader";
        endRow.Cells[1].AddParagraph(endDate.ToString("yyyy-MM-dd HH:mm:ss"));

        // Add some space after header
        section.AddParagraph().Format.SpaceAfter = Unit.FromCentimeter(0.5);
    }

    /// <summary>Adds the alarm data table with header and data rows.</summary>
    private void AddReportBody(Section section, List<AlarmRecord> alarmData)
    {
        // Section title
        Paragraph bodyTitle = section.AddParagraph("Alarm History");
        bodyTitle.Format.Font.Size = 14;
        bodyTitle.Format.Font.Bold = true;
        bodyTitle.Format.SpaceAfter = Unit.FromCentimeter(0.3);

        // Create the alarm table
        MigraDocTable alarmTable = section.AddTable();
        alarmTable.Borders.Width = 0.5;
        alarmTable.Borders.Color = TableBorderColor;
        alarmTable.Format.Font.Size = 9;

        // Define columns
        MigraDocColumn colTime = alarmTable.AddColumn(Unit.FromCentimeter(4));       // Activation Date/Time
        MigraDocColumn colText = alarmTable.AddColumn(Unit.FromCentimeter(9));       // Alarm Text
        MigraDocColumn colDevice = alarmTable.AddColumn(Unit.FromCentimeter(5));     // Associated Device

        // Header row
        MigraDocRow headerRow = alarmTable.AddRow();
        headerRow.HeadingFormat = true;
        headerRow.Format.Font.Bold = true;
        headerRow.Shading.Color = TableHeaderBackgroundColor;
        headerRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
        headerRow.Height = Unit.FromCentimeter(0.8);

        MigraDocCell headerCell1 = headerRow.Cells[0];
        headerCell1.AddParagraph("Activation Date/Time").Style = "TableHeader";
        headerCell1.Format.Alignment = ParagraphAlignment.Center;

        MigraDocCell headerCell2 = headerRow.Cells[1];
        headerCell2.AddParagraph("Alarm Text").Style = "TableHeader";
        headerCell2.Format.Alignment = ParagraphAlignment.Center;

        MigraDocCell headerCell3 = headerRow.Cells[2];
        headerCell3.AddParagraph("Associated Device").Style = "TableHeader";
        headerCell3.Format.Alignment = ParagraphAlignment.Center;

        // Data rows
        if (alarmData != null && alarmData.Count > 0)
        {
            for (int i = 0; i < alarmData.Count; i++)
            {
                var alarm = alarmData[i];
                MigraDocRow dataRow = alarmTable.AddRow();
                dataRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
                dataRow.Height = Unit.FromCentimeter(0.6);

                // Alternating row colors
                MigraDocColor rowColor = (i % 2 == 0) ? TableRowEvenColor : TableRowOddColor;
                dataRow.Cells[0].Shading.Color = rowColor;
                dataRow.Cells[1].Shading.Color = rowColor;
                dataRow.Cells[2].Shading.Color = rowColor;

                // Activation Time
                dataRow.Cells[0].AddParagraph(alarm.ActivationTime.ToString("yyyy-MM-dd HH:mm:ss"));
                dataRow.Cells[0].Format.Alignment = ParagraphAlignment.Center;

                // Alarm Text
                dataRow.Cells[1].AddParagraph(alarm.AlarmText ?? "");
                dataRow.Cells[1].Format.Alignment = ParagraphAlignment.Left;

                // Associated Device
                dataRow.Cells[2].AddParagraph(alarm.AssociatedDevice ?? "");
                dataRow.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            }
        }
        else
        {
            // No data row
            MigraDocRow emptyRow = alarmTable.AddRow();
            emptyRow.Cells[0].MergeRight = 2;
            emptyRow.Cells[0].AddParagraph("No alarms found for the specified date range");
            emptyRow.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            emptyRow.Cells[0].Format.Font.Italic = true;
            emptyRow.Height = Unit.FromCentimeter(1);
            emptyRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
        }

        // Add record count
        Paragraph countParagraph = section.AddParagraph();
        countParagraph.Format.SpaceBefore = Unit.FromCentimeter(0.3);
        countParagraph.Format.Font.Size = 9;
        countParagraph.Format.Font.Italic = true;
        countParagraph.AddText($"Total records: {alarmData?.Count ?? 0}");
    }

    /// <summary>Adds footer with page numbers and generation timestamp.</summary>
    private void AddReportFooter(Section section)
    {
        Paragraph footer = section.Footers.Primary.AddParagraph();
        footer.Format.Font.Size = 8;
        footer.Format.Alignment = ParagraphAlignment.Center;
        footer.AddText($"Generated on {DateTime.Now:yyyy-MM-dd HH:mm:ss} - Page ");
        footer.AddPageField();
        footer.AddText(" of ");
        footer.AddNumPagesField();
    }

    /// <summary>Renders the MigraDoc document and saves it as PDF.</summary>
    private void SaveDocument(Document document, string outputPath)
    {
        try
        {
            // Ensure output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Render the document
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();

            // Save the PDF
            renderer.PdfDocument.Save(outputPath);

            Log.Info("AlarmReport.Save", $"PDF saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Log.Error("AlarmReport.Save", $"Error saving PDF: {ex.Message}");
            throw;
        }
    }

    #region Datalogger Report Methods

    /// <summary>Orchestrates datalogger report generation: fetch data, create chart, create document, save PDF.</summary>
    private void GenerateDataloggerReport(DateTime startDate, DateTime endDate)
    {
        try
        {
            Log.Info("DataloggerReport.Generate", $"Generating datalogger report from {startDate} to {endDate}");

            var dataloggerData = FetchDataloggerData(startDate, endDate);
            if (dataloggerData == null || dataloggerData.Count == 0)
            {
                Log.Warning("DataloggerReport.Generate", "No datalogger data found for the specified date range");
            }

            // Generate trend chart as separate PDF
            string chartPdfPath = GenerateTrendChart(dataloggerData, startDate, endDate);

            // Create the main document (without chart - we'll merge later)
            Document document = CreateDataloggerDocument(startDate, endDate, dataloggerData, chartPdfPath);
            string outputPath = GetDataloggerOutputPath(startDate, endDate);
            
            // Save with chart PDF merged
            SaveDataloggerDocument(document, outputPath, chartPdfPath);

            // Clean up temporary chart PDF
            if (File.Exists(chartPdfPath))
            {
                try { File.Delete(chartPdfPath); } catch { }
            }

            Log.Info("DataloggerReport.Generate", $"Report saved: {outputPath}");
        }
        catch (Exception ex)
        {
            Log.Error("DataloggerReport.Generate", $"Error generating datalogger report: {ex.Message}");
        }
    }

    /// <summary>Fetches datalogger records from DataLogger1 within the date range.</summary>
    private List<DataloggerRecord> FetchDataloggerData(DateTime startDate, DateTime endDate)
    {
        var dataRecords = new List<DataloggerRecord>();

        try
        {
            if (_store == null)
            {
                _store = Project.Current.Get<Store>("DataStores/EmbeddedDatabase1");
            }

            if (_store == null)
            {
                Log.Error("DataloggerReport.Query", "Cannot find EmbeddedDatabase1 store");
                return dataRecords;
            }

            // Format dates as ISO 8601 with T separator (database format)
            string startDateStr = startDate.ToString("yyyy-MM-ddTHH:mm:ss");
            string endDateStr = endDate.ToString("yyyy-MM-ddTHH:mm:ss");

            string query = $@"SELECT LocalTimestamp, myVarToLog FROM DataLogger1 
                              WHERE LocalTimestamp >= ""{startDateStr}"" AND LocalTimestamp <= ""{endDateStr}""
                              ORDER BY LocalTimestamp ASC";

            Object[,] resultSet;
            String[] header;
            _store.Query(query, out header, out resultSet);

            if (resultSet == null || resultSet.GetLength(0) == 0)
            {
                return dataRecords;
            }

            Log.Info("DataloggerReport.Query", $"Found {resultSet.GetLength(0)} datalogger records");

            // Find column indices
            int timestampIdx = FindColumnIndex(header, "LocalTimestamp", "Timestamp", "Time");
            int valueIdx = FindColumnIndex(header, "myVarToLog", "Value");

            // Parse results into DataloggerRecord objects
            int rowCount = resultSet.GetLength(0);
            for (int i = 0; i < rowCount; i++)
            {
                var record = new DataloggerRecord
                {
                    Timestamp = timestampIdx >= 0 && resultSet[i, timestampIdx] != null
                        ? Convert.ToDateTime(resultSet[i, timestampIdx])
                        : DateTime.MinValue,
                    Value = valueIdx >= 0 && resultSet[i, valueIdx] != null
                        ? Convert.ToDouble(resultSet[i, valueIdx])
                        : 0.0
                };

                dataRecords.Add(record);
            }

            Log.Info("DataloggerReport.Query", $"Fetched {dataRecords.Count} datalogger records");
        }
        catch (Exception ex)
        {
            Log.Error("DataloggerReport.Query", $"Error fetching datalogger data: {ex.Message}");
        }

        return dataRecords;
    }

    /// <summary>Generates a trend chart using pure PDFsharp on an A4 portrait page.</summary>
    private string GenerateTrendChart(List<DataloggerRecord> data, DateTime startDate, DateTime endDate)
    {
        // A4 portrait dimensions in points (1 point = 1/72 inch)
        // A4 = 210mm x 297mm = 595.28 x 841.89 points
        double pageWidth = 595;
        double pageHeight = 842;
        
        // Page margins (same as MigraDoc: 1.5cm = ~42.5 points)
        double pageMargin = 42.5;
        
        // Content area width
        double contentWidth = pageWidth - (2 * pageMargin);

        // Create a temporary PDF document for the chart
        string tempDir = Path.GetTempPath();
        string chartPdfPath = Path.Combine(tempDir, $"DataloggerTrend_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
        
        using (PdfDocument chartDoc = new PdfDocument())
        {
            PdfPage page = chartDoc.AddPage();
            page.Width = XUnit.FromPoint(pageWidth);
            page.Height = XUnit.FromPoint(pageHeight);

            using (XGraphics gfx = XGraphics.FromPdfPage(page))
            {
                // Colors (matching MigraDoc styles)
                XColor headerBgColor = XColor.FromArgb(0, 51, 102);      // HeaderBackgroundColor
                XColor tableHeaderBgColor = XColor.FromArgb(0, 102, 153); // TableHeaderBackgroundColor
                XColor axisColor = XColor.FromArgb(50, 50, 50);
                XColor gridColor = XColor.FromArgb(200, 200, 200);
                XColor lineColor = XColor.FromArgb(0, 102, 153);
                XColor markerColor = XColor.FromArgb(0, 51, 102);

                // Pens and brushes
                XPen axisPen = new XPen(axisColor, 1.5);
                XPen gridPen = new XPen(gridColor, 0.5);
                XPen linePen = new XPen(lineColor, 2);
                XBrush markerBrush = new XSolidBrush(markerColor);
                XBrush textBrush = new XSolidBrush(axisColor);
                XBrush headerBgBrush = new XSolidBrush(headerBgColor);
                XBrush tableHeaderBrush = new XSolidBrush(tableHeaderBgColor);
                XPen borderPen = new XPen(gridColor, 0.5);

                // Fonts
                XFont pageTitleFont = new XFont("Arial", 24, XFontStyle.Bold);
                XFont sectionTitleFont = new XFont("Arial", 14, XFontStyle.Bold);
                XFont axisLabelFont = new XFont("Arial", 10, XFontStyle.Bold);
                XFont tickFont = new XFont("Arial", 8);
                XFont infoFont = new XFont("Arial", 10);
                XFont infoLabelFont = new XFont("Arial", 10, XFontStyle.Bold);

                // Background
                gfx.DrawRectangle(XBrushes.White, 0, 0, pageWidth, pageHeight);

                double currentY = pageMargin;

                // === PAGE HEADER (Title bar) ===
                double headerHeight = 50;
                gfx.DrawRectangle(headerBgBrush, pageMargin, currentY, contentWidth, headerHeight);
                gfx.DrawString("Datalogger Report", pageTitleFont, XBrushes.White,
                    new XRect(pageMargin, currentY, contentWidth, headerHeight), XStringFormats.Center);
                currentY += headerHeight + 15;

                // === INFO BOX (Machine Name and Date Filters) ===
                double labelWidth = 120;
                double rowHeight = 25;
                
                // Machine Name row
                gfx.DrawRectangle(tableHeaderBrush, pageMargin, currentY, labelWidth, rowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, pageMargin + labelWidth, currentY, contentWidth - labelWidth, rowHeight);
                gfx.DrawString("Machine Name:", infoLabelFont, XBrushes.White, pageMargin + 5, currentY + 7);
                gfx.DrawString(_machineName ?? "Machine Name", infoFont, textBrush, pageMargin + labelWidth + 5, currentY + 7);
                currentY += rowHeight;
                
                // Start Date row
                gfx.DrawRectangle(tableHeaderBrush, pageMargin, currentY, labelWidth, rowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, pageMargin + labelWidth, currentY, contentWidth - labelWidth, rowHeight);
                gfx.DrawString("Filter Start Date:", infoLabelFont, XBrushes.White, pageMargin + 5, currentY + 7);
                gfx.DrawString(startDate.ToString("yyyy-MM-dd HH:mm:ss"), infoFont, textBrush, pageMargin + labelWidth + 5, currentY + 7);
                currentY += rowHeight;
                
                // End Date row
                gfx.DrawRectangle(tableHeaderBrush, pageMargin, currentY, labelWidth, rowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, pageMargin + labelWidth, currentY, contentWidth - labelWidth, rowHeight);
                gfx.DrawString("Filter End Date:", infoLabelFont, XBrushes.White, pageMargin + 5, currentY + 7);
                gfx.DrawString(endDate.ToString("yyyy-MM-dd HH:mm:ss"), infoFont, textBrush, pageMargin + labelWidth + 5, currentY + 7);
                currentY += rowHeight + 20;

                // === TREND CHART SECTION ===
                // Section title
                gfx.DrawString("Trend Chart", sectionTitleFont, textBrush, pageMargin, currentY);
                currentY += 20;

                // Chart area dimensions
                double chartAreaX = pageMargin;
                double chartAreaY = currentY;
                double chartAreaWidth = contentWidth;
                double chartAreaHeight = 380;  // Chart height
                
                // Plot margins within chart area
                double marginLeft = 60;
                double marginRight = 20;
                double marginTop = 20;
                double marginBottom = 50;
                
                double plotWidth = chartAreaWidth - marginLeft - marginRight;
                double plotHeight = chartAreaHeight - marginTop - marginBottom;

                // Chart background
                gfx.DrawRectangle(XBrushes.White, chartAreaX, chartAreaY, chartAreaWidth, chartAreaHeight);

                // Calculate data bounds
                double minValue = 0, maxValue = 100;
                if (data != null && data.Count > 0)
                {
                    minValue = data[0].Value;
                    maxValue = data[0].Value;
                    foreach (var record in data)
                    {
                        if (record.Value < minValue) minValue = record.Value;
                        if (record.Value > maxValue) maxValue = record.Value;
                    }
                    // Add 10% padding
                    double range = maxValue - minValue;
                    if (range < 0.001) range = 1;
                    minValue -= range * 0.1;
                    maxValue += range * 0.1;
                }

                double timeRange = (endDate - startDate).TotalSeconds;
                if (timeRange < 1) timeRange = 1;

                // Plot area origin
                double plotX = chartAreaX + marginLeft;
                double plotY = chartAreaY + marginTop;

                // Draw grid lines (horizontal)
                int numHorizontalLines = 5;
                for (int i = 0; i <= numHorizontalLines; i++)
                {
                    double y = plotY + (plotHeight * i / numHorizontalLines);
                    gfx.DrawLine(gridPen, plotX, y, plotX + plotWidth, y);
                    
                    // Y-axis labels
                    double value = maxValue - ((maxValue - minValue) * i / numHorizontalLines);
                    gfx.DrawString(value.ToString("F1"), tickFont, textBrush, 
                        plotX - 5, y, new XStringFormat { Alignment = XStringAlignment.Far, LineAlignment = XLineAlignment.Center });
                }

                // Draw grid lines (vertical) and X-axis labels
                int numVerticalLines = 6;
                for (int i = 0; i <= numVerticalLines; i++)
                {
                    double x = plotX + (plotWidth * i / numVerticalLines);
                    gfx.DrawLine(gridPen, x, plotY, x, plotY + plotHeight);
                    
                    // X-axis labels
                    DateTime time = startDate.AddSeconds(timeRange * i / numVerticalLines);
                    string timeLabel = time.ToString("MM-dd HH:mm");
                    gfx.DrawString(timeLabel, tickFont, textBrush, 
                        x, plotY + plotHeight + 12, new XStringFormat { Alignment = XStringAlignment.Center });
                }

                // Draw axes
                gfx.DrawLine(axisPen, plotX, plotY, plotX, plotY + plotHeight); // Y-axis
                gfx.DrawLine(axisPen, plotX, plotY + plotHeight, plotX + plotWidth, plotY + plotHeight); // X-axis

                // Draw axis titles
                gfx.DrawString("Time", axisLabelFont, textBrush, 
                    plotX + plotWidth / 2, plotY + plotHeight + 35, XStringFormats.TopCenter);
                
                // Y-axis title (rotated)
                XGraphicsState state = gfx.Save();
                gfx.RotateAtTransform(-90, new XPoint(chartAreaX + 15, plotY + plotHeight / 2));
                gfx.DrawString("myVarToLog", axisLabelFont, textBrush, 
                    chartAreaX + 15, plotY + plotHeight / 2, XStringFormats.Center);
                gfx.Restore(state);

                // Draw data line and markers
                if (data != null && data.Count > 0)
                {
                    List<XPoint> points = new List<XPoint>();
                    
                    foreach (var record in data)
                    {
                        double xRatio = (record.Timestamp - startDate).TotalSeconds / timeRange;
                        double yRatio = (record.Value - minValue) / (maxValue - minValue);
                        
                        double x = plotX + (xRatio * plotWidth);
                        double y = plotY + plotHeight - (yRatio * plotHeight);
                        
                        // Clamp to plot area
                        x = Math.Max(plotX, Math.Min(plotX + plotWidth, x));
                        y = Math.Max(plotY, Math.Min(plotY + plotHeight, y));
                        
                        points.Add(new XPoint(x, y));
                    }

                    // Draw lines between points
                    if (points.Count > 1)
                    {
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                            gfx.DrawLine(linePen, points[i], points[i + 1]);
                        }
                    }

                    // Draw markers (only if not too many points)
                    if (points.Count <= 100)
                    {
                        double markerSize = 4;
                        foreach (var point in points)
                        {
                            gfx.DrawEllipse(markerBrush, point.X - markerSize/2, point.Y - markerSize/2, markerSize, markerSize);
                        }
                    }
                }
                else
                {
                    // No data message
                    gfx.DrawString("No data available", axisLabelFont, textBrush,
                        plotX + plotWidth / 2, plotY + plotHeight / 2, XStringFormats.Center);
                }

                // Draw plot area border
                gfx.DrawRectangle(axisPen, plotX, plotY, plotWidth, plotHeight);

                // === FOOTER ===
                string footerText = $"Generated on {DateTime.Now:yyyy-MM-dd HH:mm:ss} - Page 1";
                gfx.DrawString(footerText, tickFont, textBrush,
                    new XRect(0, pageHeight - 30, pageWidth, 20), XStringFormats.Center);
            }

            chartDoc.Save(chartPdfPath);
        }

        Log.Info("DataloggerReport.Chart", $"Chart PDF generated: {chartPdfPath}");
        return chartPdfPath;
    }

    /// <summary>Creates the MigraDoc document structure for datalogger report.</summary>
    private Document CreateDataloggerDocument(DateTime startDate, DateTime endDate, List<DataloggerRecord> data, string chartImagePath)
    {
        Document document = new Document();
        document.Info.Title = "Datalogger Report";
        document.Info.Author = "FTOptix";
        document.Info.Subject = $"Datalogger Report from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}";

        DefineStyles(document);

        Section section = document.AddSection();
        section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);

        AddDataloggerReportHeader(section, startDate, endDate);
        AddDataloggerTrendSection(section, chartImagePath);
        AddDataloggerDataTable(section, data);
        AddReportFooter(section);

        return document;
    }

    /// <summary>Adds header for datalogger report: logos, title, machine name, and filter dates.</summary>
    private void AddDataloggerReportHeader(Section section, DateTime startDate, DateTime endDate)
    {
        // Create a table for the header layout
        MigraDocTable headerTable = section.AddTable();
        headerTable.Borders.Visible = false;
        headerTable.Format.SpaceAfter = Unit.FromCentimeter(0.5);

        // Define columns: Logo Left | Title/Info Center | Logo Right
        MigraDocColumn col1 = headerTable.AddColumn(Unit.FromCentimeter(4));
        MigraDocColumn col2 = headerTable.AddColumn(Unit.FromCentimeter(10));
        MigraDocColumn col3 = headerTable.AddColumn(Unit.FromCentimeter(4));

        // First row with logos and title
        MigraDocRow row1 = headerTable.AddRow();
        row1.Height = Unit.FromCentimeter(2.5);
        row1.VerticalAlignment = MigraDocVerticalAlignment.Center;

        // Customer logo (left)
        MigraDocCell logoLeftCell = row1.Cells[0];
        string customerLogoFullPath = GetAbsolutePath(_customerLogoPath);

        if (!string.IsNullOrEmpty(_customerLogoPath) && File.Exists(customerLogoFullPath))
        {
            try
            {
                var image = logoLeftCell.AddImage(customerLogoFullPath);
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("DataloggerReport.Header", $"Could not load customer logo: {ex.Message}");
            }
        }

        // Title and info (center)
        MigraDocCell centerCell = row1.Cells[1];
        centerCell.Format.Alignment = ParagraphAlignment.Center;
        centerCell.Shading.Color = HeaderBackgroundColor;
        centerCell.VerticalAlignment = MigraDocVerticalAlignment.Center;

        Paragraph titleParagraph = centerCell.AddParagraph("Datalogger Report");
        titleParagraph.Style = "Title";
        titleParagraph.Format.Alignment = ParagraphAlignment.Center;

        // Last logo (right)
        MigraDocCell logoRightCell = row1.Cells[2];
        logoRightCell.Format.Alignment = ParagraphAlignment.Right;
        string lastLogoFullPath = GetAbsolutePath(_lastLogoPath);

        if (!string.IsNullOrEmpty(_lastLogoPath) && File.Exists(lastLogoFullPath))
        {
            try
            {
                var image = logoRightCell.AddImage(lastLogoFullPath);
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("DataloggerReport.Header", $"Could not load Last logo: {ex.Message}");
            }
        }

        // Add info box below header
        MigraDocTable infoTable = section.AddTable();
        infoTable.Borders.Width = 0.5;
        infoTable.Borders.Color = TableBorderColor;
        infoTable.Format.SpaceAfter = Unit.FromCentimeter(0.5);

        MigraDocColumn infoCol1 = infoTable.AddColumn(Unit.FromCentimeter(4));
        MigraDocColumn infoCol2 = infoTable.AddColumn(Unit.FromCentimeter(14));

        // Machine Name row
        MigraDocRow machineRow = infoTable.AddRow();
        machineRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        machineRow.Cells[0].AddParagraph("Machine Name:").Style = "TableHeader";
        machineRow.Cells[1].AddParagraph(_machineName);

        // Start Date row
        MigraDocRow startRow = infoTable.AddRow();
        startRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        startRow.Cells[0].AddParagraph("Filter Start Date:").Style = "TableHeader";
        startRow.Cells[1].AddParagraph(startDate.ToString("yyyy-MM-dd HH:mm:ss"));

        // End Date row
        MigraDocRow endRow = infoTable.AddRow();
        endRow.Cells[0].Shading.Color = TableHeaderBackgroundColor;
        endRow.Cells[0].AddParagraph("Filter End Date:").Style = "TableHeader";
        endRow.Cells[1].AddParagraph(endDate.ToString("yyyy-MM-dd HH:mm:ss"));

        // Add some space after header
        section.AddParagraph().Format.SpaceAfter = Unit.FromCentimeter(0.5);
    }

    /// <summary>Adds the trend chart section to the document (chart is on first page).</summary>
    private void AddDataloggerTrendSection(Section section, string chartPdfPath)
    {
        // Note: The chart is rendered on the FIRST page of the PDF (merged at index 0)
        // This section just adds a reference note pointing to the first page
        Paragraph trendTitle = section.AddParagraph("Trend Chart");
        trendTitle.Format.Font.Size = 14;
        trendTitle.Format.Font.Bold = true;
        trendTitle.Format.SpaceAfter = Unit.FromCentimeter(0.3);

        if (!string.IsNullOrEmpty(chartPdfPath) && File.Exists(chartPdfPath))
        {
            Paragraph chartNote = section.AddParagraph("(See chart on page 1)");
            chartNote.Format.Font.Italic = true;
            chartNote.Format.Alignment = ParagraphAlignment.Center;
        }
        else
        {
            Paragraph noChartParagraph = section.AddParagraph("[No chart data available]");
            noChartParagraph.Format.Font.Italic = true;
            noChartParagraph.Format.Alignment = ParagraphAlignment.Center;
        }

        section.AddParagraph().Format.SpaceAfter = Unit.FromCentimeter(0.5);
    }

    /// <summary>Adds the datalogger data table with timestamp and value columns.</summary>
    private void AddDataloggerDataTable(Section section, List<DataloggerRecord> data)
    {
        // Section title
        Paragraph bodyTitle = section.AddParagraph("Data Table");
        bodyTitle.Format.Font.Size = 14;
        bodyTitle.Format.Font.Bold = true;
        bodyTitle.Format.SpaceAfter = Unit.FromCentimeter(0.3);

        // Create the data table
        MigraDocTable dataTable = section.AddTable();
        dataTable.Borders.Width = 0.5;
        dataTable.Borders.Color = TableBorderColor;
        dataTable.Format.Font.Size = 9;

        // Define columns
        MigraDocColumn colTime = dataTable.AddColumn(Unit.FromCentimeter(6));    // Timestamp
        MigraDocColumn colValue = dataTable.AddColumn(Unit.FromCentimeter(6));   // Value

        // Header row
        MigraDocRow headerRow = dataTable.AddRow();
        headerRow.HeadingFormat = true;
        headerRow.Format.Font.Bold = true;
        headerRow.Shading.Color = TableHeaderBackgroundColor;
        headerRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
        headerRow.Height = Unit.FromCentimeter(0.8);

        MigraDocCell headerCell1 = headerRow.Cells[0];
        headerCell1.AddParagraph("LocalTimestamp").Style = "TableHeader";
        headerCell1.Format.Alignment = ParagraphAlignment.Center;

        MigraDocCell headerCell2 = headerRow.Cells[1];
        headerCell2.AddParagraph("myVarToLog").Style = "TableHeader";
        headerCell2.Format.Alignment = ParagraphAlignment.Center;

        // Data rows
        if (data != null && data.Count > 0)
        {
            for (int i = 0; i < data.Count; i++)
            {
                var record = data[i];
                MigraDocRow dataRow = dataTable.AddRow();
                dataRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
                dataRow.Height = Unit.FromCentimeter(0.5);

                // Alternating row colors
                MigraDocColor rowColor = (i % 2 == 0) ? TableRowEvenColor : TableRowOddColor;
                dataRow.Cells[0].Shading.Color = rowColor;
                dataRow.Cells[1].Shading.Color = rowColor;

                // Timestamp
                dataRow.Cells[0].AddParagraph(record.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                dataRow.Cells[0].Format.Alignment = ParagraphAlignment.Center;

                // Value
                dataRow.Cells[1].AddParagraph(record.Value.ToString("F3"));
                dataRow.Cells[1].Format.Alignment = ParagraphAlignment.Center;
            }
        }
        else
        {
            // No data row
            MigraDocRow emptyRow = dataTable.AddRow();
            emptyRow.Cells[0].MergeRight = 1;
            emptyRow.Cells[0].AddParagraph("No data found for the specified date range");
            emptyRow.Cells[0].Format.Alignment = ParagraphAlignment.Center;
            emptyRow.Cells[0].Format.Font.Italic = true;
            emptyRow.Height = Unit.FromCentimeter(1);
            emptyRow.VerticalAlignment = MigraDocVerticalAlignment.Center;
        }

        // Add record count
        Paragraph countParagraph = section.AddParagraph();
        countParagraph.Format.SpaceBefore = Unit.FromCentimeter(0.3);
        countParagraph.Format.Font.Size = 9;
        countParagraph.Format.Font.Italic = true;
        countParagraph.AddText($"Total records: {data?.Count ?? 0}");
    }

    /// <summary>Saves the datalogger document and merges the chart PDF page.</summary>
    private void SaveDataloggerDocument(Document document, string outputPath, string chartPdfPath)
    {
        try
        {
            // Ensure output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Render the MigraDoc document
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();

            // If we have a chart PDF, merge it into the document
            if (!string.IsNullOrEmpty(chartPdfPath) && File.Exists(chartPdfPath))
            {
                try
                {
                    using (PdfDocument chartDoc = PdfSharp.Pdf.IO.PdfReader.Open(chartPdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                    {
                        // Insert chart page as the FIRST page
                        if (chartDoc.PageCount > 0)
                        {
                            PdfPage chartPage = chartDoc.Pages[0];
                            renderer.PdfDocument.InsertPage(0, chartPage);
                            Log.Info("DataloggerReport.Save", "Chart page inserted as first page");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning("DataloggerReport.Save", $"Could not merge chart PDF: {ex.Message}");
                }
            }

            // Save the final PDF
            renderer.PdfDocument.Save(outputPath);

            Log.Info("DataloggerReport.Save", $"PDF saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Log.Error("DataloggerReport.Save", $"Error saving PDF: {ex.Message}");
            throw;
        }
    }

    /// <summary>Returns output path for datalogger report.</summary>
    private string GetDataloggerOutputPath(DateTime startDate, DateTime endDate)
    {
        string customPath = GetVariableValueOrDefault("DataloggerOutputPath", "");
        if (!string.IsNullOrEmpty(customPath))
            return customPath;

        string projectDir = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        string reportsDir = Path.Combine(projectDir, "Reports");
        string fileName = $"DataloggerReport_From_{startDate:yyyy-MM-dd_HHmmss}_To_{endDate:yyyy-MM-dd_HHmmss}.pdf";
        return Path.Combine(reportsDir, fileName);
    }

    #endregion

    /// <summary>Returns output path: custom OutputPath variable or ProjectFiles/Reports/.</summary>
    private string GetOutputPath(DateTime startDate, DateTime endDate)
    {
        string customPath = GetVariableValueOrDefault("OutputPath", "");
        if (!string.IsNullOrEmpty(customPath))
            return customPath;

        string projectDir = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        string reportsDir = Path.Combine(projectDir, "Reports");
        string fileName = $"AlarmReport_From_{startDate:yyyy-MM-dd_HHmmss}_To_{endDate:yyyy-MM-dd_HHmmss}.pdf";
        return Path.Combine(reportsDir, fileName);
    }

    /// <summary>Converts project-relative path or URI to absolute file path.</summary>
    private string GetAbsolutePath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return "";

        // Handle URI format (file:///C:/path/...) - convert to regular path
        if (path.StartsWith("file:///"))
        {
            path = path.Substring(8).Replace("/", "\\");
            Log.Info("AlarmReport.Path", $"Converted URI to path: {path}");
            return path;
        }

        // If it's already an absolute path, return it
        if (Path.IsPathRooted(path))
            return path;

        // Otherwise, treat as project-relative and convert URI to path
        string projectUri = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        string projectDir = projectUri.StartsWith("file:///")
            ? projectUri.Substring(8).Replace("/", "\\")
            : projectUri;
        return Path.Combine(projectDir, path);
    }

    /// <summary>Gets NetLogic variable value or returns default.</summary>
    private string GetVariableValueOrDefault(string variableName, string defaultValue)
    {
        try
        {
            var variable = LogicObject.GetVariable(variableName);
            if (variable != null && variable.Value != null)
            {
                return variable.Value.ToString();
            }
        }
        catch
        {
            // Variable doesn't exist, use default
        }

        return defaultValue;
    }

    /// <summary>Finds column index by trying possible names (case-insensitive).</summary>
    private int FindColumnIndex(string[] header, params string[] possibleNames)
    {
        if (header == null || possibleNames == null)
            return -1;

        foreach (var name in possibleNames)
        {
            for (int i = 0; i < header.Length; i++)
            {
                if (string.Equals(header[i], name, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
        }
        return -1;
    }

    /// <summary>Gets Session.ActualLocaleId for locale-aware Message column selection.</summary>
    private string GetCurrentLocale()
    {
        try
        {
            string localeId = Session.ActualLocaleId;
            if (!string.IsNullOrEmpty(localeId))
                return localeId;
        }
        catch { }

        return "en-US";
    }

    #endregion

    #region Recipe PDF Generation

    /// <summary>Generates a recipe label PDF with dynamic height for thermal printing.</summary>
    private void GenerateRecipeLabel()
    {
        try
        {
            Log.Info("RecipeLabel.Generate", "Generating recipe label PDF");

            // Create placeholder recipe data
            RecipeData recipe = CreatePlaceholderRecipe();

            // Generate the PDF
            string outputPath = GenerateRecipePdfDocument(recipe);

            Log.Info("RecipeLabel.Generate", $"Recipe label saved: {outputPath}");
        }
        catch (Exception ex)
        {
            Log.Error("RecipeLabel.Generate", $"Error generating recipe label: {ex.Message}");
        }
    }

    /// <summary>Creates placeholder recipe data for an industrial machine.</summary>
    private RecipeData CreatePlaceholderRecipe()
    {
        return new RecipeData
        {
            RecipeName = "FORMULA-2024-A",
            RecipeCode = "FRM-2024-001",
            BatchNumber = "BATCH-20260122-001",
            ProductCode = "PRD-45892",
            CreatedDate = DateTime.Now,
            OperatorName = "Operator 1",
            MachineName = _machineName ?? "CNC Machine #1",
            
            // Process parameters
            Parameters = new List<RecipeParameter>
            {
                new RecipeParameter { Name = "Temperature", Value = "185.5", Unit = "°C", SetPoint = "185.0" },
                new RecipeParameter { Name = "Pressure", Value = "4.2", Unit = "bar", SetPoint = "4.0" },
                new RecipeParameter { Name = "Speed", Value = "1500", Unit = "RPM", SetPoint = "1500" },
                new RecipeParameter { Name = "Flow Rate", Value = "12.8", Unit = "L/min", SetPoint = "12.5" },
                new RecipeParameter { Name = "Cycle Time", Value = "45", Unit = "sec", SetPoint = "45" },
                new RecipeParameter { Name = "Dwell Time", Value = "5.0", Unit = "sec", SetPoint = "5.0" },
            },
            
            // Material ingredients
            Ingredients = new List<RecipeIngredient>
            {
                new RecipeIngredient { Name = "Base Polymer A", Quantity = 45.5, Unit = "kg", LotNumber = "LOT-A-2024-156" },
                new RecipeIngredient { Name = "Additive B-12", Quantity = 2.3, Unit = "kg", LotNumber = "LOT-B-2024-089" },
                new RecipeIngredient { Name = "Colorant C-Red", Quantity = 0.5, Unit = "kg", LotNumber = "LOT-C-2024-234" },
                new RecipeIngredient { Name = "Stabilizer D", Quantity = 1.2, Unit = "kg", LotNumber = "LOT-D-2024-045" },
            },
            
            Notes = "Standard production run. Quality check required after cycle 50."
        };
    }

    /// <summary>Generates the recipe PDF document with dynamic height.</summary>
    private string GenerateRecipePdfDocument(RecipeData recipe)
    {
        // Label width for KFI NAUT250F: 100mm = 283.46 points (1mm = 2.8346 points)
        double labelWidthMm = 100;
        double labelWidthPt = labelWidthMm * 2.8346;
        
        // Margins
        double marginMm = 3;
        double marginPt = marginMm * 2.8346;
        double contentWidth = labelWidthPt - (2 * marginPt);
        
        // Calculate dynamic height based on content
        double headerHeight = 60;      // Title + recipe info
        double infoSectionHeight = 80; // Batch, operator, date info
        double paramHeaderHeight = 25; // "Process Parameters" title
        double paramRowHeight = 18;    // Each parameter row
        double paramTableHeight = paramHeaderHeight + 20 + (recipe.Parameters.Count * paramRowHeight);
        double ingredHeaderHeight = 25; // "Materials" title
        double ingredRowHeight = 18;   // Each ingredient row
        double ingredTableHeight = ingredHeaderHeight + 20 + (recipe.Ingredients.Count * ingredRowHeight);
        double notesHeight = string.IsNullOrEmpty(recipe.Notes) ? 0 : 50;
        double footerHeight = 30;
        double spacing = 40;           // Total spacing between sections
        
        double totalHeightPt = marginPt + headerHeight + infoSectionHeight + paramTableHeight + 
                               ingredTableHeight + notesHeight + footerHeight + spacing + marginPt;
        
        // Output path - same location as other reports (ProjectFiles/Reports/)
        string projectDir = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        string reportsDir = projectDir.StartsWith("file:///")
            ? projectDir.Substring(8).Replace("/", "\\")
            : projectDir;
        reportsDir = Path.Combine(reportsDir, "Reports");
        Directory.CreateDirectory(reportsDir);
        string outputPath = Path.Combine(reportsDir, $"Recipe_{recipe.RecipeCode}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");

        using (PdfDocument pdfDoc = new PdfDocument())
        {
            pdfDoc.Info.Title = $"Recipe: {recipe.RecipeName}";
            pdfDoc.Info.Author = "FTOptix";
            
            PdfPage page = pdfDoc.AddPage();
            page.Width = XUnit.FromPoint(labelWidthPt);
            page.Height = XUnit.FromPoint(totalHeightPt);

            using (XGraphics gfx = XGraphics.FromPdfPage(page))
            {
                // Colors
                XColor headerBgColor = XColor.FromArgb(0, 51, 102);       // Dark blue
                XColor sectionBgColor = XColor.FromArgb(0, 102, 153);     // Medium blue
                XColor labelBgColor = XColor.FromArgb(230, 240, 250);     // Light blue
                XColor borderColor = XColor.FromArgb(150, 150, 150);
                XColor textColor = XColor.FromArgb(30, 30, 30);
                
                XBrush headerBgBrush = new XSolidBrush(headerBgColor);
                XBrush sectionBgBrush = new XSolidBrush(sectionBgColor);
                XBrush labelBgBrush = new XSolidBrush(labelBgColor);
                XBrush textBrush = new XSolidBrush(textColor);
                XPen borderPen = new XPen(borderColor, 0.5);
                XPen thickBorderPen = new XPen(borderColor, 1);
                
                // Fonts
                XFont titleFont = new XFont("Arial", 14, XFontStyle.Bold);
                XFont subtitleFont = new XFont("Arial", 10, XFontStyle.Bold);
                XFont labelFont = new XFont("Arial", 8, XFontStyle.Bold);
                XFont valueFont = new XFont("Arial", 8);
                XFont smallFont = new XFont("Arial", 7);
                XFont sectionFont = new XFont("Arial", 9, XFontStyle.Bold);
                
                double currentY = marginPt;
                
                // === HEADER SECTION ===
                // Title bar
                gfx.DrawRectangle(headerBgBrush, marginPt, currentY, contentWidth, 28);
                gfx.DrawString(recipe.RecipeName, titleFont, XBrushes.White,
                    new XRect(marginPt, currentY, contentWidth, 28), XStringFormats.Center);
                currentY += 28;
                
                // Recipe code bar
                gfx.DrawRectangle(sectionBgBrush, marginPt, currentY, contentWidth, 20);
                gfx.DrawString($"Code: {recipe.RecipeCode}", subtitleFont, XBrushes.White,
                    new XRect(marginPt, currentY, contentWidth, 20), XStringFormats.Center);
                currentY += 20 + 8;
                
                // === INFO SECTION ===
                double labelColWidth = 70;
                double valueColWidth = contentWidth - labelColWidth;
                double infoRowHeight = 16;
                
                // Batch Number
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, labelColWidth, infoRowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, marginPt + labelColWidth, currentY, valueColWidth, infoRowHeight);
                gfx.DrawString("Batch No:", labelFont, textBrush, marginPt + 3, currentY + 4);
                gfx.DrawString(recipe.BatchNumber, valueFont, textBrush, marginPt + labelColWidth + 3, currentY + 4);
                currentY += infoRowHeight;
                
                // Product Code
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, labelColWidth, infoRowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, marginPt + labelColWidth, currentY, valueColWidth, infoRowHeight);
                gfx.DrawString("Product:", labelFont, textBrush, marginPt + 3, currentY + 4);
                gfx.DrawString(recipe.ProductCode, valueFont, textBrush, marginPt + labelColWidth + 3, currentY + 4);
                currentY += infoRowHeight;
                
                // Machine
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, labelColWidth, infoRowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, marginPt + labelColWidth, currentY, valueColWidth, infoRowHeight);
                gfx.DrawString("Machine:", labelFont, textBrush, marginPt + 3, currentY + 4);
                gfx.DrawString(recipe.MachineName, valueFont, textBrush, marginPt + labelColWidth + 3, currentY + 4);
                currentY += infoRowHeight;
                
                // Operator
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, labelColWidth, infoRowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, marginPt + labelColWidth, currentY, valueColWidth, infoRowHeight);
                gfx.DrawString("Operator:", labelFont, textBrush, marginPt + 3, currentY + 4);
                gfx.DrawString(recipe.OperatorName, valueFont, textBrush, marginPt + labelColWidth + 3, currentY + 4);
                currentY += infoRowHeight;
                
                // Date/Time
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, labelColWidth, infoRowHeight);
                gfx.DrawRectangle(borderPen, XBrushes.White, marginPt + labelColWidth, currentY, valueColWidth, infoRowHeight);
                gfx.DrawString("Date/Time:", labelFont, textBrush, marginPt + 3, currentY + 4);
                gfx.DrawString(recipe.CreatedDate.ToString("yyyy-MM-dd HH:mm"), valueFont, textBrush, marginPt + labelColWidth + 3, currentY + 4);
                currentY += infoRowHeight + 8;
                
                // === PROCESS PARAMETERS SECTION ===
                // Section header
                gfx.DrawRectangle(sectionBgBrush, marginPt, currentY, contentWidth, 18);
                gfx.DrawString("PROCESS PARAMETERS", sectionFont, XBrushes.White,
                    new XRect(marginPt, currentY, contentWidth, 18), XStringFormats.Center);
                currentY += 18;
                
                // Table header
                double paramCol1 = 80;  // Parameter name
                double paramCol2 = 50;  // Value
                double paramCol3 = 35;  // Unit
                double paramCol4 = contentWidth - paramCol1 - paramCol2 - paramCol3; // SetPoint
                
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, contentWidth, 14);
                gfx.DrawLine(borderPen, marginPt, currentY, marginPt + contentWidth, currentY);
                gfx.DrawString("Parameter", smallFont, textBrush, marginPt + 2, currentY + 3);
                gfx.DrawString("Value", smallFont, textBrush, marginPt + paramCol1 + 2, currentY + 3);
                gfx.DrawString("Unit", smallFont, textBrush, marginPt + paramCol1 + paramCol2 + 2, currentY + 3);
                gfx.DrawString("SetPt", smallFont, textBrush, marginPt + paramCol1 + paramCol2 + paramCol3 + 2, currentY + 3);
                currentY += 14;
                
                // Parameter rows
                bool altRow = false;
                foreach (var param in recipe.Parameters)
                {
                    XBrush rowBg = altRow ? new XSolidBrush(XColor.FromArgb(245, 245, 245)) : XBrushes.White;
                    gfx.DrawRectangle(rowBg, marginPt, currentY, contentWidth, paramRowHeight);
                    gfx.DrawLine(borderPen, marginPt, currentY + paramRowHeight, marginPt + contentWidth, currentY + paramRowHeight);
                    
                    gfx.DrawString(param.Name, valueFont, textBrush, marginPt + 2, currentY + 5);
                    gfx.DrawString(param.Value, valueFont, textBrush, marginPt + paramCol1 + 2, currentY + 5);
                    gfx.DrawString(param.Unit, valueFont, textBrush, marginPt + paramCol1 + paramCol2 + 2, currentY + 5);
                    gfx.DrawString(param.SetPoint, valueFont, textBrush, marginPt + paramCol1 + paramCol2 + paramCol3 + 2, currentY + 5);
                    
                    currentY += paramRowHeight;
                    altRow = !altRow;
                }
                
                // Border around parameters table
                gfx.DrawRectangle(thickBorderPen, marginPt, currentY - (recipe.Parameters.Count * paramRowHeight) - 14, contentWidth, (recipe.Parameters.Count * paramRowHeight) + 14);
                currentY += 8;
                
                // === MATERIALS/INGREDIENTS SECTION ===
                // Section header
                gfx.DrawRectangle(sectionBgBrush, marginPt, currentY, contentWidth, 18);
                gfx.DrawString("MATERIALS", sectionFont, XBrushes.White,
                    new XRect(marginPt, currentY, contentWidth, 18), XStringFormats.Center);
                currentY += 18;
                
                // Table header
                double ingredCol1 = 90;  // Material name
                double ingredCol2 = 40;  // Quantity
                double ingredCol3 = 30;  // Unit
                double ingredCol4 = contentWidth - ingredCol1 - ingredCol2 - ingredCol3; // Lot
                
                gfx.DrawRectangle(labelBgBrush, marginPt, currentY, contentWidth, 14);
                gfx.DrawLine(borderPen, marginPt, currentY, marginPt + contentWidth, currentY);
                gfx.DrawString("Material", smallFont, textBrush, marginPt + 2, currentY + 3);
                gfx.DrawString("Qty", smallFont, textBrush, marginPt + ingredCol1 + 2, currentY + 3);
                gfx.DrawString("Unit", smallFont, textBrush, marginPt + ingredCol1 + ingredCol2 + 2, currentY + 3);
                gfx.DrawString("Lot #", smallFont, textBrush, marginPt + ingredCol1 + ingredCol2 + ingredCol3 + 2, currentY + 3);
                currentY += 14;
                
                // Ingredient rows
                altRow = false;
                foreach (var ingred in recipe.Ingredients)
                {
                    XBrush rowBg = altRow ? new XSolidBrush(XColor.FromArgb(245, 245, 245)) : XBrushes.White;
                    gfx.DrawRectangle(rowBg, marginPt, currentY, contentWidth, ingredRowHeight);
                    gfx.DrawLine(borderPen, marginPt, currentY + ingredRowHeight, marginPt + contentWidth, currentY + ingredRowHeight);
                    
                    gfx.DrawString(ingred.Name, valueFont, textBrush, marginPt + 2, currentY + 5);
                    gfx.DrawString(ingred.Quantity.ToString("F2"), valueFont, textBrush, marginPt + ingredCol1 + 2, currentY + 5);
                    gfx.DrawString(ingred.Unit, valueFont, textBrush, marginPt + ingredCol1 + ingredCol2 + 2, currentY + 5);
                    gfx.DrawString(ingred.LotNumber, smallFont, textBrush, marginPt + ingredCol1 + ingredCol2 + ingredCol3 + 2, currentY + 5);
                    
                    currentY += ingredRowHeight;
                    altRow = !altRow;
                }
                
                // Border around ingredients table
                gfx.DrawRectangle(thickBorderPen, marginPt, currentY - (recipe.Ingredients.Count * ingredRowHeight) - 14, contentWidth, (recipe.Ingredients.Count * ingredRowHeight) + 14);
                currentY += 8;
                
                // === NOTES SECTION ===
                if (!string.IsNullOrEmpty(recipe.Notes))
                {
                    gfx.DrawRectangle(labelBgBrush, marginPt, currentY, contentWidth, 14);
                    gfx.DrawString("Notes:", labelFont, textBrush, marginPt + 3, currentY + 3);
                    currentY += 14;
                    
                    gfx.DrawRectangle(borderPen, XBrushes.White, marginPt, currentY, contentWidth, 28);
                    gfx.DrawString(recipe.Notes, smallFont, textBrush,
                        new XRect(marginPt + 3, currentY + 3, contentWidth - 6, 24), XStringFormats.TopLeft);
                    currentY += 28 + 5;
                }
                
                // === FOOTER ===
                gfx.DrawLine(new XPen(headerBgColor, 1), marginPt, currentY, marginPt + contentWidth, currentY);
                currentY += 3;
                string footerText = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                gfx.DrawString(footerText, smallFont, textBrush,
                    new XRect(marginPt, currentY, contentWidth, 12), XStringFormats.Center);
            }
            
            pdfDoc.Save(outputPath);
        }
        
        return outputPath;
    }

    #endregion
}

#region Data Classes

/// <summary>Data model for a single alarm record.</summary>
public class AlarmRecord
{
    public DateTime ActivationTime { get; set; }
    public string AlarmText { get; set; }
    public string AssociatedDevice { get; set; }
}

/// <summary>Data model for a single datalogger record.</summary>
public class DataloggerRecord
{
    public DateTime Timestamp { get; set; }
    public double Value { get; set; }
}

/// <summary>Data model for a machine recipe.</summary>
public class RecipeData
{
    public string RecipeName { get; set; }
    public string RecipeCode { get; set; }
    public string BatchNumber { get; set; }
    public string ProductCode { get; set; }
    public DateTime CreatedDate { get; set; }
    public string OperatorName { get; set; }
    public string MachineName { get; set; }
    public List<RecipeParameter> Parameters { get; set; } = new List<RecipeParameter>();
    public List<RecipeIngredient> Ingredients { get; set; } = new List<RecipeIngredient>();
    public string Notes { get; set; }
}

/// <summary>A single process parameter in a recipe.</summary>
public class RecipeParameter
{
    public string Name { get; set; }
    public string Value { get; set; }
    public string Unit { get; set; }
    public string SetPoint { get; set; }
}

/// <summary>A single ingredient/material in a recipe.</summary>
public class RecipeIngredient
{
    public string Name { get; set; }
    public double Quantity { get; set; }
    public string Unit { get; set; }
    public string LotNumber { get; set; }
}

#endregion
