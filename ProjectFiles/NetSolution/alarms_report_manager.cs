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
// Aliases to resolve namespace conflicts
using MigraDocColor = MigraDoc.DocumentObjectModel.Color;
using MigraDocColors = MigraDoc.DocumentObjectModel.Colors;
using MigraDocTable = MigraDoc.DocumentObjectModel.Tables.Table;
using MigraDocColumn = MigraDoc.DocumentObjectModel.Tables.Column;
using MigraDocRow = MigraDoc.DocumentObjectModel.Tables.Row;
using MigraDocCell = MigraDoc.DocumentObjectModel.Tables.Cell;
using MigraDocVerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment;
#endregion

/// <summary>
/// Generates PDF alarm reports from EmbeddedDatabase1.AlarmsEventLogger1.
/// Uses MigraDoc/PDFsharp for PDF generation with locale-aware message columns.
/// </summary>
public class alarms_report_manager : BaseNetLogic
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
        try
        {
            GenerateReport(startDate, endDate);
        }
        catch (Exception ex)
        {
            Log.Error("AlarmReport.Generate", $"Error generating report: {ex.Message}");
        }
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

    #endregion

    #region Private Methods

    /// <summary>Orchestrates report generation: fetch data, create document, save PDF.</summary>
    private void GenerateReport(DateTime startDate, DateTime endDate)
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
}

#region Data Classes

/// <summary>Data model for a single alarm record.</summary>
public class AlarmRecord
{
    public DateTime ActivationTime { get; set; }
    public string AlarmText { get; set; }
    public string AssociatedDevice { get; set; }
}

#endregion
