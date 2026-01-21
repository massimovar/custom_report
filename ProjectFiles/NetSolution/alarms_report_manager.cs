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
/// Alarm Report Generator for FTOptix
/// Generates PDF reports from alarm history data stored in EmbeddedDatabase1.AlarmsEventLogger1
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
        // Enable System.Drawing support for PDFsharp/MigraDoc on .NET Core/.NET 5+
        AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);
        
        // Register CodePages encoding provider for Windows-1252 support (required by PDFsharp)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        // Initialize the store reference
        _store = Project.Current.Get<Store>("DataStores/EmbeddedDatabase1");
        
        // Set default values - these can be overridden via NetLogic variables
        _machineName = GetVariableValueOrDefault("MachineName", "Machine Name");
        _customerLogoPath = GetVariableValueOrDefault("CustomerLogoPath", "%PROJECTDIR&%/imgs/logo1.svg");
        _lastLogoPath = GetVariableValueOrDefault("LastLogoPath", "%PROJECTDIR&%/imgs/logo2.svg");
        
        Log.Info("AlarmReport.Init", "Alarm Report Manager initialized");
    }

    public override void Stop()
    {
        // Cleanup if needed
    }

    #region Public Methods (Exported)

    /// <summary>
    /// Generates the alarm report PDF with DateTime parameters directly
    /// </summary>
    /// <param name="startDate">Start date for the filter</param>
    /// <param name="endDate">End date for the filter</param>
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

    /// <summary>
    /// Sets the customer logo path for the report header
    /// </summary>
    /// <param name="logoPath">Path to the customer logo image</param>
    [ExportMethod]
    public void SetCustomerLogo(string logoPath)
    {
        _customerLogoPath = logoPath;
        Log.Info("AlarmReport.Config", $"Customer logo path set to: {logoPath}");
    }

    /// <summary>
    /// Sets the Last logo path for the report header
    /// </summary>
    /// <param name="logoPath">Path to the Last logo image</param>
    [ExportMethod]
    public void SetLastLogo(string logoPath)
    {
        _lastLogoPath = logoPath;
        Log.Info("AlarmReport.Config", $"Last logo path set to: {logoPath}");
    }

    /// <summary>
    /// Sets the machine name for the report header
    /// </summary>
    /// <param name="machineName">Machine name to display in the report</param>
    [ExportMethod]
    public void SetMachineName(string machineName)
    {
        _machineName = machineName;
        Log.Info("AlarmReport.Config", $"Machine name set to: {machineName}");
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Main report generation logic
    /// </summary>
    private void GenerateReport(DateTime startDate, DateTime endDate)
    {
        Log.Info("AlarmReport.Generate", $"Generating alarm report from {startDate} to {endDate}");
        
        // Fetch alarm data from database
        var alarmData = FetchAlarmData(startDate, endDate);
        
        if (alarmData == null || alarmData.Count == 0)
        {
            Log.Warning("AlarmReport.Generate", "No alarm data found for the specified date range");
            // Still generate report with empty table
        }
        
        // Create the PDF document
        Document document = CreateDocument(startDate, endDate, alarmData);
        
        // Render and save the PDF
        string outputPath = GetOutputPath(startDate, endDate);
        SaveDocument(document, outputPath);
        
        Log.Info("AlarmReport.Generate", $"Report generated successfully: {outputPath}");
    }

    /// <summary>
    /// Fetches alarm data from the EmbeddedDatabase1.AlarmsEventLogger1 table
    /// </summary>
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
            
            // First, let's query all data to debug - check what's in the table
            Object[,] debugResultSet;
            String[] debugHeader;
            string debugQuery = "SELECT * FROM AlarmsEventLogger1 LIMIT 5";
            
            try
            {
                _store.Query(debugQuery, out debugHeader, out debugResultSet);
                if (debugHeader != null && debugHeader.Length > 0)
                {
                    Log.Info("AlarmReport.Query", $"Available columns in AlarmsEventLogger1: {string.Join(", ", debugHeader)}");
                    if (debugResultSet != null && debugResultSet.GetLength(0) > 0)
                    {
                        Log.Info("AlarmReport.Query", $"Table has {debugResultSet.GetLength(0)} sample rows");
                        
                        // Log the first row's LocalTime value to see the actual format
                        int ltIdx = Array.IndexOf(debugHeader, "LocalTime");
                        if (ltIdx >= 0 && debugResultSet[0, ltIdx] != null)
                        {
                            var sampleLocalTime = debugResultSet[0, ltIdx];
                            Log.Info("AlarmReport.Query", $"Sample LocalTime value: '{sampleLocalTime}' (Type: {sampleLocalTime.GetType().Name})");
                        }
                    }
                    else
                    {
                        Log.Warning("AlarmReport.Query", "Table AlarmsEventLogger1 is EMPTY - no rows found");
                    }
                }
            }
            catch (Exception debugEx)
            {
                Log.Warning("AlarmReport.Query", $"Debug query failed: {debugEx.Message}");
            }
            
            // Log filter parameters
            Log.Info("AlarmReport.Query", $"Filter Start: {startDate} | Filter End: {endDate}");
            
            // Query WITHOUT date filter first to get total count
            Object[,] countResultSet;
            String[] countHeader;
            string countQuery = "SELECT COUNT(*) as TotalCount FROM AlarmsEventLogger1";
            _store.Query(countQuery, out countHeader, out countResultSet);
            if (countResultSet != null && countResultSet.GetLength(0) > 0)
            {
                Log.Info("AlarmReport.Query", $"Total rows in table (no filter): {countResultSet[0, 0]}");
            }
            
            // Build the query using LocalTime with proper date format
            // Database uses ISO 8601 format with T separator: 2026-01-21T14:32:10.1762615
            string startDateStr = startDate.ToString("yyyy-MM-ddTHH:mm:ss");
            string endDateStr = endDate.ToString("yyyy-MM-ddTHH:mm:ss");
            
            Log.Info("AlarmReport.Query", $"Querying alarms WHERE LocalTime >= '{startDateStr}' AND LocalTime <= '{endDateStr}'");
            
            // Build the query using LocalTime column
            string query = $@"SELECT * FROM AlarmsEventLogger1 
                              WHERE LocalTime >= ""{startDateStr}"" AND LocalTime <= ""{endDateStr}""
                              ORDER BY LocalTime DESC";
            
            Log.Info("AlarmReport.Query", $"Executing query: {query}");
            
            // Execute the query
            Object[,] resultSet;
            String[] header;
            _store.Query(query, out header, out resultSet);
            
            if (resultSet == null || resultSet.GetLength(0) == 0)
            {
                Log.Info("AlarmReport.Query", "Query returned no results for the date range");
                return alarmRecords;
            }
            
            // Log available columns for debugging
            Log.Info("AlarmReport.Query", $"Query returned columns: {string.Join(", ", header)}");
            Log.Info("AlarmReport.Query", $"Query returned {resultSet.GetLength(0)} rows");
            
            // Get the current session's locale for localized message column
            string currentLocale = GetCurrentLocale();
            string localizedMessageColumn = $"Message_{currentLocale}";
            Log.Info("AlarmReport.Query", $"Current locale: {currentLocale}, looking for column: {localizedMessageColumn}");
            
            // Find column indices - try localized message column first, then fallback to generic Message
            int localTimeIdx = FindColumnIndex(header, "LocalTime", "Time", "Timestamp", "EventTime");
            int messageIdx = FindColumnIndex(header, localizedMessageColumn);
            
            // If localized column not found, try fallback options
            if (messageIdx < 0)
            {
                Log.Warning("AlarmReport.Query", $"Localized column '{localizedMessageColumn}' not found, trying fallback columns");
                messageIdx = FindColumnIndex(header, "Message", "Text", "Description", "AlarmMessage");
            }
            
            int sourceNameIdx = FindColumnIndex(header, "SourceName", "Source", "DeviceName", "AlarmName");
            int sourcePathIdx = FindColumnIndex(header, "SourcePath", "Path", "ObjectPath");
            
            Log.Info("AlarmReport.Query", $"Column indices - Time:{localTimeIdx}, Message:{messageIdx} (column: {(messageIdx >= 0 ? header[messageIdx] : "N/A")}), SourceName:{sourceNameIdx}, SourcePath:{sourcePathIdx}");
            
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
                    AssociatedDevice = sourceNameIdx >= 0 && resultSet[i, sourceNameIdx] != null 
                        ? resultSet[i, sourceNameIdx].ToString() 
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

    /// <summary>
    /// Creates the MigraDoc document with header and body
    /// </summary>
    private Document CreateDocument(DateTime startDate, DateTime endDate, List<AlarmRecord> alarmData)
    {
        Document document = new Document();
        document.Info.Title = "Alarm Report";
        document.Info.Author = "FTOptix";
        document.Info.Subject = $"Alarm Report from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}";
        
        // Define styles
        DefineStyles(document);
        
        // Create the main section
        Section section = document.AddSection();
        section.PageSetup.TopMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.LeftMargin = Unit.FromCentimeter(1.5);
        section.PageSetup.RightMargin = Unit.FromCentimeter(1.5);
        
        // Add Header
        AddReportHeader(section, startDate, endDate);
        
        // Add Body (Alarm Table)
        AddReportBody(section, alarmData);
        
        // Add Footer with page numbers
        AddReportFooter(section);
        
        return document;
    }

    /// <summary>
    /// Defines document styles
    /// </summary>
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

    /// <summary>
    /// Adds the report header section with logos, title, and filter dates
    /// </summary>
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
        if (!string.IsNullOrEmpty(_customerLogoPath) && File.Exists(GetAbsolutePath(_customerLogoPath)))
        {
            try
            {
                var image = logoLeftCell.AddImage(GetAbsolutePath(_customerLogoPath));
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("AlarmReport.Header", $"Could not load customer logo: {ex.Message}");
            }
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
        if (!string.IsNullOrEmpty(_lastLogoPath) && File.Exists(GetAbsolutePath(_lastLogoPath)))
        {
            try
            {
                var image = logoRightCell.AddImage(GetAbsolutePath(_lastLogoPath));
                image.Width = Unit.FromCentimeter(3);
                image.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                Log.Warning("AlarmReport.Header", $"Could not load Last logo: {ex.Message}");
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

    /// <summary>
    /// Adds the report body with the alarm data table
    /// </summary>
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

    /// <summary>
    /// Adds footer with page numbers and generation timestamp
    /// </summary>
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

    /// <summary>
    /// Renders and saves the document to PDF
    /// </summary>
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

    /// <summary>
    /// Generates the output file path for the report
    /// </summary>
    private string GetOutputPath(DateTime startDate, DateTime endDate)
    {
        // Try to get custom output path from NetLogic variable
        string customPath = GetVariableValueOrDefault("OutputPath", "");
        
        if (!string.IsNullOrEmpty(customPath))
        {
            return customPath;
        }
        
        // Default: save to ProjectFiles/Reports folder
        string projectDir = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        string reportsDir = Path.Combine(projectDir, "Reports");
        string fileName = $"AlarmReport_{startDate:yyyyMMdd}_{endDate:yyyyMMdd}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
        
        return Path.Combine(reportsDir, fileName);
    }

    /// <summary>
    /// Gets the absolute path from a project-relative or absolute path
    /// </summary>
    private string GetAbsolutePath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return "";
            
        // If it's already an absolute path, return it
        if (Path.IsPathRooted(path))
            return path;
        
        // Otherwise, treat as project-relative
        string projectDir = new ResourceUri(ResourceUri.FromProjectRelativePath("")).Uri;
        return Path.Combine(projectDir, path);
    }

    /// <summary>
    /// Gets a variable value from the NetLogic or returns the default value
    /// </summary>
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

    /// <summary>
    /// Finds a column index by trying multiple possible column names (case-insensitive)
    /// </summary>
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

    /// <summary>
    /// Gets the current session's ActualLocale (e.g., "it-IT", "en-US")
    /// Used to determine which localized Message column to read from the database
    /// </summary>
    private string GetCurrentLocale()
    {
        try
        {
            // Get the ActualLocaleId from the current session
            string localeId = Session.ActualLocaleId;
            if (!string.IsNullOrEmpty(localeId))
            {
                Log.Info("AlarmReport.Locale", $"Session ActualLocaleId: {localeId}");
                return localeId;
            }
        }
        catch (Exception ex)
        {
            Log.Warning("AlarmReport.Locale", $"Error getting session locale: {ex.Message}");
        }
        
        // Default fallback
        Log.Warning("AlarmReport.Locale", "Could not determine locale, using default 'en-US'");
        return "en-US";
    }

    #endregion
}

#region Data Classes

/// <summary>
/// Represents a single alarm record from the database
/// </summary>
public class AlarmRecord
{
    public DateTime ActivationTime { get; set; }
    public string AlarmText { get; set; }
    public string AssociatedDevice { get; set; }
}

#endregion
