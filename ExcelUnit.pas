{$reference 'Microsoft.Office.Interop.Excel.dll'}
{$reference 'System.Drawing.dll'}

unit ExcelUnit;

interface
uses Microsoft.Office.Interop.Excel, system, system.Runtime, system.IO, system.Drawing, system.Runtime.InteropServices;

type
  BorderRoute = Microsoft.Office.Interop.Excel.XlBordersIndex;
  
  FormatType =
  (
    General,
    Text,
    Number,
    Percent,
    Money,
    Date
  );
  
  Design = class
    private
      rang: Microsoft.Office.Interop.Excel.Range;

      procedure SetCellWidth(value: real);
      procedure SetCellHeight(value: real);
      procedure SetBorder(b: BorderRoute; value: boolean);
      procedure SetCellColor(value: color);
      procedure SetBorderColor(b: BorderRoute; value: color);
      procedure SetTextSize(value: real);
      procedure SetTextBold(value: boolean);
      procedure SetTextColor(value: color);
      
      function GetCellWidth:= convert.ToDouble(rang.ColumnWidth);
      function GetCellHeight:= convert.ToDouble(rang.RowHeight);
      function GetBorder(b: BorderRoute): boolean;
      function GetCellColor:= colortranslator.FromOle(convert.ToInt32(rang.Interior.Color));
      function GetBorderColor(b: BorderRoute):= colortranslator.FromOle(convert.ToInt32(rang.Borders[b].Color));
      function GetTextSize:= convert.ToDouble(rang.Font.Size);
      function GetTextBold:= convert.ToBoolean(rang.Font.Bold);
      function GetTextColor:= colortranslator.FromOle(convert.ToInt32(rang.Font.Color));
    public
      ///Задаёт или возвращает ширину столбцов(0-255)
      property Width: real read GetCellWidth write SetCellWidth;
      ///Задаёт или возвращает высоту строк(0-409)
      property Height: real read GetCellHeight write SetCellHeight;
      ///Задаёт или возвращает наличие границы
      property Border[b: BorderRoute]: boolean read GetBorder write SetBorder;
      ///Задаёт или возвращает цвет заливки
      property CellColor: Color read GetCellColor write SetCellColor;
      ///Задаёт или возвращает цвет границы
      property BorderColor[b: BorderRoute]: Color read GetBorderColor write SetBorderColor;
      ///Задаёт или возвращает размер шрифта
      property TextSize: real read GetTextSize write SetTextSize;
      ///Задаёт или возвращает налилчие жирности текста
      property TextBold: boolean read GetTextBold write SetTextBold;
      ///Задаёт или возвращает цвет текста
      property TextColor: Color read GetTextColor write SetTextColor;

      ///Централизует содержимое ячейки
      procedure Centralize;
      ///Автовыводит ширину и высоту
      procedure AutoSize;
      ///Задаёт формат ячеек
      procedure CellFormat(format: FormatType);
      ///<summary>Задаёт формат ячеек</summary>
      ///<param name="number">Кол-во значящих цифр после запятой(для числового и процентного типов)</param>
      procedure CellFormat(format: FormatType; number: integer);
  end;
  
  Cell = class
    private 
      cel: Microsoft.Office.Interop.Excel.Range;
      
      procedure SetCellVal(value: object);
      
      function GetCelVal:= cel.Value2;
      function GetCellDesign: Design;
    public
      ///Возвращает или задаёт значение ячейки
      property Val: object read GetCelVal write SetCellVal;
      ///Предоставляет методы и свойства для оформления ячейки
      property CellDesign: Design read GetCellDesign;
            
      ///Очищает ячейку
      procedure Clear:= cel.Clear;
  end;

  Range = class
    private
      rang: Microsoft.Office.Interop.Excel.Range;
      
      procedure SetRangeVal(value: array[,] of object);
      procedure SetRangeMerge(value: boolean);
      
      function GetRangeVal: array[,] of object;
      function GetRangeMerge:= convert.ToBoolean(rang.MergeCells);
      function GetRangeDesign: Design;
    public
      ///Возвращает массив значений диапазона ячеек
      property Val: array [,] of object read GetRangeVal write SetRangeVal;
      ///Возвращает или задёт группировку диапозона ячеек
      property Merge: boolean read GetRangeMerge write SetRangeMerge;
      ///Предоставляет методы и свойства для оформления диапазона ячеек
      property RangeDesign: Design read GetRangeDesign;
      
      ///Очищает диапазон ячеек
      procedure Clear:= rang.Clear;
  end;

  ExcelApp = class
    private
      app: Microsoft.Office.Interop.Excel.ApplicationClass;
      ws: Microsoft.Office.Interop.Excel.Worksheet;
      
      procedure SetBook(value: string);
      procedure SetSheet(value: integer);
      procedure SetSheetName(value: string);
      
      function GetBook:= app.Workbooks[1].FullName;
      function GetSheet:= ws.Index;
      function GetSheetName:= ws.Name;
      function GetCell(i, j: integer): Cell;
      function GetRange(i, j, ii, jj: integer): Range;
    public
      ///<summary>Создаёт новый экземпляр класса</summary>
      constructor Create;
      ///<summary>Создаёт новый экземпляр класса и открывает Excel с указаной книгой</summary>
      ///<param name="path">Путь к книге</param>
      constructor Create(path: string);
      ///<summary>Создаёт новый экземпляр класса и открывает Excel с указанными книгой и листом</summary>
      ///<param name="path">Путь к книге</param>
      ///<param name="sheet">Номер листа(начиная с 1)</param>
      constructor Create(path: string; sheet: integer);
      
      ///Возвращает или задаёт адрес книги
      property Book: string read GetBook write SetBook;
      ///Возвращает или задёт номер активного листа в книге(начиная с 1)
      property Sheet: integer read GetSheet write SetSheet;
      ///Возвращает или задёт имя активного листа в книге
      property SheetName: string read GetSheetName write SetSheetName;
      ///Возвращает ячейку с адресом [y, x] (начиная с [1, 1])
      property CellOne[i, j: integer]: Cell read GetCell; default;
      ///Возвращает диапазон ячеек начиная с [y, x] по [y2, x2]
      property CellRange[i, j, ii, jj: integer]: Range read GetRange;
      
      ///<summary>Открывает Excel с указанной книгой</summary>
      ///<param name="path">Путь к книге</param>
      procedure Open(path: string);
      ///<summary>Открывает Excel с указанными книгой и листом</summary>
      ///<param name="path">Путь к книге</param>
      ///<param name="sheet">Номер листа(начиная с 1)</param>
      procedure Open(path: string; sheet: integer);
      ///Удаляет активный лист из книги
      procedure SheetDel:= ws.Delete;
      ///<summary>Удаляет столбец</summary>
      ///<param name="number">Номер столбца(начиная с 1)</param>
      procedure ColumnDel(number: integer):= (ws.Columns[number, system.Type.Missing] as  Microsoft.Office.Interop.Excel.Range).delete;
      ///<summary>Удаляет строку</summary>
      ///<param name="number">Номер строки(начиная с 1)</param>
      procedure RowDel(number: integer):= (ws.Rows[number, system.Type.Missing] as  Microsoft.Office.Interop.Excel.Range).delete;
      ///Сохраняет изменения и закрывает книгу
      procedure Save:= app.Workbooks[1].Close(true);
      ///Закрывает Excel без сохраениения изменений
      procedure Close;
  end;
    
implementation
{$region ExelApp}
  constructor ExcelApp.Create;
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
  end;
  
  constructor ExcelApp.Create(path: string);
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
    open(path);
  end;
  
  constructor Excelapp.Create(path: string; sheet: integer);
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
    open(path);
    SetSheet(sheet);
  end;
  
  procedure Excelapp.Open(path: string);
  begin
    if not system.IO.File.Exists(path) then
      begin
        var x:= new Microsoft.Office.Interop.Excel.ApplicationClass;
        x.Workbooks.Add;
        x.Workbooks[1].SaveAs(path);
        x.Quit;
      end;
    
    if app.Workbooks.Count > 0 then
      app.Workbooks[1].Close(false);
    
    app.Workbooks.Open(path);
    app.DisplayAlerts:= false;
    ws:= app.Workbooks[1].Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  procedure ExcelApp.Open(path: string; sheet: integer);
  begin
    Open(path);
    SetSheet(sheet);
  end;
  
  procedure ExcelApp.Close;
  begin
    if app.Workbooks.Count > 0 then
      app.Workbooks[1].Close(false);
    app.Quit;
  end;

  procedure ExcelApp.SetBook(value: string);
  begin
    if app.Workbooks.Count > 0 then
      app.Workbooks[1].Close(false);
    app.Workbooks.Open(value);
    ws:= app.Workbooks[1].Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  procedure ExcelApp.SetSheet(value: integer);
  begin
    if app.Workbooks[1].Worksheets.Count < value then
      for var i:= app.Workbooks[1].Worksheets.Count to value-1 do
        app.Workbooks[1].Worksheets.Add(system.Type.Missing, app.Workbooks[1].Worksheets[i]);
    ws:= app.workbooks[1].Worksheets[value] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  procedure ExcelApp.SetSheetName(value: string);
  begin
    ws.Name:= value;
  end;
  
  function ExcelApp.GetCell(i, j: integer): Cell;
  begin
    result:= new Cell;
    result.cel:= ws.get_range(ws.Cells[i, j], ws.Cells[i, j]);
  end;
  
  function ExcelApp.GetRange(i, j, ii, jj: integer): Range;
  begin
    result:= new Range;
    result.rang:= ws.get_range(ws.Cells[i, j], ws.Cells[ii, jj]);
  end;
{$endregion}

{$region Range}
  procedure Range.SetRangeVal(value: array[,] of object);
  begin
    rang.Value2:= value;
  end;
  
  procedure Range.SetRangeMerge(value: boolean);
  begin
    if value then
      rang.Merge
    else
      rang.UnMerge;
  end;
  
  function Range.GetRangeVal: array[,] of object;
  begin
    result:= new object[rang.Columns.Count, rang.Rows.Count];
    result:= rang.Value2 as array[,] of object;
  end;
  
  function Range.GetRangeDesign: Design;
  begin
    result:= new Design;
    result.rang:= rang;
  end;
{$endregion}

{$region Cell}
  procedure Cell.SetCellVal(value: object);
  begin
    cel.Value2:= value;
  end;
  
  function Cell.GetCellDesign: Design;
  begin
    result:= new Design;
    result.rang:= cel;
  end;
{$endregion}  

{$region Design}
  procedure Design.Centralize;
  begin
    rang.HorizontalAlignment:= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
    rang.VerticalAlignment:= Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
  end;
 
  procedure Design.AutoSize;
  begin
    rang.Columns.AutoFit;
    rang.Rows.AutoFit;
  end;
  
  procedure Design.CellFormat(format: FormatType);
  begin
    if format = formattype.General then
      rang.NumberFormat:= 'general'
    else if format = formattype.Text then
      rang.NumberFormat:= '@'#0
    else if format = formattype.Date then
      rang.NumberFormat:= 'm/d/yyyy'
    else if format = formattype.Money then
      rang.NumberFormat:= '#,##0 $'
    else if format = formattype.Number then
      rang.NumberFormat:= '0'#0
    else if format = formattype.Percent then
      rang.NumberFormat:= '0%';
  end;  
    
  procedure Design.CellFormat(format: FormatType; number: integer);
  begin
    if (format = formattype.Date) or (format = formattype.General) or (format = formattype.Money) or (format = formattype.Text) then
      exit;
    
    if number < 1 then
      begin
        Cellformat(format);
        exit;
      end;
      
    var s:= '0.'+('0'*number);
    
    if format = formattype.Number then
      rang.NumberFormat:= s
    else
      rang.NumberFormat:= s+'%';
  end;

  procedure Design.SetCellWidth(value: real);
  begin
    rang.ColumnWidth:= value;
  end;
  
  procedure Design.SetCellHeight(value: real);
  begin
    rang.RowHeight:= value;
  end;
  
  procedure Design.SetBorder(b: BorderRoute; value: boolean);
  begin
    if value then
      rang.Borders[b].LineStyle:= Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    else 
      rang.Borders[b].LineStyle:= Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
  end;
  
  procedure Design.SetCellColor(value: color);
  begin
    if (value.R = 255) and (value.G = 255) and (value.B = 255) then
      rang.Interior.ColorIndex:=0
    else
      rang.Interior.Color:= colortranslator.ToOle(value);
  end;
  
  procedure Design.SetBorderColor(b: BorderRoute; value: color);
  begin
    rang.Borders[b].Color:= colortranslator.ToOle(value);
  end;
  
  procedure Design.SetTextSize(value: real);
  begin
    rang.Font.Size:= value;
  end;
  
  procedure Design.SetTextBold(value: boolean);
  begin
    rang.Font.Bold:= value;
  end;
  
  procedure Design.SetTextColor(value: color);
  begin
    rang.Font.Color:= colortranslator.ToOle(value);
  end;

  function Design.GetBorder(b: BorderRoute): boolean;
  begin
    result:=
      if rang.Borders[b].LineStyle.ToString='-4142' then
        false
      else
        true;
  end;
{$endregion}
end.