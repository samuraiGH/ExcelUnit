{$reference 'Microsoft.Office.Interop.Excel.dll'}
{$reference 'System.Drawing.dll'}

unit ExcelUnitPlus;

interface
uses Microsoft.Office.Interop.Excel, system, system.Drawing;

type
  BorderRoute=Microsoft.Office.Interop.Excel.XlBordersIndex;
  
  FormatType=
  (
    General,
    Text,
    Number,
    Percent,
    Money,
    Date
  );
  
  Design=class
    private
      rang: Microsoft.Office.Interop.Excel.Range;
      
      procedure SetTextColor(value: color);
      procedure SetTextBold(value: boolean);
      procedure SetTextSize(value: real);
      procedure SetCellColor(value: color);
      procedure SetCellWidth(value: real);
      procedure SetCellHeight(value: real);
      procedure SetBorder(b: BorderRoute; value: boolean);
      procedure SetBorderColor(b: BorderRoute; value: color);
      
      function GetCellColor:= colortranslator.FromOle(convert.ToInt32(rang.Interior.Color));
      function GetTextColor:= colortranslator.FromOle(convert.ToInt32(rang.Font.Color));
      function GetBorderColor(b: BorderRoute):= colortranslator.FromOle(convert.ToInt32(rang.Borders[b].Color));
      function GetTextBold:= convert.ToBoolean(rang.Font.Bold);
      function GetCellWidth:= convert.ToDouble(rang.ColumnWidth);
      function GetCellHeight:= convert.ToDouble(rang.RowHeight);
      function GetTextSize:= convert.ToDouble(rang.Font.Size);
      function GetBorder(b: BorderRoute): boolean;
    public
      ///Задаёт или возвращает цвет заливки
      property CellColor: Color read GetCellColor write SetCellColor;
      ///Задаёт или возвращает цвет текста
      property TextColor: Color read GetTextColor write SetTextColor;
      ///Задаёт или возвращает налилчие жирности текста
      property TextBold: boolean read GetTextBold write SetTextBold;
      ///Задаёт или возвращает размер шрифта
      property TextSize: real read GetTextSize write SetTextSize;
      ///Задаёт или возвращает ширину столбцов(0-255)
      property Width: real read GetCellWidth write SetCellWidth;
      ///Задаёт или возвращает высоту строк(0-409)
      property Height: real read GetCellHeight write SetCellHeight;
      ///Задаёт или возвращает наличие границы
      property Border[b: BorderRoute]: boolean read GetBorder write SetBorder;
      ///Задаёт или возвращает цвет границы
      property BorderColor[b: BorderRoute]: Color read GetBorderColor write SetBorderColor;
      
      ///Централизует содержимое ячейки
      procedure Centralize;
      ///Автовыводит ширину и высоту
      procedure AutoSize;
      ///Задаёт формат ячеек
      procedure CellFormat(format: FormatType);
      ///<summary>Задаёт формат ячеек</summary>
      ///<param name="number">количество значящих цифр после запятой для числового и процентного типов</param>
      procedure CellFormat(format: FormatType; number: integer);
  end;
  
  Cell=class
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
      procedure Clear:=cel.Clear;
  end;

  Range=class
    private
      rang: Microsoft.Office.Interop.Excel.Range;
      
      procedure SetRangeVal(value: array[,] of object);
      procedure SetRangeMerge(value: boolean);
      
      function GetRangeVal: array[,] of object;
      function GetRangeMerge:= convert.ToBoolean(rang.MergeCells);
      function GetRangeDesign: Design;
    public
      ///Возвращает или задёт группировку диапозона ячеек
      property Merge: boolean read GetRangeMerge write SetRangeMerge;
      ///Возвращает массив значений диапазона ячеек
      property Val: array [,] of object read GetRangeVal write SetRangeVal;
      ///Предоставляет методы и свойства для оформления диапазона ячеек
      property RangeDesign: Design read GetRangeDesign;
      
      ///Очищает диапазон ячеек
      procedure Clear:=rang.Clear;
  end;

  ExcelApp=class
    private
      app: Microsoft.Office.Interop.Excel.ApplicationClass;
      ws: Microsoft.Office.Interop.Excel.Worksheet;
      
      procedure SetBook(value: string);
      procedure SetSheet(index: integer);
      
      function GetBook:= app.Workbooks[1].FullName;
      function GetSheet:= ws.Index;
      function GetCell(i, j: integer): Cell;
      function GetRange(i, j, ii, jj: integer): Range;
    public
      ///<summary>Создаёт новый экземпляр класса</summary>
      constructor Create;
      ///<summary>Создаёт новый экземпляр класса и открывает Excel с указаной книгой</summary>
      ///<param name="path">Путь к книге</param>
      constructor Create(path: string);
      ///<summary>Создаёт новый экземпляр класса и открывает Excel с указаной книгой</summary>
      ///<param name="path">Путь к книге</param>
      ///<param name="sheet">Номер листа</param>
      constructor Create(path: string; sheet: integer);
      
      ///Возвращает или задаёт адрес книги
      property Book: string read GetBook write SetBook;
      ///Возвращает или задёт номер листа в книге, начиная с 1
      property Sheet: integer read GetSheet write SetSheet;
      ///Возвращает ячейку с адресом [y, x] начиная с 1
      property CellOne[i, j: integer]: Cell read GetCell;
      ///Возвращает диапазон ячеек начиная с [y, x] по [y2, x2]
      property CellRange[i, j, ii, jj: integer]: Range read GetRange;
      
      ///<summary>Открывает Excel с указаной книгой</summary>
      ///<param name="path">Путь к книге</param>
      procedure Open(path: string);
      ///Сохраняет изменения и закрывает книгу
      procedure Save:=app.Workbooks[1].Close(true);
      ///Закрывает Excel без сохраениения изменений
      procedure Close;
  end;
    
implementation
//реализация приложения
  constructor excelapp.Create;
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
  end;
  
  constructor ExcelApp.Create(path: string);
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
    open(path);
  end;
  
  constructor excelapp.Create(path: string; sheet: integer);
  begin
    app:= new Microsoft.Office.Interop.Excel.ApplicationClass;
    open(path);
    SetSheet(sheet);
  end;
  
  procedure excelapp.open(path: string);
  begin
    if system.IO.File.Exists(path)=false then
      begin
        var x:=Microsoft.Office.Interop.Excel.ApplicationClass.Create.Workbooks.Add;
        x.SaveAs(path);
      end;
    
    if app.Workbooks.Count>0 then
      app.Workbooks[1].Close(false);
    
    app.Workbooks.Open(path);
    app.DisplayAlerts:=false;
    ws:=app.Workbooks[1].Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  procedure ExcelApp.Close;
  begin
    if app.Workbooks.Count>0 then
      app.Workbooks[1].Close(false);
    app.Quit;
  end;

  procedure ExcelApp.SetBook(value: string);
  begin
    if app.Workbooks.Count>0 then
      app.Workbooks[1].Close(false);
    app.Workbooks.Open(value);
    ws:=app.Workbooks[1].Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  procedure ExcelApp.SetSheet(index: integer);
  begin
    if app.Workbooks[1].Worksheets.Count<index-1 then
      for var i:=app.Workbooks[1].Worksheets.Count to index do
        app.Workbooks[1].Worksheets.Add(system.Reflection.Missing.Value, app.Workbooks[1].Worksheets[i]);
    ws:=app.workbooks[1].Worksheets[index] as Microsoft.Office.Interop.Excel.Worksheet;
  end;
  
  function ExcelApp.GetCell(i, j: integer): Cell;
  begin
    result:= new Cell;
    result.cel:=ws.get_range(ws.Cells[i, j], ws.Cells[i, j]);
  end;
  
  function ExcelApp.GetRange(i, j, ii, jj: integer): Range;
  begin
    result:= new Range;
    result.rang:=ws.get_range(ws.Cells[i, j], ws.Cells[ii, jj]);
  end;
  
  //реализация диапозона
  procedure Range.SetRangeMerge(value: boolean);
  begin
    if value then
      rang.Merge
    else
      rang.UnMerge;
  end;
  
  procedure Range.SetRangeVal(value: array[,] of object);
  begin
    rang.Value2:=value;
  end;
  
  function Range.GetRangeVal: array[,] of object;
  begin
    result:=new object[rang.Columns.Count, rang.Rows.Count];
    result:=rang.Value2 as array[,] of object;
  end;
  
  function Range.GetRangeDesign: Design;
  begin
    result:= new Design;
    result.rang:=rang;
  end;
  
  //реализация ячейки
  procedure Cell.SetCellVal(value: object);
  begin
    cel.Value2:=value;
  end;
  
  function Cell.GetCellDesign: Design;
  begin
    result:= new Design;
    result.rang:=cel;
  end;
  
  //реализация дизайна
  procedure Design.SetCellColor(value: color);
  begin
    if (value.R=255) and (value.G=255) and (value.B=255) then
      rang.Interior.ColorIndex:=0
    else
      rang.Interior.Color:=colortranslator.ToOle(value);
  end;
  
  procedure Design.SetTextColor(value: color);
  begin
    rang.Font.Color:=colortranslator.ToOle(value);
  end;
  
  procedure Design.SetCellWidth(value: real);
  begin
    rang.ColumnWidth:=value;
  end;
  
  procedure Design.SetTextBold(value: boolean);
  begin
    rang.Font.Bold:=value;
  end;
  
  procedure Design.SetTextSize(value: real);
  begin
    rang.Font.Size:=value;
  end;
  
  procedure Design.SetCellHeight(value: real);
  begin
    rang.RowHeight:=value;
  end;

  procedure Design.Centralize;
  begin
    rang.HorizontalAlignment:=Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
    rang.VerticalAlignment:=Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
  end;

  procedure Design.SetBorder(b: BorderRoute; value: boolean);
  begin
    if value then
      rang.Borders[b].LineStyle:=Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    else 
      rang.Borders[b].LineStyle:=Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
  end;

  procedure Design.SetBorderColor(b: BorderRoute; value: color);
  begin
    rang.Borders[b].Color:=colortranslator.ToOle(value);
  end;
    
  procedure Design.AutoSize;
  begin
    rang.Columns.AutoFit;
    rang.Rows.AutoFit;
  end;
  
  procedure Design.CellFormat(format: FormatType);
    begin
      if format=formattype.General then
        rang.NumberFormat:='general'
      else if format=formattype.Text then
        rang.NumberFormat:='@'#0
      else if format=formattype.Date then
        rang.NumberFormat:='m/d/yyyy'
      else if format=formattype.Money then
        rang.NumberFormat:='#,##0 $'
      else if format=formattype.Number then
        rang.NumberFormat:='0'#0
      else if format=formattype.Percent then
        rang.NumberFormat:='0%';
    end;  
    
  procedure Design.CellFormat(format: FormatType; number: integer);
  begin
    if (number<1) or (format=formattype.Date) or (format=formattype.General) or (format=formattype.Money) or (format=formattype.Text) then
      begin
        Cellformat(format);
        exit;
      end;
      
    var s:='0.';
    for var i:=1 to number do
      s+='0';
    
    if format=formattype.Number then
      rang.NumberFormat:=s
    else if format=formattype.Percent then
      rang.NumberFormat:=s+'%';
  end;

  function Design.GetBorder(b: BorderRoute): boolean;
    begin
      result:=if rang.Borders[b].LineStyle.ToString='-4142' then
        false
      else
        true;
    end;

end.