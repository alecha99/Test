(*
  07.02.2018
  Перелыгин А.Н.
  Удобная оболочка формирования отчетов по шаблонам

  Возможности:
  1. Заполнение одной и более таблиц , отдельных ячеек на одном и более вкладок шаблона,
  без требований к порядку заполнения.

  2. Удаление/зачистка шаблонных элементов в автоматическом режиме, при открытии результата.

  3. Работа с размножаемыми столбцами для каждой таблицы, и обеспечение единообразного представления,
  смещаемых и размножаемых столбцов.

  4. Возможность вывода итогов суммирования отдельных строк.

  5. Работа с одной или более вкладок шаблонов.

  6. Возможность создавать и работать с несколькими копиями вкладок,
  при условии последовательного заполнения каждой.

*)

unit uFlexCelDocument;

interface

uses
  Windows,
  Forms,
  SysUtils,
  System.Classes,
  Graphics,
  Controls,
  Math,
  StrUtils,
  Generics.Collections,
  Generics.Defaults,
  uFlexCelReport,
  _UXlsAdapter.XlsFile,
  _UCoreEnums.TFlxInsertMode,
  _ExcelAdapter.ExcelFile,
  VCL.FlexCel.Core,
  uExceptions,
  Variants,
  CodeList,
  System.Types;

{$REGION ' Инструкция важная к прочтению ... '}
(*
  1. Объекты возвращаемые компонентом являются только шаблонами и зачищаются при удалении  родительского объекта.
  В свою очередь только они работают с итоговыми листами результата, определяя сетку координат относительно текущего положения.
  Все шаблонные элементы (строки, вкладки, столбики) будут корректно автоматически удалены объектом и не требуют дополнительных действий.
  Все шаблонные элементы (строки, вкладки, столбики) должны быть добавлены хотя бы один раз на каждом листе для заполнения (Шаблонные элементы не подлежат корректировки)
  Каждый шаблонный элемент имеет свой набор итераций на каждом добавленном листе и каждый лист имеет свой набор элементов.
  2. Не инициализированные вкладки будут автоматически удалены при выводе отчета
*)
{$ENDREGION}

type

  TSheetTemplate = class;
  TRowTable = class;
  TTableSheet = class;

{$REGION ' Базовые типы для внутреннего использования ....'}
  TCorrectionPos = (tcpLeftY, tcpLeftX, tcpRightY, tcpRightX);
  TTypeSheet = (ttsBase, ttsNoBase);
  TTypePoint = (ttpLeft, ttpRight);
  tBaseObj = (tboBase, tboCurrent);

  TBasePointCells = record
    FConst: string;
    FBasePoint: TPoint;
    FValue: Variant;
    procedure Assign(var APoint: TBasePointCells);
  end;

  TBaseRect = record
    FLeft: TBasePointCells;
    FRight: TBasePointCells;
    procedure AssignLeft(var APoint: TBasePointCells);
    procedure AssignRight(var APoint: TBasePointCells);
  end;
{$ENDREGION}

  // Разамножение/копирование необходимого числа колонок таблицы со сдвигом остатка таблицы в право
  TExtColumns = class(TComponent)
  strict private
    FOwner: TTableSheet;
    FIDColumn: Integer;
    FBaseColumn: Integer;
    FCountColumn: Integer;
    FCurrentIter: Integer;

    // Переменные работающие с текущей строкой
    FTable: TTableSheet;
    FColumn: TExtColumns;
    FSheet: TSheetTemplate;
  private
    FIterariton: Integer;

    // конструктор для базового элемента
    constructor Create(AOwner: TTableSheet; const AOneColumn: Integer;
      const ACountColumn: Integer = 1); reintroduce; overload;
    // конструктор для элемента-наследника в дочерних таблицах
    constructor Create(AOwner: TTableSheet; ABaseColumn: TExtColumns); reintroduce; overload;

    procedure GetCurrentColumn;
    property BaseColumn: Integer read FBaseColumn;
    property CountColumn: Integer read FCountColumn;
    property Iterariton: Integer Read FIterariton;
    property CurrentIter: Integer read FCurrentIter;
    property IDColumn: Integer read FIDColumn;
  public
    procedure BeforeDestruction; override;

{$REGION ' Работа с итерациями обьекта '}
    /// <summary> Установить текущий номер итерации для колонок
    /// </summary>
    /// <param name="AIteration"> Номер итерации начиная с 1
    /// </param>
    /// <remarks> Это необходимо что бы таблица воспринимала указанный номер колонки в разрезе итерации (указанной в параметре)
    /// </remarks>
    procedure SetCurrentIter(const AIteration: Integer);
    /// <summary> Вставка новых колонок
    /// </summary>
    /// <remarks> Вставка следующих колонок и установка их как текущих
    /// </remarks>
    /// <returns> Номер текущей итерации (новой)
    /// </returns>
    function InsertColumn: Integer;
    /// <summary> Получить номер текущей итерации
    /// </summary>
    /// <remarks> Получаем именно текущий, а не последней.
    /// </remarks>
    /// <returns> Номер текущей итерации
    /// </returns>
    function GetIteration: Integer;
    /// <summary> Получить общее кол-во итераций
    /// </summary>
    /// <returns> Кол-во итераций
    /// </returns>
    function IteraritonCount: Integer;
{$ENDREGION}
  end;

  // Основной объект-контейнер при создании  необходим файл шаблона
  TFlexCelTemplate = class(TComponent)
  strict private
    FFLEX: TXLSFileReport;
    FOpenResult: Boolean;
    // Получить список листов шаблона
    procedure GetSheetList;
    // Очистить отчет от мусора
    procedure ClearReportGarbage;
  private
    FListBaseSheet: TObjectList<TSheetTemplate>;
    FListSheet: TObjectList<TSheetTemplate>;
  protected
    // Преобразование адресов Excel для внутреннего использования
    class function GetPointToAddr(const PointCell: TBasePointCells): string; virtual;
    class procedure GetAddrToPoint(const AXlsAddr: string; out APointCell: TBasePointCells); inline;
  public
{$REGION ' Создание/удаление объекта '}
    /// <summary>Создание базового обьекта отчета
    /// </summary>
    /// <param name="AOwner">Обьект Owner
    /// </param>
    /// <param name="AFileTemplate">Полный путь к шаблону файла Excel (наличие шаблона обязательно)
    /// </param>
    /// <remarks>  Если файла шаблона не найден или путь не определён будет сгенерировано исключение
    /// </remarks>
    constructor Create(AOwner: TComponent; const AFileTemplate: string); reintroduce;
    destructor Destroy; override;
{$ENDREGION}
{$REGION ' Работа с адресам формата Excel '}
    /// <summary> Получение положения ячейки строковому адресу формата Excel
    /// </summary>
    /// <param name="AAdress"> Строковое значение адреса ячейки формата Excel
    /// </param>
    /// <returns> TPoint Номер столбика и строки (числовые) от 0
    /// </returns>
    class function AddrToPoint(const AAdress: string): TPoint; inline;
    /// <summary>  Получение положения ячейки строковому адресу формата Excel
    /// </summary>
    /// <param name="APoint"> TPoint Номер столбика и строки (числовые) от 0
    /// </param>
    /// <returns> string Строковое значение адреса ячейки формата Excel
    /// </returns>
    class function PointToAddr(const APoint: TPoint): string; inline;
{$ENDREGION}
{$REGION ' Получить ссылку на базовые шаблоны '}
    /// <summary> Получение ссылки на шаблон листа по индексу
    /// </summary>
    /// <param name="AIndex"> Номер листа-шаблона, начиная с 1
    /// </param>
    /// <returns> TSheetTemplate Базовый объект листа-шаблона
    /// </returns>
    /// <remarks> Sheet создается автоматически при открытии шаблона
    /// так что возвращается ссылка только уже на созданный объект.
    /// <br> Все листы шаблонов при открытии файла будут удалены </br>
    /// </remarks>
    function GetSheetByIndex(const AIndex: Integer): TSheetTemplate;
    /// <summary> Получение ссылки на шаблон листа по названию листа
    /// </summary>
    /// <param name="ABaseName">Наименование листа-шаблона
    /// </param>
    /// <returns>TSheetTemplate Базовый объект листа-шаблона
    /// </returns>
    /// <remarks> Sheet создается автоматически при открытии шаблона
    /// так что возвращается ссылка только уже на созданный объект.
    /// <br> Все листы шаблонов при открытии файла результата будут удалены!</br>
    /// </remarks>
    function GetSheetByName(const ABaseName: string): TSheetTemplate;
{$ENDREGION}
{$REGION ' Открыть файл результата }
    /// <summary> OpenResult Открыть готовый отчет
    /// </summary>
    /// <remarks> После открытия отчета в Excel, работа с отчетом более недопустимо, так как
    /// <br> элементы-шаблоны будут удалены, а созданные листы проходят очистку от мусора.</br>
    /// </remarks>
    procedure OpenResult;
{$ENDREGION}
    // Основной компонет обертки (Для дополнительного функционала)
    property FlexCel: TXLSFileReport read FFLEX;
  end;

  (* Указатель на прямую ячейку создается только при инициализации *)
  TCellSheet = class(TComponent)
  strict private
    FOwner: TSheetTemplate;
    FBasePos: TBasePointCells;
    FIncY: Integer;
    FDef: Variant;
    FIDCell: Integer;
    // Переменные, работающие с текущей строкой
    FCell: TCellSheet;
    FSheet: TSheetTemplate;
    function GetPosition: TBasePointCells;
  protected
    constructor Create(AOwner: TComponent; ABaseSheet: TSheetTemplate; const ABasePos: TCellSheet;
      const AIDCell: Integer); reintroduce; overload; virtual;
  private
    constructor Create(AOwner: TComponent; ABaseSheet: TSheetTemplate;
      const ABasePos: TBasePointCells; const AIDCell: Integer); reintroduce; overload; virtual;

    procedure IncY(const AStep: Integer = 1);

    property IDCell: Integer read FIDCell;
    property Position: TBasePointCells read GetPosition;
    property Def: Variant read FDef write FDef;
    property BasePos: TBasePointCells read FBasePos;
    procedure GetCurrentElement;
  public
    procedure BeforeDestruction; override;
{$REGION ' Работа с ячейками Excel для пользователя '}
    /// <summary> SetValue установить значение ячейки
    /// </summary>
    /// <param name="AValue"> Устанавливаемое значение
    /// </param>
    procedure SetValue(const AValue: Variant);
    /// <summary> Установить цвет заливки ячейки
    /// </summary>
    /// <param name="AColor"> Цвет заливки
    /// </param>
    procedure SetColor(const AColor: TColor);
    /// <summary> Нарисовать штрих-код сверху ячейки (на документе по положению ячейки)
    /// </summary>
    /// <param name="ABarCode"> Int64 штрих-код
    /// </param>
    /// <param name="AForeColor"> TColor цвет штрих-кода
    /// </param>
    procedure SetBarCodeEAN13Image(ABarCode: Int64; AForeColor: TColor; aDy, aDx: Integer);
    /// <summary> Записать формулу суммы для всех итераций, любой шаблонной строки любой таблицы текущего листа
    /// </summary>
    /// <param name="AInRow"> TRowTable Объект шаблонной строки таблицы
    /// </param>
    /// <param name="AInColumn"> Номер колонки таблицы начиная с 1
    /// </param>
    /// <param name="ADefault"> Значенние по умолчанию, если строка на листе не имела ни одной итерации
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer; const ADefault: Variant);
    /// <summary> Получить указатель на ячейку по осям Х,Y
    /// </summary>
    /// <returns> TPoint текущее положение ячейки на текущем sheet
    /// </returns>
    function GetCell: TPoint;
{$ENDREGION}
  end;

  (* Строка таблицы  создается создается только при инициализации *)
  TRowTable = class(TComponent)
  strict private
    FOwner: TTableSheet;
    FBaseRow: Integer;
    FListRow: TList<Integer>;
    FIDRow: Integer;
    // Переменные, работабщие с текущей строкой
    FTable: TTableSheet;
    FRow: TRowTable;
    FSheet: TSheetTemplate;
    // Получить текущий элемент с адресами
    procedure GetCurrentElement;
  private
    FCorrenTRow: Integer;

    // создание дочернего object на основании базового object
    constructor Create(AOwner: TTableSheet; const ABaseRow: TRowTable); reintroduce; overload;

    // Получить текущий object, для текущей sheet и активация вкладки
    function CurrenTRow: Integer;
    // Получение строки формулы для дочернего текущего object
    function GetFormula(const AColumn: Integer; AListIterations: TIntegerList;
      out AValid: Boolean): string;
    // Создание базового object, на основании координат
    constructor Create(AOwner: TTableSheet; const ABaseRow: Integer; const AIDRow: Integer);
      reintroduce; overload;

    property BaseRow: Integer read FBaseRow;
    procedure CurretBasePos(const AOfSet: Integer);
    property IDRow: Integer read FIDRow;
    property LisTRow: TList<Integer> read FLisTRow;
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;

{$REGION ' Стилизация ячеек строк таблицы '}
    /// <summary> Установить значение ячейки для колонки текущей строки
    /// </summary>
    /// <param name="ANumColumn"> Номер колонки
    /// (Внимание номерация колонок таблицы с 1)
    /// </param>
    /// <param name="AValue"> устанавливаемое значение
    /// </param>
    procedure SetValue(const ANumColumn: Integer; const AValue: Variant);
    /// <summary> Записать формулу суммы для всех итераций, любой шаблонной строки любой таблицы текущего листа
    /// для определенной ячейки текущей строки.
    /// </summary>
    /// <param name="AInRow"> TRowTable Объект шаблонной строки таблицы, по которой выводим формулу
    /// </param>
    /// <param name="AInColumn"> Номер колонки строки, по которой выводим формулу
    /// </param>
    /// <param name="ADefault"> Значение, вводимое в случае если строка не имела не одной итерации
    /// </param>
    /// <param name="AResColumn"> Номер колонки для текущей строки куда вписываем формулу
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer; const ADefault: Variant;
      const AResColumn: Integer = -1); overload;
    /// <summary> Записать формулу суммы для листа итераций, любой шаблонной строки любой таблицы текущего листа
    /// для определеной ячейки текущей строки.
    /// </summary>
    /// <param name="AInRow"> TRowTable Объект шаблонной строки таблицы, по которой выводим формулу
    /// </param>
    /// <param name="AInColumn"> Номер колонки строки, по которой выводим формулу
    /// </param>
    /// <param name="AListIterations">  TIntegerList лист выбранных итераций, для строки по которой создаем формулу
    /// </param>
    /// <param name="ADefault"> Значение вводимое в случае если строка не имела не одной итерации
    /// </param>
    /// <param name="AResColumn"> Номер колонки для текущей строки куда вписываем формулу
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer;
      AListIterations: TIntegerList; const ADefault: Variant;
      const AResColumn: Integer = -1); overload;
    /// <summary> Установить определенный шрифт для ячейки (необходим для выделения записей)
    /// </summary>
    /// <param name="AFonts"> Устанавливаемый шрифт
    /// </param>
    /// <param name="AColumn"> Номер колонки  (Внимание номерация колонок таблицы с 1)
    /// </param>
    /// Вывод формулы в ячейку
    procedure SetFont(const AFonts: TXLSFontData; const AColumn: Integer = 0);
    /// <summary> Установить определенный цвет шрифта для ячейки (необходим для выделения записей)
    /// </summary>
    /// <param name="AColor"> Устанавливаемый цвет шрифта
    /// </param>
    /// <param name="AColumn"> Номер колонки (Внимание номерация колонок таблицы с 1)
    /// </param>
    procedure SetColorFont(const AColor: TColor; const AColumn: Integer); overload;
    /// <summary> Установить определенный цвет шрифта для всех ячеек строки таблицы (необходим для выделения записей)
    /// </summary>
    /// <param name="AColor"> Устанавливаемый цвет шрифта
    /// </param>
    /// <remarks>  Цвет шрифта будет изменён от первой до последней колонки таблицы,
    /// по расчету крайних точек самой таблицы
    /// </remarks>
    procedure SetColorFont(const AColor: TColor); overload;
    /// <summary> Установить определенный цвет заливки ячейки (необходим для выделения записей)
    /// </summary>
    /// <param name="ANumColumn"> Номер колонки (Внимание номерация колонок таблицы с 1)
    /// </param>
    /// <param name="AColor"> Устанавливаемый цвет
    /// </param>
    procedure SetColor(const ANumColumn: Integer; const AColor: TColor); overload;
    /// <summary> Установить определенный цвет заливки для всех ячеек строки таблицы (необходим для выделения записей)
    /// </summary>
    /// <param name="AColor"> Устанавливаемый цвет шрифта
    /// </param>
    /// <remarks>  Цвет заливки будет изменён от первой до последней колонки таблицы,
    /// по расчету крайних точек самой таблицы
    /// </remarks>
    procedure SetColor(const AColor: TColor); overload;
{$ENDREGION}
{$REGION ' Работа со строками таблиц '}
    /// <summary> Вставка/копирование строки из базового элемента
    /// </summary>
    /// <remarks> Работа возможна только если строка хотя бы раз была скопирована,
    /// работа с шаблоном вызовет exception
    /// </remarks>
    /// <returns> integer Возвращается именно номер новой итерации, а не номер новой строки,
    /// при многолинейном документе следует это учесть!
    /// </returns>
    function InsertRow: Integer;
    /// <summary> Получить номер строки для строки по определённой итерации
    /// </summary>
    /// <param name="AIteration"> Номер итерации
    /// </param>
    /// <returns> integer Номер строки, начиная с 1
    /// </returns>
    function GetNumRow(const AIteration: Integer = -1): Integer;
    /// <summary> Получить ко-во итераций для текущего листа
    /// </summary>
    /// <returns> integer кол-во итераций
    /// </returns>
    function IterationCount: Integer;
    /// <summary> Получить ко-во итераций для текущего листа
    /// </summary>
    /// <returns> integer кол-во итераций
    /// </returns>
    function RowCount: Integer;
    /// <summary> Получить текущую итерацию
    /// </summary>
    /// <returns> integer номер итерации
    /// </returns>
    function GetIteration: Integer;
    /// <summary> Обьеденение ячеек в одной строки
    /// </summary>
    /// <param name="ABegColumn"> Номер начальной колонки
    /// </param>
    /// <param name="AEndColumn"> Номер конечной колонки
    /// </param>
    /// <remarks> Только склеивание колонок (Внимание все нумерации колонок таблиц начинаются с 1)
    /// </remarks>
    procedure SetColumnMerge(const ABegColumn: Integer = -1; const AEndColumn: Integer = -1);
    /// <summary> Объединение ячеек в строки и ввод в нее значения
    /// </summary>
    /// <param name="ABegColumn"> Номер начальной колонки
    /// </param>
    /// <param name="AEndColumn"> Номер конечной колонки
    /// </param>
    /// <param name="AValue"> Устанавливаемое значение
    /// </param>
    /// <remarks> Внимание все номерации колонок таблиц начинаються с 1
    /// </remarks>
    procedure SetValueAndMerge(const ABegColumn, AEndColumn: Integer; const AValue: Variant);
{$ENDREGION}
  end;

  (* Таблица создается только при инициализации *)
  TTableSheet = class(TComponent)
  strict private
    FOwner: TSheetTemplate;
    FRectangle: TBaseRect;
    FOffsetBase: array [0 .. 3] of Integer;
    FRow: TObjectList<TRowTable>;
    FExtCol: TObjectList<TExtColumns>;
  private
    FOffSet: array [0 .. 3] of Integer;
    FNumTable: Integer;
    FLasTRow: Integer;
    FRowHistory: TList<Integer>;
    FOffSetTableName: Integer;
    // Контроль удаления таблицы
    FDropDef: Boolean;
    FDropTable: Boolean;

    // Создание базовой таблицы по абсалютным координатам
    constructor Create(AOwner: TSheetTemplate; const ALeftPos, ARigthPos: TBasePointCells;
      const ANumTable: Integer = -1); reintroduce; overload;

{$REGION ' Функции экспортирующие функционал от базовых объектов в текущие элементы '}
    function ExtColumnCount: Integer;
    function GetFormula(const AColumn: Integer; AListIterations: TList<Integer>): string;
    // Добавление полей и колонок
    function IncRow(ABaseSheet: TSheetTemplate; ARow: TRowTable): Integer;
    Procedure IncColumn(ABaseSheet: TSheetTemplate; AColumn: TExtColumns);
    // Установка цветов заливки и штрифта на конечном object (Импортированный из Row)
    procedure SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable; const ANumColumn: Integer;
      const AColor: TColor); overload;
    procedure SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      const AColor: TColor); overload;
    // Установка штрифта на конечном object (Импортированный из Row)
    procedure SetFont(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      AFonts: TXLSFontData);
    // Установка значений ячеек на конечном object (Импортированный из Row)
    procedure SetValue(ABaseSheet: TSheetTemplate; ARow: TRowTable; const ANumColumn: Integer;
      const AValue: Variant);
    // Установка конечном object (Импортированный из Row)
    procedure SetFormula(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      const AFormula: string);
    // Корректировка смещения таблицы в период добавления строк
    procedure incOfSeTRow;
    // Добавление базовых строк в таблицу
    procedure AddBaseRow(ABaseRow: TRowTable); overload;
    procedure AddBaseExtColumn(ABaseColumn: TExtColumns); overload;
    // Перегрузка строк и колонок
    procedure OverloadRowCol(ARecipient: TTableSheet);
    // определение местоположения строки
    function GetDesineRow(const ATypeRow: tBaseObj; const ABaseRow: Integer): Integer;
    // Определение координат полей
    function GetDesineColumn(const ACoumn: Integer): Integer;
    // Получить текущее кол-во колонок в таблице
    function GetColumnCount: Integer;
    // Определение текущих координат таблицы
    function GetDesineTable(const ATypePos: TCorrectionPos): Integer;
    // Получаем значение текущего обьекта
    function GeTRow(const AIDRow: Integer): TRowTable;
    // Получить текущий обьект размножаемой колонки
    function GetColumns(ABaseColun: TExtColumns): TExtColumns;
    // Получаем базовое место-положение таблицы
    function GetBasePosition(const ATypePoint: TTypePoint): TBasePointCells;

    // Склеить Column на строки
    procedure SetColumnMerge(ABaseSheet: TSheetTemplate; ARow: TRowTable;
      const ABegColumn, AEndColumn: Integer);
    property SheetTemplate: TSheetTemplate read FOwner;
    function Position: TPoint;
    function FPositionRigth: TPoint;
    procedure CleaningTable;
    function LastRow: Integer;
    procedure SetCurrentSheet(ASheet: TSheetTemplate);
    function GetCurrentSheet: TSheetTemplate;
    function GetCurrentTable(ASheet: TSheetTemplate): TTableSheet;
    function GetCurrenTRow(ARow: TRowTable; const AIteration: Integer): Integer;

    // Номер таблицы по порядку создания
    property NumTable: Integer read FNumTable;
{$ENDREGION}
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;

{$REGION ' пользовательские функции работы с таблицами '}
    /// <summary> Сделать выпадающий список по строки
    /// </summary>
    /// <param name="ARow"> Номер Первой строки таблицы (Он будет оставатся видимым)
    /// </param>
    /// <param name="AEndNumRow"> Номер последней строки таблицы, он будет скрыт включитьельно
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br> 1. Позволяется пока только один уровень складывания.  </br>
    /// <br> 2. Берутся не итерации, а именно готовые строки таблицы, что бы вернуть текущую </br>
    /// строку у TRowTable, воспользуйтесь функцией GetNumRow у TRowTable
    /// </remarks>
    procedure SetGroupRows(ARow: TRowTable; const AEndNumRow: Integer);
    /// <summary> Вернуть текущий адрес на определённую ячейку определённой строки
    /// (Номерация не по итерации а по последовательности)
    /// </summary>
    /// <param name="ARow"> Номер строки таблицы
    /// </param>
    /// <param name="ANumColumn"> Номер колонки таблицы (Номерация колонок таблиц с 1)
    /// </param>
    /// <returns> TPoint нумерация по осям x,y
    /// </returns>
    /// <remarks>
    /// <red> Важное замечание: </red> Берутся не итерации, а именно готовые строки таблицы, что бы вернуть текущую
    /// строку у TRowTable воспользуйтесь функцией GetNumRow у TRowTable.
    /// Функция так же Sheet для текущей таблицы делает текущим.
    /// </remarks>
    function GetCell(const ANumRow: Integer; const ANumColumn: Integer = -1): TPoint; overload;
    /// <summary> Получить углы таблицы (адресс одного из углов таблицы)
    /// <param name="ATypePos"> Тип угла
    /// <br> ttpLeft= левый верхний угол ttpRight= нижний правый угол</br>
    /// </param>
    /// <returns><b>Returns: TPoint</b> нумерация по осям x,y
    /// </returns>
    /// </summary>
    /// <remarks>
    /// Адреса рассчитываются и для TExtColumns, и всех текущих TRowTable,
    /// после очистки шаблонных строк адреса будут смещены.
    /// </remarks>
    function GetCell(const ATypePos: TTypePoint): TPoint; overload;
    /// <summary> Инициализировать строку шаблона в таблице
    /// </summary>
    /// <param name="ANumRow"> Номер строки от текущего верхнего угла таблицы
    /// </param>
    /// <returns> TRowTable Cсылка на базовую строку для таблицы
    /// </returns>
    /// <remarks><red> Важное замечание: </red>
    /// <br> 1. Cмещение адреса строки шаблона происходит и при корректировки смещения таблицы.</br>
    /// <br> 2. В случае если таблица имеет более одной шаблонной строки но при этом небыла заполненной,
    /// зачистка Sheet оставит только одну и добавленную первой. </br>
    /// <br> 3. Удалять возвращаемый object нельзя. </br>
    /// <br> 4. Удаление базовых строк из таблицы обеспечивает сама таблица.</br>
    /// </remarks>
    function AddBaseRow(const ANumRow: Integer): TRowTable; overload;
    /// <summary> Инициализировать колонки шаблона в таблице
    /// </summary>
    /// <param name="ABegColumn"> Номер первой размножаемой колонки
    /// </param>
    /// <param name="AACountColumn"> Номер крайней размножаемой колонки(для размножения блоками)
    /// </param>
    /// <returns> TExtColumns cсылка на базовый размножаемый столбик для таблицы
    /// </returns>
    /// <remarks> <red> Важное замечание: </red>
    /// <br> 1. Нумерация колонок остается неизменной для любого кол-во размножений(как если бы это были обычные столбики)
    /// а навигация по размноженным столбикам прооизводиться путем установки теущей итерации для TExtColumns (ф.SetCurrentIter)</br>
    /// <br> 2. В случае если таблица имеет более одной шаблонной строки, но при этом небыла заполненной,
    /// зачистка Sheet удалит все TExtColumns. </br>
    /// <br> 3. Удалять возвращаемый object нельзя </br>
    /// <br> 4. Удаление базовых TExtColumns из таблицы обеспечивает сама таблица </br>
    /// </remarks>
    function AddBaseExtColumn(const ABegColumn: Integer; const ACountColumn: Integer = 1)
      : TExtColumns; overload;
    /// <summary> Скорректировать местоположения таблицы
    /// <param name="AView"> Вид угла смещения
    /// tcpLeftY= Левый верхний угол по оси Y, tcpLeftX= Левый верхний угол по оси X,
    /// tcpRightY= Правый нижний угол по оси Y,  tcpRightX= Правый нижний угол по оси X
    /// </param>
    /// <param name="AValue"> Величина смещения смещения
    /// </param>
    /// </summary>
    procedure CorretionBasePos(const AView: TCorrectionPos; const AValue: Integer);
    /// <summary> Заполнить титульную часть таблицы
    /// <param name="AColumn"> Номер колонки в таблице (начиная с 1), независимо от отступа таблицы от края
    /// </param>
    /// <param name="AValue"> Значение ячейки в таблице
    /// </param>
    /// <param name="ARowTitle"> Если строк в титульной части более 1, то можно указать считая от верхней
    /// </param>
    /// </summary>
    procedure SetTitleColumns(const AColumn: Integer; const AValue: Variant;
      const ARowTitle: Integer = 0);
    /// <summary> Поставить итог по таблице в строки результата
    /// <param name="ARow"> TRowTable все итерации которого будут вписаны формулой (Для простых отчетов)
    /// </param>
    /// <param name="AColumn">Номер колонки в таблице (начиная с 1)
    /// </param>
    /// <param name="ADefault">Значение по умолчанию, если на текущем листе TRowTable небыло ни одной итерации
    /// </param>
    /// <param name="RowResult">Если итоговых строк более одной то можно указать номер(Отсчет от последней добавленой TRowTable)
    /// </param>
    /// </summary>
    procedure SetSumTable(ARow: TRowTable; const AColumn: Integer; const ADefault: Variant;
      const RowResult: Integer = 1);
    /// <summary> Удалить таблицу, при условии что нет ни одной записи, учитывая строку описания таблицы
    /// <param name="OffSetTName"> Сколько строк в верх стоит удалить от верхнего угла (Обычно это наименование таблицы)
    /// </param>
    /// </summary>
    /// <remarks><red> Важное замечание: </red>
    /// <br>1. Данное свойство устанавливается на все наследники шаблона таблицы, подсчет итераций производиться для каждой новой страницы.</br>
    /// <br>2. Удаление смещает область документа на необходимое кол-во строк (Это именно удаление строк а не зачистка)</br>
    /// </remarks>
    procedure DropDef(const OffSetTName: Integer = 0);
    /// <summary> Удалить таблицу в любом случае на текущем Sheet
    /// <param name="OffSetTName"> Сколько строк в верх стоит удалить от верхнего угла (Обычно это наименование таблицы)
    /// </param>
    /// </summary>
    /// <remarks><red> Важное замечание: </red>
    /// <br>1. Удаление смещает область документа на необходимое кол-во строк (Это именно удаление строк а не зачистка).</br>
    /// <br>2. Таблица на текущем Sheet будет удалена в любом случае (Даже если в ней были добавлены строки)</br>
    /// </remarks>
    procedure DropTable(const OffSetTName: Integer = 0);
    /// <summary> Склеить столбцы строки таблицы
    /// <param name="ABegNumRow"> Начальная строка склеивания
    /// </param>
    /// <param name="AEndNumRow"> Крайняя строка склеивания
    /// </param>
    /// <param name="ABeginCol"> Начальная колонка склеивания
    /// </param>
    /// <param name="AEndCol"> Крайняя колонка склеивания
    /// </param>
    /// </summary>
    /// <remarks>
    /// <red> Важное замечание: </red>
    /// <br> 1. Метод склеивает именно строки таблицы, а не номера итераций строк.</br>
    /// <br> 2. Номерация колонок начинается с 1 </br>
    /// <br> 3. Механизм расчитает начальную и конечную колонки склеивания, при размножаемых колонках возможен казус(Шаблонные колонки будут удалены)</br>
    /// </remarks>
    procedure SetMergeGroupRow(const ABegNumRow, AEndNumRow: Integer;
      const ABeginCol, AEndCol: Integer);
{$ENDREGION}
  end;

  (* Sheet создается только при инициализации и для каждого листа-шаблона должен быть свой *)
  TSheetTemplate = class(TComponent)
  strict private
    FBaseName: string;
    FSheetName: string;
    FBaseSheet: Integer;
    FSheet: TList<Integer>;
    FListOwner: TList<Integer>;
    FOwner: TFlexCelTemplate;
    FTable: TObjectList<TTableSheet>;
    FListCell: TObjectList<TCellSheet>;
    FSheetCurrent: TSheetTemplate;
    FSheetNum: Integer;
    FSheetType: TTypeSheet;
    // Основное добавление таблицы без поиска в документе по готовым координатам
    function AddBaseTable(const ALeftPos, ARigthPos: TBasePointCells; const ANumTable: Integer = -1)
      : TTableSheet; overload;
    // Основное добавление ссылок по координатам
    function AddBaseCells(ABaseSheet: TSheetTemplate; const AFBasePos: TBasePointCells;
      const ADefault: Variant; const ASheetType: TTypeSheet; const AIDCell: Integer)
      : TCellSheet; overload;
  protected
    // Меняем текущую вкладку по умолчанию соответствубщую данному шаблону
    procedure ReassignSheet(const ANum: Integer); virtual;
  private
    // Конструкторы и деструкторы только в привате
    constructor Create(AOwner: TFlexCelTemplate; const AIndex: Integer; const ABaseName: string;
      const ATypeSheet: TTypeSheet = ttsNoBase); reintroduce;

    // Зачистка Sheet от базовых элементов
    procedure DeleteBaseRow(ATable: TTableSheet; ARow: TRowTable);
    procedure DeleteBaseColumns(ATable: TTableSheet; AColumns: TExtColumns);
    // Получение позиции указателя на вкладке по константе
    function GetPosConst(const AConst: string): TBasePointCells;
    function IncRow(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable): Integer;
    procedure IncColumn(ABaseSheet: TSheetTemplate; ATable: TTableSheet; AColumn: TExtColumns);
    // Закрасить ячейку у строки
    procedure SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ANumColumn: Integer; const AColor: TColor); overload;
    // Закрасить строку
    procedure SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColor: TColor); overload;
    // Сменить цвет шрифта
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; const AColor: TColor); overload;
    procedure SetFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; AFonts: TXLSFontData);
    // Склеить строки в таблице по колонкам
    procedure SetColumnMerge(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ABegColumn, AEndColumn: Integer);
    // присвоить значение строки
    procedure SetValue(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ANumColumn: Integer; const AValue: Variant); overload;
    // Присвоить значение ячейки с константой TCellSheet
    procedure SetValue(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AValue: Variant); overload;
    // Присвоить значение ячейки по адресу
    procedure SetValue(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
      const AValue: Variant); overload;
    // Присвоить значение цвета ячейки с константой TCellSheet
    procedure SetColor(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AColor: TColor); overload;
    // Нарисовать ШК
    procedure SetBarCodeEAN13Image(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const ABarCode: Int64; const AForeColor: TColor; const aDy, aDx: Integer);
    // Ввести форулу
    procedure SetFormula(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; const AFormula: string); overload;
    procedure SetFormula(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
      const AFormula: string); overload;
    procedure SetFormula(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AFormula: String); overload;
    // Присвоить значение по указателю
    procedure SetValue(const AFBasePos: TBasePointCells; const AValue: Variant;
      const ATypeSheet: TTypeSheet); overload;

    procedure DropTable(ATable: TTableSheet);
    function GetTable(const ANumTable: Integer): TTableSheet;
    function GetCell(const AIDCell: Integer): TCellSheet;
    // зачистка шаблона производиться автоматически
    procedure CleaningSheet;
    // Установка текущего Sheet (Внутренений метод)
    procedure SetSheet(ASheet: TSheetTemplate);
    // добавление элемента в дочерние Sheet
    procedure AddBaseCells(ABaseSheet: TSheetTemplate; ACells: TCellSheet); overload;

    property FlexCelTemplate: TFlexCelTemplate read FOwner;
    // тип вкладки
    property SheetType: TTypeSheet read FSheetType;
    // Текущая страница
    property Sheet: TSheetTemplate read FSheetCurrent;
    // Текущий номер вкладки
    property SheetNum: Integer read FSheetNum;
    // Ссылка на базовую вкладку
    property BaseSheet: Integer read FBaseSheet;
    property BaseName: string read FBaseName;

  protected
    // Прописать значение по константе
    procedure SetPointValue(const APoint: TPoint; const AValue: Variant);
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;
{$REGION ' Методы таблиц доступные для интеграции '}
    /// <summary> Добавление базовой TCellSheet с поиском по константе/Tаg (Инициализация)
    /// </summary>
    /// <param name="ATag"> Значение Tag(константы)
    /// </param>
    /// <param name="ADefault"> Значение ячейки по умолчанию
    /// </param>
    /// <returns>  TCellSheet Возвращаемый базовый TCellSheet
    /// </returns>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять только для базовых Sheet.  </br>
    /// <br>2. Все по умолчанию будут иметь ADefault </br>
    /// <br>3. object создает и очищает TSheetTemplate а вы работаете только с шаблоном</br>
    /// </remarks>
    function AddBaseCells(const ATag: string; const ADefault: Variant): TCellSheet; overload;
    /// <summary> Добавление базовой TCellSheet по указанному адресу Excel (Инициализация)
    /// </summary>
    /// <param name="AXlsAddr"> Адрес ячейки формата Excel
    /// </param>
    /// <param name="ADefault"> Значение ячейки по умолчанию
    /// </param>
    /// <returns> TCellSheet Возвращаемый базовый TCellSheet
    /// </returns>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять только для базовых Sheet.  </br>
    /// <br>2. Все по умолчанию будут иметь ADefault </br>
    /// <br>3. object создает и очищает TSheetTemplate а вы работаете только с шаблоном</br>
    /// </remarks>
    function AddBaseCellsByAddr(const AXlsAddr: string; const ADefault: Variant)
      : TCellSheet; overload;
    /// <summary> Добавление новой вкладки на основании шаблона
    /// </summary>
    /// <param name="ASheetName"> Строковое имя новой вкладки
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Если имя Sheet дублирует базовый? то базовый будет изменён, если дочерний то будет создан с префиксом #x где x номер дубликата имени.  </br>
    /// <br>2. При добавлении все компоненты Sheet (TCellSheet,TTableSheet...) переходят в исходное состояние </br>
    /// <br>3. Все итерации дочерних компонентов будут сброшены к 0</br>
    /// </remarks>
    procedure CopyAddSheet(const ASheetName: string);
    /// <summary> Добавление таблицы на текущую вкладку по координатам(Инициализация)
    /// </summary>
    /// <param name="ALeftRow"> Номер ячейки по оси Y для верхнего левого угла
    /// </param>
    /// <param name="ALeftCol"> Номер ячейки по oси X для верхнего левого угла
    /// </param>
    /// <param name="ARigthRow"> Номер ячейки по оси Y для нижнего правого угла
    /// </param>
    /// <param name="ARigthCol"> Номер ячейки по oси X для нижнего правого угла
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять, только для базовых Sheet.</br>
    /// <br>2. Инициализаций строк таблицы, делается в этом компоненте.</br>
    /// <br>3. object создает и очищает TSheetTemplate, а вы работаете только с шаблоном.</br>
    /// </remarks>
    function AddBaseTable(const ALeftRow, ALeftCol, ARigthRow, ARigthCol: Integer)
      : TTableSheet; overload;
    /// <summary> Добавление таблицы на текущую вкладку по двум константам/tag (Инициализация)
    /// </summary>
    /// <param name="ALeftConst">Константа/значение Tag, верхнего левого угла
    /// </param>
    /// <param name="ARigthConst">Константа/значение Tag, нижнего правого угла
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять только для базовых Sheet.</br>
    /// <br>2. Инициализаций строк таблицы делаеться в этом компоненте.</br>
    /// <br>3. object создает и очищает TSheetTemplate, а вы работаете только с шаблоном.</br>
    /// </remarks>
    function AddBaseTable(const ALeftConst, ARigthConst: string): TTableSheet; overload;
    /// <summary> Добавление таблицы на текущую вкладку по константе/tag и смещению(Инициализация)
    /// </summary>
    /// <param name="ALeftConst">Константа/значение Tag верхнего левого угла
    /// </param>
    /// <param name="ARigthOffset">Адреса смещения для правого нижнего угла по осям X,Y
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять только для базовых Sheet.</br>
    /// <br>2. Инициализаций строк таблицы, делается в этом компоненте.</br>
    /// <br>3. object создает и очищает TSheetTemplate, а вы работаете только с шаблоном.</br>
    /// </remarks>
    function AddBaseTable(const ALeftConst: string; ARigthOffset: TPoint): TTableSheet; overload;
    /// <summary> Добавление таблицы на текущую вкладку по координатам формата Excel(Инициализация)
    /// </summary>
    /// <param name="ALeftConst">Координаты верхнего левого угла
    /// </param>
    /// <param name="ARigthConst">Координаты нижнего правого угла
    /// </param>
    /// <remarks> <red> Важное замечание: </red>
    /// <br>1. Позволяет добавлять только для базовых Sheet.</br>
    /// <br>2. Инициализаций строк таблицы делается в этом компоненте.</br>
    /// <br>3. object создает и очищает TSheetTemplate, а вы работаете только с шаблоном.</br>
    /// </remarks>
    function AddBaseTableByAddr(const ALeftTop, ARigthBottom: string): TTableSheet;
{$ENDREGION}
  end;

implementation

{$REGION ' TBaseRect ....}

procedure TBaseRect.AssignLeft(var APoint: TBasePointCells);
begin
  Self.FLeft.Assign(APoint);
end;

procedure TBaseRect.AssignRight(var APoint: TBasePointCells);
begin
  Self.FRight.Assign(APoint);
end;
{$ENDREGION}
{$REGION ' TBasePointCells ....}

procedure TBasePointCells.Assign(var APoint: TBasePointCells);
begin
  APoint.FConst := Self.FConst;
  APoint.FValue := Self.FValue;
  APoint.FBasePoint.X := Self.FBasePoint.X;
  APoint.FBasePoint.Y := Self.FBasePoint.Y;
end;
{$ENDREGION}
{$REGION ' TFlexCelTemplate ....}

procedure TFlexCelTemplate.ClearReportGarbage;
var
  cycle: Integer;
  TmpSheet: TSheetTemplate;
begin
  for TmpSheet in FListSheet do
    TmpSheet.CleaningSheet;

  for cycle := 0 to Self.FListBaseSheet.Count - 1 do
  begin
    if (FListSheet.Count = 0) and (Self.FListBaseSheet[cycle].BaseSheet = 0) then
      Break
    else
      Self.FFLEX.DeleteSheet(Self.FListBaseSheet[cycle].BaseSheet);
  end;
end;

constructor TFlexCelTemplate.Create(AOwner: TComponent; const AFileTemplate: string);
begin
  inherited Create(AOwner);

  FOpenResult := False;

  if not FileExists(AFileTemplate) then
    raise EASDException.Create('Не найден файл шаблона отчета Excel ' + sLineBreak + '«' +
      AFileTemplate + '»');

  try
    FFLEX := TXLSFileReport.Create(AOwner, AFileTemplate);
  except
    on E: Exception do
      raise EASDException.Create('Ошибка при загрузке шаблона отчета: ' + sLineBreak + '«' +
        E.Message + '»' + sLineBreak + 'возможно он открыт другим пользователем');
  end;

  // Создаем контроль базовых листов необходимо учитывать все листы шаблона
  FListBaseSheet := TObjectList<TSheetTemplate>.Create(TComparer<TSheetTemplate>.Construct(
    function(const Left, Right: TSheetTemplate): Integer
    begin
      Result := CompareValue(Right.BaseSheet, Left.BaseSheet);
    end), True);

  // Создаем контроль базовых листов необходимо учитывать все листы шаблона
  FListSheet := TObjectList<TSheetTemplate>.Create(True);

  // Получаем листы шаблона докумета
  GetSheetList;
end;

destructor TFlexCelTemplate.Destroy;
begin
  FreeAndNil(FFLEX);
  FreeAndNil(FListSheet);
  FreeAndNil(FListBaseSheet);
  inherited;
end;

class procedure TFlexCelTemplate.GetAddrToPoint(const AXlsAddr: string;
out APointCell: TBasePointCells);
var
  Position: TPoint;
begin
  Position := TFlexCelTemplate.AddrToPoint(AXlsAddr);
  APointCell.FBasePoint.X := Position.X;
  APointCell.FBasePoint.Y := Position.Y - 1;
  APointCell.FValue := null;
  APointCell.FConst := AXlsAddr
end;

class function TFlexCelTemplate.GetPointToAddr(const PointCell: TBasePointCells): string;
begin
  Result := PointToAddr(PointCell.FBasePoint);
end;

function TFlexCelTemplate.GetSheetByIndex(const AIndex: Integer): TSheetTemplate;
begin
  if AIndex > FListBaseSheet.Count then
    raise EASDException.Create
      ('TFlexCelTemplate.GetSheetIndex Обращение к несуществующему листу шаблона');

  Result := FListBaseSheet[(FListBaseSheet.Count - 1) - AIndex + 1];
end;

function TFlexCelTemplate.GetSheetByName(const ABaseName: string): TSheetTemplate;
var
  iteration: Integer;
begin
  Result := nil;

  for iteration := 0 to FListBaseSheet.Count - 1 do
    if SameText(ABaseName, FListBaseSheet[iteration].BaseName) then
      Result := FListBaseSheet[iteration];

  if not Assigned(Result) then
    raise EASDException.Create
      ('TFlexCelTemplate.GetSheetName Обращение к несуществующему листу шаблона');
end;

procedure TFlexCelTemplate.GetSheetList;
var
  iteration, sh: Integer;
  tsh: TSheetTemplate;
begin
  sh := FFLEX.SheetCount;
  for iteration := 0 to sh - 1 do
  begin
    FFLEX.SheetIndex := iteration;
    tsh := TSheetTemplate.Create(Self, iteration, FFLEX.SheetName, ttsBase);
    tsh.Name := 'BaseSheet_' + IntToStr(iteration + 1);
    Self.FListBaseSheet.Add(tsh);
  end;
  Self.FListBaseSheet.Sort;
end;

procedure TFlexCelTemplate.OpenResult;
begin
  if Self.FOpenResult then
    raise EASDException.Create
      ('TFlexCelTemplate.OpenResult Повторное открытие результата, недопустимо');

  ClearReportGarbage;

  Self.FFLEX.ExportData(TXLSReportFunc.xfSaveToFileWODialog);
  Self.FOpenResult := False;
end;

class function TFlexCelTemplate.PointToAddr(const APoint: TPoint): string;
var
  Addr: string;
  X, Y: Integer;
begin
  X := APoint.X + 1;
  Y := APoint.Y;

  SetLength(Addr, 0);

  X := X - 1;
  while X >= 26 do
  begin
    Addr := chr(65 + (X mod 26)) + Addr;
    X := (X div 26) - 1;
  end;
  Result := chr(65 + X) + Addr + IntToStr(Y);
end;

class function TFlexCelTemplate.AddrToPoint(const AAdress: string): TPoint;
var
  Addr: string;
  i: Integer;
  W: Integer;
begin
  Addr := UpperCase(Trim(AAdress));

  SetLength(Addr, Length(Addr));
  Result.X := 0;
  for i := 1 to Length(Addr) do
  begin
    W := Ord(Addr[i]);
    if (W > 64) and (W < 91) then
      Result.X := (Result.X * 26) + (W - 64)
    else
    begin
      Addr := Copy(Addr, i);
      Break;
    end;
  end;

  // Сам эксел начинает отсчет с 0 а отображение с 1 по оси X
  Result.X := Result.X - 1;
  Result.Y := StrToIntDef(Addr, -1);

  Assert((Result.Y > -1) or (Result.Y > -1),
    'TFlexCelTemplate.AddrToPoint Ошибка формата координат ячейки Excel');
end;
{$ENDREGION}
{$REGION ' TSheetTemplate ....}

// Добавляем таблицу по поиску констант
function TSheetTemplate.AddBaseTable(const ALeftConst, ARigthConst: string): TTableSheet;
begin
  Result := AddBaseTable(GetPosConst(ALeftConst), GetPosConst(ARigthConst));
end;

// Добавляем таблицу по координатам
function TSheetTemplate.AddBaseTable(const ALeftRow, ALeftCol, ARigthRow, ARigthCol: Integer)
  : TTableSheet;
var
  Left, Right: TBasePointCells;
  lNum: Integer;
begin
  Left.FConst := '';
  Right.FConst := '';
  Left.FValue := null;
  Right.FValue := null;
  Left.FBasePoint.Y := ALeftRow;
  Left.FBasePoint.X := ALeftCol;
  Right.FBasePoint.Y := ARigthRow;
  Right.FBasePoint.X := ARigthCol;
  lNum := Self.FTable.Count + 1;
  Result := AddBaseTable(Left, Right, lNum);
end;

function TSheetTemplate.AddBaseTable(const ALeftPos, ARigthPos: TBasePointCells;
const ANumTable: Integer = -1): TTableSheet;
begin
  Result := TTableSheet.Create(Self, ALeftPos, ARigthPos, ANumTable);
  FTable.Add(Result);
  FTable.Sort;
end;

function TSheetTemplate.AddBaseCells(ABaseSheet: TSheetTemplate; const AFBasePos: TBasePointCells;
const ADefault: Variant; const ASheetType: TTypeSheet; const AIDCell: Integer): TCellSheet;
var
  num: Integer;
begin
  num := ifthen(ASheetType = ttsBase, Self.FListCell.Count + 1, AIDCell);

  Result := TCellSheet.Create(Self, ABaseSheet, AFBasePos, num);

  Result.Name := Self.Name + '_Cell' + IntToStr(num);

  Result.Def := ADefault;
  Self.FListCell.Add(Result);
  Self.FListCell.Sort;

  if ASheetType = ttsBase then
    SetValue(AFBasePos, ADefault, ttsBase);
end;

function TSheetTemplate.AddBaseCells(const ATag: string; const ADefault: Variant): TCellSheet;
var
  bpc: TBasePointCells;
begin
  bpc := GetPosConst(ATag);
  Result := Self.AddBaseCells(Self, bpc, ADefault, ttsBase, -1);
end;

procedure TSheetTemplate.AddBaseCells(ABaseSheet: TSheetTemplate; ACells: TCellSheet);
begin
  Self.AddBaseCells(ABaseSheet, ACells.BasePos, ACells.Def, ttsNoBase, ACells.IDCell);
end;

function TSheetTemplate.AddBaseCellsByAddr(const AXlsAddr: string; const ADefault: Variant)
  : TCellSheet;
var
  bpc: TBasePointCells;
begin
  TFlexCelTemplate.GetAddrToPoint(AXlsAddr, bpc);
  Result := Self.AddBaseCells(Self, bpc, ADefault, ttsBase, -1);
  bpc.FValue := ADefault;
end;

function TSheetTemplate.AddBaseTable(const ALeftConst: string; ARigthOffset: TPoint): TTableSheet;
var
  Left, Right: TBasePointCells;
  lNum: Integer;
begin
  Left := GetPosConst(ALeftConst);
  Right.FConst := '';
  Right.FValue := null;
  Right.FBasePoint.Y := ARigthOffset.Y;
  Right.FBasePoint.X := ARigthOffset.X;
  lNum := Self.FTable.Count + 1;
  Result := AddBaseTable(Left, Right, lNum);
  Result.Name := 'Sheet_' + IntToStr(Self.FSheetNum) + '_BaseTable' + IntToStr(lNum);
end;

function TSheetTemplate.AddBaseTableByAddr(const ALeftTop, ARigthBottom: string): TTableSheet;
var
  Left, Right: TBasePointCells;
  lNum: Integer;
begin
  lNum := Self.FTable.Count + 1;
  TFlexCelTemplate.GetAddrToPoint(ALeftTop, Left);
  TFlexCelTemplate.GetAddrToPoint(ARigthBottom, Right);
  Right.FBasePoint.X := Right.FBasePoint.X - Left.FBasePoint.X;
  Right.FBasePoint.Y := Right.FBasePoint.Y - Left.FBasePoint.Y;
  Result := AddBaseTable(Left, Right, lNum);
  Result.Name := 'Sheet_' + IntToStr(Self.FSheetNum) + '_BaseTable' + IntToStr(lNum);
end;

procedure TSheetTemplate.BeforeDestruction;
begin
  if (Owner = nil) or (csDestroying in Owner.ComponentState) then
    inherited
  else
    raise EASDException.Create
      ('TSheetTemplate.BeforeDestruction объекты данного класса удаляются только родителями');
end;

procedure TSheetTemplate.CleaningSheet;
var
  tmpTable: TTableSheet;
begin
  for tmpTable in FTable do
    if ((tmpTable.FDropDef) and (tmpTable.FRowHistory.Count = 0)) or (tmpTable.FDropTable) then
      Self.DropTable(tmpTable)
    else
      tmpTable.CleaningTable;
end;

procedure TSheetTemplate.CopyAddSheet(const ASheetName: string);
var
  sh: Integer;
  tsh: TSheetTemplate;
  SName: string;
  tmpTb, RlsTb: TTableSheet;
  tmpCell: TCellSheet;
begin
  FOwner.FlexCel.SheetIndex := FOwner.FlexCel.SheetCount - 1;

  if ASheetName = '' then
    SName := Self.FBaseName
  else
    SName := ASheetName;
  if SameText(SName, Self.FBaseName) then
  begin
    FOwner.FlexCel.SheetIndex := Self.FBaseSheet;
    FOwner.FlexCel.SheetName := 'RenameSheet-' + IntToHex(Self.FBaseSheet, 2);
    Self.FSheetName := FOwner.FlexCel.SheetName;
  end;

  FOwner.FlexCel.InsertAndCopySheet(FBaseSheet, SName, FOwner.FlexCel.SheetCount - 1);
  SName := FOwner.FlexCel.SheetName;
  sh := FOwner.FlexCel.SheetIndex;

  FSheet.Add(sh);
  FOwner.FlexCel.LoadVarListFromXLSFile;
  tsh := TSheetTemplate.Create(FOwner, Self.FBaseSheet, Self.BaseName);
  tsh.FSheetName := SName;
  tsh.FBaseName := Self.FBaseName;

  inc(Self.FSheetNum);

  FOwner.FlexCel.SheetIndex := FSheet[FSheet.Count - 1];
  tsh.FSheetNum := FOwner.FlexCel.SheetCount - 1;
  tsh.Name := 'CurrentSheet_' + IntToStr(FSheet.Count + 1) + '_ofBaseSheet_' +
    IntToStr(Self.FBaseSheet);

  FOwner.FListSheet.Add(tsh);
  Self.FListOwner.Add(FOwner.FListSheet.Count - 1);
  Self.FSheetCurrent := tsh;

  // Создаем указатели на заполняемую область таблиц
  for tmpTb in Self.FTable do
  begin
    RlsTb := tsh.AddBaseTable(tmpTb.GetBasePosition(ttpLeft), tmpTb.GetBasePosition(ttpRight),
      tmpTb.NumTable);
    RlsTb.Name := 'Sheet' + IntToStr(tsh.FSheetNum) + '_' + 'CurrentTable' +
      IntToStr(tmpTb.NumTable);
    RlsTb.FDropDef := tmpTb.FDropDef;
    tmpTb.OverloadRowCol(RlsTb);
  end;

  // Создаем указатели на константные ячейки
  for tmpCell in Self.FListCell do
    tsh.AddBaseCells(Self, tmpCell);
end;

constructor TSheetTemplate.Create(AOwner: TFlexCelTemplate; const AIndex: Integer;
const ABaseName: string; const ATypeSheet: TTypeSheet = ttsNoBase);
begin
  inherited Create(AOwner);
  FSheetCurrent := nil;
  FOwner := AOwner;
  FSheetType := ATypeSheet;
  Self.FBaseSheet := AIndex;
  Self.FBaseName := ABaseName;
  FSheet := TList<Integer>.Create;
  FSheetNum := 0;
  FListOwner := TList<Integer>.Create;

  // Список таблиц входящий в базовую форму выстраиваем для определения положения
  FTable := TObjectList<TTableSheet>.Create(TComparer<TTableSheet>.Construct(
    function(const Left, Right: TTableSheet): Integer
    begin
      Result := CompareValue(Right.Position.Y, Left.Position.Y);
    end), True);

  // Список одинночных ячеек констат
  FListCell := TObjectList<TCellSheet>.Create(TComparer<TCellSheet>.Construct(
    function(const Left, Right: TCellSheet): Integer
    begin
      Result := CompareValue(Left.BasePos.FBasePoint.Y, Right.BasePos.FBasePoint.Y);
      if Result <> 0 then
        Exit;
      Result := CompareValue(Right.BasePos.FBasePoint.X, Left.BasePos.FBasePoint.X);
    end), True);
end;

procedure TSheetTemplate.DeleteBaseColumns(ATable: TTableSheet; AColumns: TExtColumns);
var
  SrcRange: TXlsCellRange;
  Left, Right: TBasePointCells;
  rb, re, cb: Integer;
begin
  SetSheet(Self);

  // Определяем координаты таблицы
  Left := ATable.GetBasePosition(ttpLeft);
  Right := ATable.GetBasePosition(ttpRight);

  // высчитываем верхний и нижний угол
  rb := Left.FBasePoint.Y + ATable.FOffSet[0];
  re := Left.FBasePoint.Y + ATable.FOffSet[0] + Right.FBasePoint.Y + ATable.FRowHistory.Count + 1;
  cb := ATable.GetDesineColumn(AColumns.BaseColumn - 1) + 1;

  with TFlexCelTemplate(FOwner).FlexCel do
  begin
    SrcRange := TXlsCellRange.Create(rb, cb, re, (cb + (AColumns.CountColumn)) - 1);
    XF.DeleteRange(SrcRange, TFlxInsertMode.ShiftRangeRight);
  end;
end;

procedure TSheetTemplate.DeleteBaseRow(ATable: TTableSheet; ARow: TRowTable);
begin
  SetSheet(Self);
  TFlexCelTemplate(FOwner).FlexCel.DeleteRows(ATable.GetDesineRow(tboBase, ARow.BaseRow),
    ATable.GetDesineRow(tboBase, ARow.BaseRow));
end;

destructor TSheetTemplate.Destroy;
begin
  FTable.Free;
  FSheet.Free;
  FListOwner.Free;
  FListCell.Free;
  inherited;
end;

procedure TSheetTemplate.DropTable(ATable: TTableSheet);
var
  Left, Right: TBasePointCells;
begin
  Left := ATable.GetBasePosition(ttpLeft);
  Right := ATable.GetBasePosition(ttpRight);

  SetSheet(Self);
  TFlexCelTemplate(FOwner).FlexCel.DeleteRows(Left.FBasePoint.Y + ATable.FOffSet[0] -
    ATable.FOffSetTableName, Left.FBasePoint.Y + ATable.FOffSet[0] + Right.FBasePoint.Y +
    ATable.FRowHistory.Count);
end;

function TSheetTemplate.GetCell(const AIDCell: Integer): TCellSheet;
var
  tmp: TCellSheet;
begin
  Result := nil;
  for tmp in FListCell do
    if tmp.IDCell = AIDCell then
    begin
      Result := tmp;
      Break;
    end;

  if not Assigned(Result) then
    raise EASDException.Create('TSheetTemplate.GetCell не найден элемент константы');
end;

function TSheetTemplate.GetPosConst(const AConst: string): TBasePointCells;
var
  sh: Integer;
begin
  if Assigned(FSheetCurrent) then
    raise EASDException.Create
      ('TSheetTemplate.GetPosConst Инициализация всех констант возможно только в шаблоне и только до первой вставки');

  sh := FBaseSheet;

  Result.FConst := AConst;
  Result.FValue := null;
  with TFlexCelTemplate(FOwner).FlexCel do
  begin
    if not(SheetIndex = sh) then
    begin
      SheetIndex := sh;
      LoadVarListFromXLSFile;
    end;

    if not VRL.SearchVar(AConst, Result.FBasePoint.Y, Result.FBasePoint.X) then
      raise EASDException.Create
        ('Ошибка TSheetTemplate.GetPosConst Не найдена константа на форме Excel ' + AConst);

    // Зачищаем найденый указатель
    SetValue(Result.FBasePoint.Y, Result.FBasePoint.X, null);
  end;
end;

function TSheetTemplate.GetTable(const ANumTable: Integer): TTableSheet;
begin
  Result := nil;
  for Result in FTable do
    if Result.NumTable = ANumTable then
      Break;

  if (Result = nil) or not(Result.NumTable = ANumTable) then
    raise EASDException.Create('TSheetTemplate.GetTable в списке таблиц не найден искомый элемент');
end;

procedure TSheetTemplate.IncColumn(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
AColumn: TExtColumns);
var
  SrcRange: TXlsCellRange;
  Left, Right: TBasePointCells;
  rb, re, cb: Integer;
begin
  SetSheet(Self);

  // Определяем координаты таблицы
  Left := ATable.GetBasePosition(ttpLeft);
  Right := ATable.GetBasePosition(ttpRight);

  // высчитываем верхний и нижний угол
  rb := Left.FBasePoint.Y + ATable.FOffSet[0];
  re := Left.FBasePoint.Y + ATable.FOffSet[0] + Right.FBasePoint.Y + ATable.FRowHistory.Count + 1;
  cb := ATable.GetDesineColumn(AColumn.BaseColumn - 1) + 1;

  SrcRange := TXlsCellRange.Create(rb, cb, re, (cb + (AColumn.CountColumn)) - 1);

  with ABaseSheet.FlexCelTemplate.FlexCel do
  begin
    XF.InsertAndCopyRange(SrcRange,
    // параметры смещения
    rb, cb + ((AColumn.FIterariton + 1) * AColumn.CountColumn), 1, TFlxInsertMode.ShiftRangeRight,
      TRangeCopyMode.All, XF, SheetIndex + 1, nil);
  end;
end;

function TSheetTemplate.IncRow(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable): Integer;
var
  tmpTable: TTableSheet;
  TmpCelll: TCellSheet;
begin
  tmpTable := nil;
  for tmpTable in FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      tmpTable.FRowHistory.Add(ARow.IDRow);

      // здесь стоит вставить необходимую строку
      ABaseSheet.FlexCelTemplate.FlexCel.InsertAndCopyRow(tmpTable.GetDesineRow(tboBase,
        ARow.BaseRow), tmpTable.LastRow);

      // Увеличиваем счетчик текущей строки
      inc(tmpTable.FLasTRow);
      Result := tmpTable.FRowHistory.Count;
      ARow.FCorrenTRow := Result;

      Break;
    end
    else
      tmpTable.incOfSeTRow;
  end;

  if not Assigned(tmpTable) then
    raise EASDException.Create('TSheetTemplate.IncRow Не найдена необходимая вкладка')
  else
  begin
    // Kонтроль констант для изменения их текущего места положения
    for TmpCelll in Self.FListCell do
      if TmpCelll.Position.FBasePoint.Y >= tmpTable.Position.Y then
        TmpCelll.IncY;
    Result := tmpTable.FLasTRow;
  end;
end;

procedure TSheetTemplate.SetValue(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const ANumColumn: Integer; const AValue: Variant);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь вписываем значения
      ABaseSheet.FlexCelTemplate.FlexCel.SetValue(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineColumn(ANumColumn) - 1, AValue);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const ANumColumn: Integer; const AColor: TColor);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь ставим цвет ячейки
      ABaseSheet.FlexCelTemplate.FlexCel.SetColor(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineColumn(ANumColumn) - 1, AColor);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const AColor: TColor);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь ставим цвет строки
      ABaseSheet.FlexCelTemplate.FlexCel.SetColor(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineRow(tboCurrent, ARow.CurrenTRow),
        tmpTable.GetDesineTable(tcpLeftX), tmpTable.GetDesineTable(tcpRightX), AColor);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetBarCodeEAN13Image(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const ABarCode: Int64; const AForeColor: TColor; const aDy, aDx: Integer);
begin
  SetSheet(Self);
  // вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.InsertEAN13Image(ABarCode, AForeColor,
    ACell.Position.FBasePoint.Y, ACell.Position.FBasePoint.X, aDy, aDx);
end;

procedure TSheetTemplate.SetColor(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const AColor: TColor);
begin
  SetSheet(Self);
  // здесь вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.SetColor(ACell.Position.FBasePoint.Y,
    ACell.Position.FBasePoint.X, AColor);
end;

procedure TSheetTemplate.SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const AColumn: Integer; const AColor: TColor);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь ставим цвет шрифта
      ABaseSheet.FlexCelTemplate.FlexCel.SetColorFont(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineRow(tboCurrent, ARow.CurrenTRow),
        tmpTable.GetDesineColumn(AColumn) - 1, tmpTable.GetDesineColumn(AColumn) - 1, AColor);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
const AColumn: Integer; AFonts: TXLSFontData);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь ставим цвет шрифта
      ABaseSheet.FlexCelTemplate.FlexCel.SetFont(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineRow(tboCurrent, ARow.CurrenTRow),
        tmpTable.GetDesineColumn(AColumn), tmpTable.GetDesineColumn(AColumn), AFonts);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetColumnMerge(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const ABegColumn, AEndColumn: Integer);
var
  row, col1, col2: Integer;
begin
  SetSheet(Self);
  // здесь вписываем значения
  row := ATable.GetDesineRow(tboCurrent, ARow.CurrenTRow);
  col1 := ifthen(ABegColumn < 1, 1, ABegColumn);
  col2 := ifthen((AEndColumn < 1), ATable.GetColumnCount, AEndColumn);
  ABaseSheet.FlexCelTemplate.FlexCel.XF.MergeCells(row + 1, ATable.GetDesineColumn(col1), row + 1,
    ATable.GetDesineColumn(col2));
end;

procedure TSheetTemplate.SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const AColor: TColor);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь ставим цвет шрифта
      ABaseSheet.FlexCelTemplate.FlexCel.SetColorFont(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineRow(tboCurrent, ARow.CurrenTRow),
        tmpTable.GetDesineTable(tcpLeftX), tmpTable.GetDesineTable(tcpRightX), AColor);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetFormula(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
ARow: TRowTable; const AColumn: Integer; const AFormula: string);
var
  tmpTable: TTableSheet;
begin
  for tmpTable in Self.FTable do
  begin
    if tmpTable.NumTable = ATable.NumTable then
    begin
      SetSheet(Self);
      // здесь вписываем значения
      ABaseSheet.FlexCelTemplate.FlexCel.SetFormula(tmpTable.GetDesineRow(tboCurrent,
        ARow.CurrenTRow), tmpTable.GetDesineColumn(AColumn) - 1, AFormula);
      Break;
    end;
  end;
end;

procedure TSheetTemplate.SetFormula(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
const AFormula: string);
begin
  SetSheet(Self);
  // здесь вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.SetFormula(ARow, AColumn, AFormula);
end;

procedure TSheetTemplate.SetPointValue(const APoint: TPoint; const AValue: Variant);
begin
  with Self.FlexCelTemplate.FlexCel do
  begin
    SetSheet(Self);
    // здесь вписываем значения
    SetValue(APoint.Y, APoint.X, AValue);
  end;
end;

procedure TSheetTemplate.SetSheet(ASheet: TSheetTemplate);
begin
  if not(Self.FlexCelTemplate.FlexCel.SheetIndex = ASheet.SheetNum) then
  begin
    Self.FlexCelTemplate.FlexCel.SheetIndex := ASheet.SheetNum;
    Self.FlexCelTemplate.FlexCel.LoadVarListFromXLSFile;
  end;
end;

procedure TSheetTemplate.SetValue(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const AValue: Variant);
begin
  SetSheet(Self);
  // здесь вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.SetValue(ACell.Position.FBasePoint.Y,
    ACell.Position.FBasePoint.X, AValue);
end;

procedure TSheetTemplate.SetFormula(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const AFormula: String);
begin
  SetSheet(Self);
  // здесь вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.SetFormula(ACell.Position.FBasePoint.Y,
    ACell.Position.FBasePoint.X, AFormula);
end;

procedure TSheetTemplate.SetValue(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
const AValue: Variant);
begin
  SetSheet(Self);
  // здесь вписываем значения
  ABaseSheet.FlexCelTemplate.FlexCel.SetValue(ARow, AColumn, AValue);
end;

procedure TSheetTemplate.ReassignSheet(const ANum: Integer);
begin
  if ANum = 0 then
    raise EASDException.Create('TSheetTemplate.ReassignSheet Базовый шаблон не подлежит изменению');

  if not Assigned(Self.FSheetCurrent) then
    raise EASDException.Create
      ('TSheetTemplate.ReassignSheet Заполнение возможно только для добавленных листов');

  if (FListOwner.Count = 0) or (FListOwner.Count < ANum) then
    raise EASDException.Create('TSheetTemplate.ReassignSheet Обращение к несуществующему листу');

  FSheetCurrent := FOwner.FListSheet[FListOwner[ANum - 1]];
end;

procedure TSheetTemplate.SetValue(const AFBasePos: TBasePointCells; const AValue: Variant;
const ATypeSheet: TTypeSheet);
var
  sh: Integer;
begin
  if (ATypeSheet = ttsBase) then
    sh := FBaseSheet
  else
    sh := FSheetNum;

  with TFlexCelTemplate(FOwner).FlexCel do
  begin
    if not(SheetIndex = sh) then
    begin
      SheetIndex := sh;
      LoadVarListFromXLSFile;
    end;
    // устанавливаем значение по указателю
    SetValue(AFBasePos.FBasePoint.Y, AFBasePos.FBasePoint.X, AValue);
  end;
end;
{$ENDREGION}
{$REGION ' TTableSheet ....}

procedure TTableSheet.AddBaseExtColumn(ABaseColumn: TExtColumns);
var
  TmpCol: TExtColumns;
begin
  TmpCol := TExtColumns.Create(Self, ABaseColumn);
  Self.FExtCol.Add(TmpCol);
end;

function TTableSheet.AddBaseExtColumn(const ABegColumn: Integer; const ACountColumn: Integer = 1)
  : TExtColumns;
var
  index: Integer;
  TmpCol: TExtColumns;
begin
  TmpCol := TExtColumns.Create(Self, ABegColumn, ACountColumn);
  if not Self.FExtCol.BinarySearch(TmpCol, index) then
    Self.FExtCol.Insert(index, TmpCol)
  else
    raise EASDException.Create
      ('TTableSheet.AddBaseExtColumn Попытка добавить вторую колонку с базовым индексом');

  Result := TmpCol;
end;

function TTableSheet.AddBaseRow(const ANumRow: Integer): TRowTable;
begin
  if not(Self.FOwner.SheetNum = 0) then
    raise EASDException.Create
      ('TTableSheet.AddBaseRow Добавление строк таблиц возможно до первого заполнения отчета');

  Result := TRowTable.Create(Self, ANumRow, Self.FRow.Count + 1);
  Self.FRow.Add(Result);
  Self.FRow.Sort;
  Result.Name := 'BaseRow' + IntToStr(Self.FRow.Count);
end;

procedure TTableSheet.BeforeDestruction;
begin
  if (Owner = nil) or (csDestroying in Owner.ComponentState) then
    inherited
  else
    raise EASDException.Create
      ('TTableSheet.BeforeDestruction Объекты данного класса удаляются только родителями');
end;

procedure TTableSheet.AddBaseRow(ABaseRow: TRowTable);
var
  lt: TRowTable;
begin
  lt := TRowTable.Create(Self, ABaseRow.BaseRow, ABaseRow.IDRow);
  Self.FRow.Add(lt);
  lt.Name := 'CurrenRow' + IntToStr(Self.FRow.Count);
end;

procedure TTableSheet.CleaningTable;
var
  tmpRow: TRowTable;
  TmpCol: TExtColumns;
begin
  if TSheetTemplate(Self.Owner).SheetType = ttsBase then
    Exit;

  for TmpCol in FExtCol do
    FOwner.DeleteBaseColumns(Self, TmpCol);

  for tmpRow in FRow do
    if (Self.FRowHistory.Count > 0) or (tmpRow.IDRow > 1) then
    begin
      FOwner.DeleteBaseRow(Self, tmpRow);
    end;
end;

procedure TTableSheet.CorretionBasePos(const AView: TCorrectionPos; const AValue: Integer);
var
  tmp: TRowTable;
begin
  with Self.FRectangle do
  begin
    case AView of
      tcpLeftY:
        begin
          Self.FOffsetBase[0] := Self.FOffsetBase[0] + AValue;
          Self.FOffsetBase[2] := Self.FOffsetBase[2] - AValue;
          // При коректировки верхнего угла таблицы по горизонтали корректируем отсчет строк
          for tmp in FRow do
            tmp.CurretBasePos(AValue);
          FRow.Sort;
        end;
      tcpLeftX:
        Self.FOffsetBase[1] := AValue;
      tcpRightY:
        Self.FOffsetBase[2] := AValue;
      tcpRightX:
        Self.FOffsetBase[3] := AValue;
    end;
  end;
end;

constructor TTableSheet.Create(AOwner: TSheetTemplate; const ALeftPos, ARigthPos: TBasePointCells;
const ANumTable: Integer = -1);
begin
  if ANumTable < 0 then
    raise EASDException.Create('TTableSheet.Create Необходимо указать номер таблицы на листе');

  inherited Create(AOwner);
  FOffSetTableName := 0;
  FOwner := AOwner;

  FDropDef := False;
  FDropTable := False;

  ALeftPos.Assign(FRectangle.FLeft);
  ARigthPos.Assign(FRectangle.FRight);

  FLasTRow := 0;
  FOffsetBase[0] := 0;
  FOffsetBase[1] := 0;
  FOffsetBase[2] := 0;
  FOffsetBase[3] := 0;

  FOffSet[0] := 0;
  FOffSet[1] := 0;
  FOffSet[2] := 0;
  FOffSet[3] := 0;

  Self.FNumTable := ANumTable;

  // Лист истории для каждой строки в таблице необходимы для сумарных значений и расчета текущего мп
  FRowHistory := TList<Integer>.Create;

  // Создаем список строк таблицы
  FRow := TObjectList<TRowTable>.Create(TComparer<TRowTable>.Construct(
    function(const Left, Right: TRowTable): Integer
    begin
      Result := CompareValue(Right.BaseRow, Left.BaseRow);
    end), True);

  // Множитель  колонок
  FExtCol := TObjectList<TExtColumns>.Create(TComparer<TExtColumns>.Construct(
    function(const Left, Right: TExtColumns): Integer
    begin
      Result := CompareValue(Right.BaseColumn, Left.BaseColumn);
    end), True);
end;

destructor TTableSheet.Destroy;
begin
  FRow.Free;
  FRowHistory.Free;
  FExtCol.Free;
  inherited;
end;

procedure TTableSheet.DropDef(const OffSetTName: Integer);
begin
  Self.FOffSetTableName := OffSetTName;
  Self.FDropDef := True;
end;

procedure TTableSheet.DropTable(const OffSetTName: Integer = 0);
var
  tmpTable: TTableSheet;
begin
  tmpTable := Self.FOwner.Sheet.GetTable(Self.NumTable);
  tmpTable.FOffSetTableName := OffSetTName;
  tmpTable.FDropTable := True;
end;

function TTableSheet.ExtColumnCount: Integer;
begin
  Result := FExtCol.Count;
end;

function TTableSheet.Position: TPoint;
begin
  Result.Y := Self.FRectangle.FLeft.FBasePoint.Y + Self.FOffsetBase[0] + Self.FOffSet[0];
  Result.X := Self.FRectangle.FLeft.FBasePoint.X + Self.FOffsetBase[1] + Self.FOffSet[1];
end;

function TTableSheet.FPositionRigth: TPoint;
begin
  Result.Y := Self.Position.Y + Self.FRectangle.FRight.FBasePoint.Y + Self.FOffsetBase[2] +
    Self.FOffSet[2];
  Result.X := Self.Position.X + Self.FRectangle.FRight.FBasePoint.X + Self.FOffsetBase[3] +
    Self.FOffSet[3];
end;

function TTableSheet.GetBasePosition(const ATypePoint: TTypePoint): TBasePointCells;
begin
  with FRectangle do
  begin
    Result.FBasePoint.X := ifthen(ATypePoint = ttpLeft, FLeft.FBasePoint.X, FRight.FBasePoint.X);
    Result.FBasePoint.Y := ifthen(ATypePoint = ttpLeft, FLeft.FBasePoint.Y, FRight.FBasePoint.Y);
    Result.FConst := ifthen(ATypePoint = ttpLeft, FLeft.FConst, FRight.FConst);

    if ATypePoint = ttpLeft then
    begin
      Result.FValue := FLeft.FValue;
      Result.FBasePoint.Y := Result.FBasePoint.Y + FOffsetBase[0];
      Result.FBasePoint.X := Result.FBasePoint.X + FOffsetBase[1];
    end
    else
    begin
      Result.FValue := FRight.FValue;
      Result.FBasePoint.Y := Result.FBasePoint.Y + FOffsetBase[2];
      Result.FBasePoint.X := Result.FBasePoint.X + FOffsetBase[3];
    end;
  end;
end;

function TTableSheet.GetCell(const ATypePos: TTypePoint): TPoint;
var
  tmpCells: TBasePointCells;
  tmpTable: TTableSheet;
  Sheet: TSheetTemplate;
begin
  Sheet := GetCurrentSheet;
  tmpTable := GetCurrentTable(Sheet);
  tmpCells := tmpTable.GetBasePosition(ATypePos);
  Result := tmpCells.FBasePoint;
end;

function TTableSheet.GetColumnCount: Integer;
begin
  Result := Self.GetDesineTable(tcpRightX) - Self.GetDesineTable(tcpLeftX) + 1;
end;

function TTableSheet.GetColumns(ABaseColun: TExtColumns): TExtColumns;
var
  tmp: TExtColumns;
begin
  tmp := nil;
  for tmp in FExtCol do
    if (tmp.IDColumn = ABaseColun.IDColumn) then
      Break;

  if not Assigned(tmp) or not(tmp.IDColumn = ABaseColun.IDColumn) then
    raise EASDException.Create('TTableSheet.GeTRow В целевой таблице не найдена строка');

  Result := tmp;
end;

function TTableSheet.GetDesineColumn(const ACoumn: Integer): Integer;
var
  TmpCol: TExtColumns;
  IncCol: Integer;
begin
  IncCol := 0;
  for TmpCol in FExtCol do
  begin
    if ((ACoumn >= TmpCol.BaseColumn) and (ACoumn <= (TmpCol.BaseColumn + TmpCol.CountColumn) - 1))
      and ((TmpCol.Iterariton < 1) or (TmpCol.CurrentIter < 1)) then
      raise EASDException.Create
        ('TTableSheet.GetDesineColumn Базовый элемент не подлежит изменению');

    if (ACoumn < TmpCol.BaseColumn) or (TmpCol.CurrentIter = 0) then
      Continue;

    if ACoumn = TmpCol.BaseColumn then
    begin
      IncCol := IncCol + (TmpCol.CurrentIter * TmpCol.CountColumn);
      Continue;
    end;

    if ACoumn > TmpCol.BaseColumn then
    begin
      IncCol := IncCol + (TmpCol.Iterariton * TmpCol.CountColumn);
      Continue;
    end;
  end;

  Result := Self.Position.X + IncCol + ACoumn;
end;

function TTableSheet.GetDesineRow(const ATypeRow: tBaseObj; const ABaseRow: Integer): Integer;
begin
  if FRow.Count = 0 then
    raise EASDException.Create('TTableSheet.GetDisineRow Не найден список строк таблицы');

  if (ATypeRow = tboBase) then
    Result := Self.Position.Y + ABaseRow
  else
    Result := Self.Position.Y + FRow[0].BaseRow + ABaseRow;
end;

function TTableSheet.GetDesineTable(const ATypePos: TCorrectionPos): Integer;
begin
  case ATypePos of
    // tcpLeftY:
    tcpLeftX:
      Result := GetDesineColumn(0);
    // tcpRightY:
    tcpRightX:
      Result := Self.FPositionRigth.X;
  else
    Result := -1;
  end;
end;

function TTableSheet.GetFormula(const AColumn: Integer; AListIterations: TList<Integer>): string;
var
  iteration: Integer;
  Incessant: Boolean;
  Position: TPoint;
begin
  // Проверяем является ли лист последовательным
  Incessant := (AListIterations[AListIterations.Count - 1] - AListIterations.Count)
    = AListIterations[0] - 1;

  Position.X := Self.GetDesineColumn(AColumn);
  Result := '=SUM(';

  if Incessant then
  begin
    Position.Y := Self.GetDesineRow(tboBase, Self.FRow[0].BaseRow) + AListIterations[0] + 1;
    Result := Result + TFlexCelTemplate.PointToAddr(Position);
    Result := Result + ': ';
    Position.Y := Self.GetDesineRow(tboBase, Self.FRow[0].BaseRow) + AListIterations
      [AListIterations.Count - 1] + 1;
    Result := Result + TFlexCelTemplate.PointToAddr(Position);
  end
  else
  begin
    Position.Y := Self.GetDesineRow(tboBase, Self.FRow[0].BaseRow) + AListIterations[0] + 1;
    Result := Result + TFlexCelTemplate.PointToAddr(Position);
    for iteration := 1 to AListIterations.Count - 1 do
    begin
      Position.Y := Self.GetDesineRow(tboBase, Self.FRow[0].BaseRow) + AListIterations
        [iteration] + 1;
      Result := Result + ';' + TFlexCelTemplate.PointToAddr(Position);
    end;
  end;

  Result := Result + ')';
end;

function TTableSheet.GeTRow(const AIDRow: Integer): TRowTable;
begin
  Result := nil;
  for Result in FRow do
    if (Result.IDRow = AIDRow) then
      Break;

  if not Assigned(Result) or not(Result.IDRow = AIDRow) then
    raise EASDException.Create('TTableSheet.GeTRow в целевой таблице ненайдена строка');
end;

function TTableSheet.GetCell(const ANumRow: Integer; const ANumColumn: Integer = -1): TPoint;
var
  tmpTable: TTableSheet;
  Sheet: TSheetTemplate;
  nc: Integer;
begin
  Sheet := GetCurrentSheet;
  tmpTable := GetCurrentTable(Sheet);

  if ANumRow > tmpTable.FRowHistory.Count then
    raise EASDException.Create
      ('TTableSheet.GetPoinTRow Обращение к несуществующей строки таблицы');

  nc := ifthen(ANumColumn < 1, 1, ANumColumn);
  Result.Y := tmpTable.Position.Y + tmpTable.FRow[0].BaseRow + ANumRow + 1;
  Result.X := tmpTable.GetDesineColumn(nc) - 1;

  // Переключаем FlexCel на нужную вкладку
  with Sheet.FlexCelTemplate.FlexCel do
  begin
    if not(SheetIndex = Sheet.SheetNum) then
    begin
      SheetIndex := Sheet.SheetNum;
      LoadVarListFromXLSFile;
    end;
  end;
end;

function TTableSheet.GetCurrenTRow(ARow: TRowTable; const AIteration: Integer): Integer;
begin
  if (ARow.LisTRow.Count < AIteration) or (ARow.LisTRow.Count < 1) or
    (ARow.LisTRow.Count < AIteration) then
    raise EASDException.Create('TTableSheet.GetRealRow Запрос на несуществующую строка');

  if (AIteration < 1) then
    Result := ARow.LisTRow[ARow.LisTRow.Count - 1]
  else
    Result := ARow.LisTRow[AIteration - 1];
end;

function TTableSheet.GetCurrentSheet: TSheetTemplate;
begin
  if not Assigned(SheetTemplate.Sheet) then
    raise EASDException.Create('TTableSheet.SetSumTable Обращение к несуществующей вкладке');

  Result := SheetTemplate.Sheet;
end;

function TTableSheet.GetCurrentTable(ASheet: TSheetTemplate): TTableSheet;
var
  tmpTable: TTableSheet;
begin
  tmpTable := ASheet.GetTable(Self.FNumTable);
  if not Assigned(tmpTable) then
    raise EASDException.Create('TTableSheet.SetSumTable Обращение к несуществующей таблице');

  Result := tmpTable;
end;

procedure TTableSheet.SetSumTable(ARow: TRowTable; const AColumn: Integer;
const ADefault: Variant; const RowResult: Integer = 1);
var
  tmpTable: TTableSheet;
  Formula: string;
  valid: Boolean;
  iter: Integer;
  Sheet: TSheetTemplate;
  LisTRow: TList<Integer>;
begin
  Sheet := GetCurrentSheet;
  tmpTable := GetCurrentTable(Sheet);

  SetLength(Formula, 0);

  if Assigned(ARow) and (ARow.RowCount > 0) then
  begin
    tmpTable.GeTRow(ARow.IDRow);
    Formula := ARow.GetFormula(AColumn - 1, nil, valid);
  end
  else
    if not Assigned(ARow) and (tmpTable.FRowHistory.Count > 0) then
    begin
      valid := True;
      LisTRow := TList<Integer>.Create;
      try
        for iter := 0 to tmpTable.FRowHistory.Count - 1 do
          LisTRow.Add(iter + 1);

        Formula := tmpTable.GetFormula(AColumn - 1, LisTRow);
      finally
        LisTRow.Free;
      end;
    end
    else
      valid := False;

  // Если нет ни одного значения
  if not(valid) then
    Sheet.SetValue(SheetTemplate, tmpTable.LastRow + RowResult, tmpTable.GetDesineColumn(AColumn) -
      1, ADefault)
  else
    Sheet.SetFormula(SheetTemplate, tmpTable.LastRow + RowResult, tmpTable.GetDesineColumn(AColumn)
      - 1, Formula);
end;

procedure TTableSheet.IncColumn(ABaseSheet: TSheetTemplate; AColumn: TExtColumns);
begin
  Self.FOwner.IncColumn(ABaseSheet, Self, AColumn);
end;

function TTableSheet.IncRow(ABaseSheet: TSheetTemplate; ARow: TRowTable): Integer;
begin
  Result := Self.FOwner.IncRow(ABaseSheet, Self, ARow);
end;

procedure TTableSheet.SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const ANumColumn: Integer; const AColor: TColor);
begin
  Self.FOwner.SetColor(ABaseSheet, Self, ARow, ANumColumn, AColor);
end;

procedure TTableSheet.SetGroupRows(ARow: TRowTable; const AEndNumRow: Integer);
var
  Sheet: TSheetTemplate;
  lTop, llower: Integer;
begin
  Sheet := Self.GetCurrentSheet;
  Self.GetCurrentTable(Sheet);
  lTop := Self.GetCell(ARow.GetNumRow).Y;
  llower := Self.GetCell(AEndNumRow).Y;
  SetCurrentSheet(Sheet);
  if lTop < llower then
  begin
    Sheet.FlexCelTemplate.FlexCel.XF.OutlineSummaryRowsBelowDetail := False;
    Sheet.FlexCelTemplate.FlexCel.XF.SetRowOutlineLevel(lTop + 1, llower, 1);
  end;
end;

procedure TTableSheet.SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColor: TColor);
begin
  Self.FOwner.SetColor(ABaseSheet, Self, ARow, AColor);
end;

procedure TTableSheet.SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const AColor: TColor);
begin
  Self.FOwner.SetColorFont(ABaseSheet, Self, ARow, AColor);
end;

procedure TTableSheet.SetColumnMerge(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const ABegColumn, AEndColumn: Integer);
begin
  Self.FOwner.SetColumnMerge(ABaseSheet, Self, ARow, ABegColumn, AEndColumn);
end;

procedure TTableSheet.SetCurrentSheet(ASheet: TSheetTemplate);
begin
  with ASheet.FlexCelTemplate.FlexCel do
  begin
    if not(SheetIndex = ASheet.SheetNum) then
    begin
      SheetIndex := ASheet.SheetNum;
      LoadVarListFromXLSFile;
    end;
  end;
end;

procedure TTableSheet.SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const AColumn: Integer; const AColor: TColor);
begin
  Self.FOwner.SetColorFont(ABaseSheet, Self, ARow, AColumn, AColor);
end;

procedure TTableSheet.SetFont(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
AFonts: TXLSFontData);
begin
  Self.FOwner.SetFont(ABaseSheet, Self, ARow, AColumn, AFonts);
end;

procedure TTableSheet.SetFormula(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const AColumn: Integer; const AFormula: string);
begin
  Self.FOwner.SetFormula(ABaseSheet, Self, ARow, AColumn, AFormula);
end;

procedure TTableSheet.SetMergeGroupRow(const ABegNumRow, AEndNumRow: Integer;
const ABeginCol, AEndCol: Integer);
var
  tmpTable: TTableSheet;
  Sheet: TSheetTemplate;
  PBeg, PEnd: TPoint;
begin
  Sheet := GetCurrentSheet;
  tmpTable := Self.GetCurrentTable(Sheet);

  SetCurrentSheet(Sheet);

  PBeg := Self.GetCell(ABegNumRow, ABeginCol);
  PEnd := Self.GetCell(AEndNumRow, AEndCol);
  FOwner.FlexCelTemplate.FlexCel.XF.MergeCells(PBeg.Y, tmpTable.GetDesineColumn(PBeg.X) - 1, PEnd.Y,
    tmpTable.GetDesineColumn(PEnd.X) - 2);
end;

procedure TTableSheet.SetTitleColumns(const AColumn: Integer; const AValue: Variant;
const ARowTitle: Integer = 0);
var
  tmpTable: TTableSheet;
  Sheet: TSheetTemplate;
begin
  Sheet := GetCurrentSheet;
  tmpTable := Self.GetCurrentTable(Sheet);

  Sheet.SetValue(SheetTemplate, (tmpTable.Position.Y + ARowTitle),
    tmpTable.GetDesineColumn(AColumn) - 1, AValue);
end;

procedure TTableSheet.SetValue(ABaseSheet: TSheetTemplate; ARow: TRowTable;
const ANumColumn: Integer; const AValue: Variant);
begin
  Self.FOwner.SetValue(ABaseSheet, Self, ARow, ANumColumn, AValue);
end;

procedure TTableSheet.incOfSeTRow;
begin
  inc(FOffSet[0]);
end;

function TTableSheet.LastRow: Integer;
var
  RowTable: Integer;
begin
  if FRow.Count = 0 then
    raise EASDException.Create('TTableSheet.GetDisineRow Не найден список строк таблицы');

  RowTable := Self.Position.Y;
  Result := RowTable + FRow[0].BaseRow + FRowHistory.Count;
end;

procedure TTableSheet.OverloadRowCol(ARecipient: TTableSheet);
var
  tmpRow: TRowTable;
  TmpCol: TExtColumns;
begin
  for tmpRow in Self.FRow do
    ARecipient.AddBaseRow(tmpRow);

  for TmpCol in Self.FExtCol do
    ARecipient.AddBaseExtColumn(TmpCol);

  ARecipient.FRow.Sort;
  ARecipient.FExtCol.Sort;
end;
{$ENDREGION}
{$REGION ' TCellSheet ....}

constructor TCellSheet.Create(AOwner: TComponent; ABaseSheet: TSheetTemplate;
const ABasePos: TBasePointCells; const AIDCell: Integer);
begin
  if not(AOwner.InheritsFrom(TSheetTemplate)) then
    raise EASDException.Create('TFlexCell.Create Ошибка создания указателя на базовую ячейку');

  inherited Create(AOwner);
  FOwner := ABaseSheet;
  ABasePos.Assign(Self.FBasePos);
  FIDCell := AIDCell;
  FIncY := 0;
end;

procedure TCellSheet.BeforeDestruction;
begin
  if (Owner = nil) or (csDestroying in Owner.ComponentState) then
    inherited
  else
    raise EASDException.Create
      ('TCellSheet.BeforeDestruction Объекты данного класса удаляются только родителями');
end;

constructor TCellSheet.Create(AOwner: TComponent; ABaseSheet: TSheetTemplate;
const ABasePos: TCellSheet; const AIDCell: Integer);
begin
  Create(AOwner, ABaseSheet, ABasePos.FBasePos, AIDCell);
end;

function TCellSheet.GetCell: TPoint;
begin
  GetCurrentElement;

  if Not Assigned(FSheet) then
    raise EASDException.Create('TTableSheet.GetPoinTRow Обращение к несуществующей вкладкe');

  Result.Y := FCell.Position.FBasePoint.Y + 1;
  Result.X := FCell.Position.FBasePoint.X;

  // Переключаем FlexCel на нужную вкладку
  with FSheet.FlexCelTemplate.FlexCel do
  begin
    if not(SheetIndex = FSheet.SheetNum) then
    begin
      SheetIndex := FSheet.SheetNum;
      LoadVarListFromXLSFile;
    end;
  end;
end;

procedure TCellSheet.GetCurrentElement;
begin
  FSheet := FOwner.Sheet;
  FCell := FSheet.GetCell(Self.IDCell);
end;

function TCellSheet.GetPosition: TBasePointCells;
begin
  FBasePos.Assign(Result);
  Result.FBasePoint.Y := Result.FBasePoint.Y + FIncY;
end;

procedure TCellSheet.IncY(const AStep: Integer);
begin
  Self.FIncY := Self.FIncY + AStep;
end;

procedure TCellSheet.SetBarCodeEAN13Image(ABarCode: Int64; AForeColor: TColor; aDy, aDx: Integer);
begin
  GetCurrentElement;
  FSheet.SetBarCodeEAN13Image(TSheetTemplate(Self.Owner), FCell, ABarCode, AForeColor, aDy, aDx);
end;

procedure TCellSheet.SetColor(const AColor: TColor);
begin
  GetCurrentElement;
  FSheet.SetColor(TSheetTemplate(Self.Owner), FCell, AColor);
end;

procedure TCellSheet.SetFormula(AInRow: TRowTable; const AInColumn: Integer;
const ADefault: Variant);
var
  Formula: String;
  valid: Boolean;
begin
  Formula := AInRow.GetFormula(AInColumn - 1, nil, valid);
  GetCurrentElement;
  if valid then
    FSheet.SetFormula(TSheetTemplate(Self.Owner), FCell, Formula)
  else
    FSheet.SetValue(TSheetTemplate(Self.Owner), FCell, ADefault);
end;

procedure TCellSheet.SetValue(const AValue: Variant);
begin
  GetCurrentElement;
  FSheet.SetValue(TSheetTemplate(Self.Owner), FCell, AValue);
end;
{$ENDREGION}
{$REGION ' TRowTable ....}

procedure TRowTable.BeforeDestruction;
begin
  if (Owner = nil) or (csDestroying in Owner.ComponentState) then
    inherited
  else
    raise EASDException.Create
      ('TRowTable.BeforeDestruction Объекты данного класса удаляются только родителями');
end;

constructor TRowTable.Create(AOwner: TTableSheet; const ABaseRow: Integer;
const AIDRow: Integer);
begin
  if not(AOwner.InheritsFrom(TTableSheet)) then
    raise EASDException.Create('TRowTable.Create Ошибка создания указателя на базовую строку');

  inherited Create(AOwner);
  FOwner := AOwner;
  FLisTRow := TList<Integer>.Create(TComparer<Integer>.Construct(
    function(const Left, Right: Integer): Integer
    begin
      Result := CompareValue(Right, Left);
    end));
  FBaseRow := ABaseRow;
  FIDRow := AIDRow;
  FCorrenTRow := 0;
end;

procedure TRowTable.CurretBasePos(const AOfSet: Integer);
begin
  Self.FBaseRow := Self.FBaseRow - AOfSet;
end;

constructor TRowTable.Create(AOwner: TTableSheet; const ABaseRow: TRowTable);
begin
  Create(AOwner, ABaseRow.FBaseRow, ABaseRow.IDRow);
end;

function TRowTable.CurrenTRow: Integer;
begin
  if Self.FLisTRow.Count = 0 then
    raise EASDException.Create('TRowTable.GetCurrenTRow Поиск несуществующей строки');

  Result := Self.FLisTRow[Self.FLisTRow.Count - 1];
end;

destructor TRowTable.Destroy;
begin
  FLisTRow.Free;
  inherited;
end;

procedure TRowTable.GetCurrentElement;
begin
  FSheet := FOwner.SheetTemplate.Sheet;
  if not Assigned(FSheet) then
    raise EASDException.Create('TRowTable.GetCurrentElement Обращение к несуществующему листу');

  FTable := FOwner.SheetTemplate.Sheet.GetTable(Self.FOwner.NumTable);
  FRow := FTable.GeTRow(Self.IDRow);
end;

function TRowTable.GetFormula(const AColumn: Integer; AListIterations: TIntegerList;
out AValid: Boolean): string;
var
  Rowlist: TList<Integer>;
  iter, idx, num: Integer;
begin
  GetCurrentElement;
  AValid := False;

  Rowlist := TList<Integer>.Create(TComparer<Integer>.Construct(
    function(const Left, Right: Integer): Integer
    begin
      Result := CompareValue(Left, Right);
    end));

  try
    if Assigned(AListIterations) then
    begin
      Rowlist.Capacity := AListIterations.Count;

      for iter := 0 to AListIterations.Count - 1 do
      begin
        num := FRow.LisTRow[AListIterations[iter] - 1];
        if not Rowlist.BinarySearch(num, idx) then
          Rowlist.Insert(idx, num);
      end;
    end
    else
      for iter in FRow.FLisTRow do
      begin
        if not Rowlist.BinarySearch(iter, idx) then
          Rowlist.Insert(idx, iter);
      end;
    Rowlist.Sort;

    AValid := Rowlist.Count > 0;

    if AValid then
      Result := FTable.GetFormula(AColumn, Rowlist);
  finally
    FreeAndNil(Rowlist);
  end;
end;

function TRowTable.GetNumRow(const AIteration: Integer = -1): Integer;
begin
  GetCurrentElement;
  Result := FTable.GetCurrenTRow(FRow, AIteration);
end;

function TRowTable.InsertRow: Integer;
var
  index, num: Integer;
begin
  GetCurrentElement;

  index := FTable.IncRow(FSheet, FRow);
  if not FLisTRow.BinarySearch(index, num) then
    FRow.FLisTRow.Add(index)
  else
    raise EASDException.Create('TRowTable.InserTRow Ошибка добавления строки');

  Result := FRow.FLisTRow.Count;
end;

function TRowTable.IterationCount: Integer;
begin
  Result := Self.RowCount;
end;

function TRowTable.RowCount: Integer;
begin
  GetCurrentElement;
  Result := FRow.FLisTRow.Count;
end;

function TRowTable.GetIteration: Integer;
begin
  GetCurrentElement;
  Result := FRow.FLisTRow.Count;
end;

procedure TRowTable.SetColor(const ANumColumn: Integer; const AColor: TColor);
begin
  GetCurrentElement;
  FTable.SetColor(FSheet, FRow, ANumColumn, AColor);
end;

procedure TRowTable.SetColor(const AColor: TColor);
begin
  GetCurrentElement;
  FTable.SetColor(FSheet, FRow, AColor);
end;

procedure TRowTable.SetFormula(AInRow: TRowTable; const AInColumn: Integer;
const ADefault: Variant; const AResColumn: Integer = -1);
var
  Formula: string;
  rlsColumn: Integer;
  valid: Boolean;
begin
  Formula := AInRow.GetFormula(AInColumn - 1, nil, valid);
  GetCurrentElement;
  rlsColumn := ifthen(AResColumn < 0, AInColumn, AResColumn);

  if valid then
    FTable.SetFormula(FSheet, FRow, rlsColumn, Formula)
  else
    FTable.SetValue(FSheet, FRow, rlsColumn, ADefault);
end;

procedure TRowTable.SetFont(const AFonts: TXLSFontData; const AColumn: Integer);
begin
  GetCurrentElement;
  FTable.SetFont(FSheet, FRow, AColumn, AFonts);
end;

procedure TRowTable.SetColorFont(const AColor: TColor; const AColumn: Integer);
begin
  GetCurrentElement;
  FTable.SetColorFont(FSheet, FRow, AColumn, AColor);
end;

procedure TRowTable.SetColorFont(const AColor: TColor);
begin
  GetCurrentElement;
  FTable.SetColorFont(FSheet, FRow, AColor);
end;

procedure TRowTable.SetFormula(AInRow: TRowTable; const AInColumn: Integer;
AListIterations: TIntegerList; const ADefault: Variant; const AResColumn: Integer = -1);
var
  Formula: string;
  rlsColumn: Integer;
  valid: Boolean;
begin
  Formula := AInRow.GetFormula(AInColumn - 1, AListIterations, valid);
  GetCurrentElement;
  rlsColumn := ifthen(AResColumn < 0, AInColumn, AResColumn);

  if valid then
    FTable.SetFormula(FSheet, FRow, rlsColumn, Formula)
  else
    FTable.SetValue(FSheet, FRow, rlsColumn, ADefault);
end;

procedure TRowTable.SetColumnMerge(const ABegColumn: Integer = -1; const AEndColumn: Integer = -1);
begin
  GetCurrentElement;
  FTable.SetColumnMerge(FSheet, FRow, ABegColumn, AEndColumn);
end;

procedure TRowTable.SetValueAndMerge(const ABegColumn, AEndColumn: Integer; const AValue: Variant);
begin
  Self.SetColumnMerge(ABegColumn, AEndColumn);
  FTable.SetValue(FSheet, FRow, ifthen(ABegColumn < 1, 1, ABegColumn), AValue);
end;

procedure TRowTable.SetValue(const ANumColumn: Integer; const AValue: Variant);
begin
  GetCurrentElement;
  FTable.SetValue(FSheet, FRow, ANumColumn, AValue);
end;
{$ENDREGION}
{$REGION ' TExtColumns ...'}

constructor TExtColumns.Create(AOwner: TTableSheet; const AOneColumn, ACountColumn: Integer);
begin
  inherited Create(AOwner);
  Self.FOwner := AOwner;
  Self.FIDColumn := FOwner.ExtColumnCount + 1;
  Self.FBaseColumn := AOneColumn;
  Self.FCountColumn := ACountColumn;
  Self.FIterariton := 0;
  Self.Name := 'BaseColumn_' + IntToStr(FIDColumn);
end;

procedure TExtColumns.BeforeDestruction;
begin
  if (Owner = nil) or (csDestroying in Owner.ComponentState) then
    inherited
  else
    raise EASDException.Create
      ('TExtColumns.BeforeDestruction Объекты данного класса удаляются только родителями');
end;

constructor TExtColumns.Create(AOwner: TTableSheet; ABaseColumn: TExtColumns);
begin
  inherited Create(AOwner);
  Self.FOwner := AOwner;
  Self.FIDColumn := ABaseColumn.FIDColumn;
  Self.FBaseColumn := ABaseColumn.BaseColumn;
  Self.FCountColumn := ABaseColumn.CountColumn;
  Self.FIterariton := 0;
  Self.Name := 'CurrentColumn_' + IntToStr(Self.FIDColumn);
end;

procedure TExtColumns.GetCurrentColumn;
begin
  FSheet := FOwner.SheetTemplate.Sheet;
  if not Assigned(FSheet) then
    raise EASDException.Create('TRowTable.GetCurrentElement Обращение к несуществуюшему листу');

  FTable := FOwner.SheetTemplate.Sheet.GetTable(Self.FOwner.NumTable);
  FColumn := FTable.GetColumns(Self);
end;

function TExtColumns.GetIteration: Integer;
begin
  GetCurrentColumn;
  Result := FColumn.CurrentIter;
end;

function TExtColumns.InsertColumn: Integer;
begin
  GetCurrentColumn;
  FTable.IncColumn(FSheet, FColumn);
  inc(FColumn.FIterariton);
  FColumn.FCurrentIter := FColumn.FIterariton;
  Result := FColumn.FCurrentIter;
end;

function TExtColumns.IteraritonCount: Integer;
begin
  GetCurrentColumn;
  Result := FColumn.FIterariton;
end;

procedure TExtColumns.SetCurrentIter(const AIteration: Integer);
begin
  if AIteration < 1 then
    raise EASDException.Create('TExtColumns.SetCurrentIter Обращение к базовой колонке');

  GetCurrentColumn;

  if FColumn.FIterariton < AIteration then
    raise EASDException.Create('TExtColumns.SetCurrentIter Обращение к несуществующей колонке');

  FColumn.FCurrentIter := AIteration;
end;
{$ENDREGION}

end.
