(*
  07.02.2018
  ��������� �.�.
  ������� �������� ������������ ������� �� ��������

  �����������:
  1. ���������� ����� � ����� ������ , ��������� ����� �� ����� � ����� ������� �������,
  ��� ���������� � ������� ����������.

  2. ��������/�������� ��������� ��������� � �������������� ������, ��� �������� ����������.

  3. ������ � ������������� ��������� ��� ������ �������, � ����������� �������������� �������������,
  ��������� � ������������ ��������.

  4. ����������� ������ ������ ������������ ��������� �����.

  5. ������ � ����� ��� ����� ������� ��������.

  6. ����������� ��������� � �������� � ����������� ������� �������,
  ��� ������� ����������������� ���������� ������.

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

{$REGION ' ���������� ������ � ��������� ... '}
(*
  1. ������� ������������ ����������� �������� ������ ��������� � ���������� ��� ��������  ������������� �������.
  � ���� ������� ������ ��� �������� � ��������� ������� ����������, ��������� ����� ��������� ������������ �������� ���������.
  ��� ��������� �������� (������, �������, ��������) ����� ��������� ������������� ������� �������� � �� ������� �������������� ��������.
  ��� ��������� �������� (������, �������, ��������) ������ ���� ��������� ���� �� ���� ��� �� ������ ����� ��� ���������� (��������� �������� �� �������� �������������)
  ������ ��������� ������� ����� ���� ����� �������� �� ������ ����������� ����� � ������ ���� ����� ���� ����� ���������.
  2. �� ������������������ ������� ����� ������������� ������� ��� ������ ������
*)
{$ENDREGION}

type

  TSheetTemplate = class;
  TRowTable = class;
  TTableSheet = class;

{$REGION ' ������� ���� ��� ����������� ������������� ....'}
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

  // ������������/����������� ������������ ����� ������� ������� �� ������� ������� ������� � �����
  TExtColumns = class(TComponent)
  strict private
    FOwner: TTableSheet;
    FIDColumn: Integer;
    FBaseColumn: Integer;
    FCountColumn: Integer;
    FCurrentIter: Integer;

    // ���������� ���������� � ������� �������
    FTable: TTableSheet;
    FColumn: TExtColumns;
    FSheet: TSheetTemplate;
  private
    FIterariton: Integer;

    // ����������� ��� �������� ��������
    constructor Create(AOwner: TTableSheet; const AOneColumn: Integer;
      const ACountColumn: Integer = 1); reintroduce; overload;
    // ����������� ��� ��������-���������� � �������� ��������
    constructor Create(AOwner: TTableSheet; ABaseColumn: TExtColumns); reintroduce; overload;

    procedure GetCurrentColumn;
    property BaseColumn: Integer read FBaseColumn;
    property CountColumn: Integer read FCountColumn;
    property Iterariton: Integer Read FIterariton;
    property CurrentIter: Integer read FCurrentIter;
    property IDColumn: Integer read FIDColumn;
  public
    procedure BeforeDestruction; override;

{$REGION ' ������ � ���������� ������� '}
    /// <summary> ���������� ������� ����� �������� ��� �������
    /// </summary>
    /// <param name="AIteration"> ����� �������� ������� � 1
    /// </param>
    /// <remarks> ��� ���������� ��� �� ������� ������������ ��������� ����� ������� � ������� �������� (��������� � ���������)
    /// </remarks>
    procedure SetCurrentIter(const AIteration: Integer);
    /// <summary> ������� ����� �������
    /// </summary>
    /// <remarks> ������� ��������� ������� � ��������� �� ��� �������
    /// </remarks>
    /// <returns> ����� ������� �������� (�����)
    /// </returns>
    function InsertColumn: Integer;
    /// <summary> �������� ����� ������� ��������
    /// </summary>
    /// <remarks> �������� ������ �������, � �� ���������.
    /// </remarks>
    /// <returns> ����� ������� ��������
    /// </returns>
    function GetIteration: Integer;
    /// <summary> �������� ����� ���-�� ��������
    /// </summary>
    /// <returns> ���-�� ��������
    /// </returns>
    function IteraritonCount: Integer;
{$ENDREGION}
  end;

  // �������� ������-��������� ��� ��������  ��������� ���� �������
  TFlexCelTemplate = class(TComponent)
  strict private
    FFLEX: TXLSFileReport;
    FOpenResult: Boolean;
    // �������� ������ ������ �������
    procedure GetSheetList;
    // �������� ����� �� ������
    procedure ClearReportGarbage;
  private
    FListBaseSheet: TObjectList<TSheetTemplate>;
    FListSheet: TObjectList<TSheetTemplate>;
  protected
    // �������������� ������� Excel ��� ����������� �������������
    class function GetPointToAddr(const PointCell: TBasePointCells): string; virtual;
    class procedure GetAddrToPoint(const AXlsAddr: string; out APointCell: TBasePointCells); inline;
  public
{$REGION ' ��������/�������� ������� '}
    /// <summary>�������� �������� ������� ������
    /// </summary>
    /// <param name="AOwner">������ Owner
    /// </param>
    /// <param name="AFileTemplate">������ ���� � ������� ����� Excel (������� ������� �����������)
    /// </param>
    /// <remarks>  ���� ����� ������� �� ������ ��� ���� �� �������� ����� ������������� ����������
    /// </remarks>
    constructor Create(AOwner: TComponent; const AFileTemplate: string); reintroduce;
    destructor Destroy; override;
{$ENDREGION}
{$REGION ' ������ � ������� ������� Excel '}
    /// <summary> ��������� ��������� ������ ���������� ������ ������� Excel
    /// </summary>
    /// <param name="AAdress"> ��������� �������� ������ ������ ������� Excel
    /// </param>
    /// <returns> TPoint ����� �������� � ������ (��������) �� 0
    /// </returns>
    class function AddrToPoint(const AAdress: string): TPoint; inline;
    /// <summary>  ��������� ��������� ������ ���������� ������ ������� Excel
    /// </summary>
    /// <param name="APoint"> TPoint ����� �������� � ������ (��������) �� 0
    /// </param>
    /// <returns> string ��������� �������� ������ ������ ������� Excel
    /// </returns>
    class function PointToAddr(const APoint: TPoint): string; inline;
{$ENDREGION}
{$REGION ' �������� ������ �� ������� ������� '}
    /// <summary> ��������� ������ �� ������ ����� �� �������
    /// </summary>
    /// <param name="AIndex"> ����� �����-�������, ������� � 1
    /// </param>
    /// <returns> TSheetTemplate ������� ������ �����-�������
    /// </returns>
    /// <remarks> Sheet ��������� ������������� ��� �������� �������
    /// ��� ��� ������������ ������ ������ ��� �� ��������� ������.
    /// <br> ��� ����� �������� ��� �������� ����� ����� ������� </br>
    /// </remarks>
    function GetSheetByIndex(const AIndex: Integer): TSheetTemplate;
    /// <summary> ��������� ������ �� ������ ����� �� �������� �����
    /// </summary>
    /// <param name="ABaseName">������������ �����-�������
    /// </param>
    /// <returns>TSheetTemplate ������� ������ �����-�������
    /// </returns>
    /// <remarks> Sheet ��������� ������������� ��� �������� �������
    /// ��� ��� ������������ ������ ������ ��� �� ��������� ������.
    /// <br> ��� ����� �������� ��� �������� ����� ���������� ����� �������!</br>
    /// </remarks>
    function GetSheetByName(const ABaseName: string): TSheetTemplate;
{$ENDREGION}
{$REGION ' ������� ���� ���������� }
    /// <summary> OpenResult ������� ������� �����
    /// </summary>
    /// <remarks> ����� �������� ������ � Excel, ������ � ������� ����� �����������, ��� ���
    /// <br> ��������-������� ����� �������, � ��������� ����� �������� ������� �� ������.</br>
    /// </remarks>
    procedure OpenResult;
{$ENDREGION}
    // �������� �������� ������� (��� ��������������� �����������)
    property FlexCel: TXLSFileReport read FFLEX;
  end;

  (* ��������� �� ������ ������ ��������� ������ ��� ������������� *)
  TCellSheet = class(TComponent)
  strict private
    FOwner: TSheetTemplate;
    FBasePos: TBasePointCells;
    FIncY: Integer;
    FDef: Variant;
    FIDCell: Integer;
    // ����������, ���������� � ������� �������
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
{$REGION ' ������ � �������� Excel ��� ������������ '}
    /// <summary> SetValue ���������� �������� ������
    /// </summary>
    /// <param name="AValue"> ��������������� ��������
    /// </param>
    procedure SetValue(const AValue: Variant);
    /// <summary> ���������� ���� ������� ������
    /// </summary>
    /// <param name="AColor"> ���� �������
    /// </param>
    procedure SetColor(const AColor: TColor);
    /// <summary> ���������� �����-��� ������ ������ (�� ��������� �� ��������� ������)
    /// </summary>
    /// <param name="ABarCode"> Int64 �����-���
    /// </param>
    /// <param name="AForeColor"> TColor ���� �����-����
    /// </param>
    procedure SetBarCodeEAN13Image(ABarCode: Int64; AForeColor: TColor; aDy, aDx: Integer);
    /// <summary> �������� ������� ����� ��� ���� ��������, ����� ��������� ������ ����� ������� �������� �����
    /// </summary>
    /// <param name="AInRow"> TRowTable ������ ��������� ������ �������
    /// </param>
    /// <param name="AInColumn"> ����� ������� ������� ������� � 1
    /// </param>
    /// <param name="ADefault"> ��������� �� ���������, ���� ������ �� ����� �� ����� �� ����� ��������
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer; const ADefault: Variant);
    /// <summary> �������� ��������� �� ������ �� ���� �,Y
    /// </summary>
    /// <returns> TPoint ������� ��������� ������ �� ������� sheet
    /// </returns>
    function GetCell: TPoint;
{$ENDREGION}
  end;

  (* ������ �������  ��������� ��������� ������ ��� ������������� *)
  TRowTable = class(TComponent)
  strict private
    FOwner: TTableSheet;
    FBaseRow: Integer;
    FListRow: TList<Integer>;
    FIDRow: Integer;
    // ����������, ���������� � ������� �������
    FTable: TTableSheet;
    FRow: TRowTable;
    FSheet: TSheetTemplate;
    // �������� ������� ������� � ��������
    procedure GetCurrentElement;
  private
    FCorrenTRow: Integer;

    // �������� ��������� object �� ��������� �������� object
    constructor Create(AOwner: TTableSheet; const ABaseRow: TRowTable); reintroduce; overload;

    // �������� ������� object, ��� ������� sheet � ��������� �������
    function CurrenTRow: Integer;
    // ��������� ������ ������� ��� ��������� �������� object
    function GetFormula(const AColumn: Integer; AListIterations: TIntegerList;
      out AValid: Boolean): string;
    // �������� �������� object, �� ��������� ���������
    constructor Create(AOwner: TTableSheet; const ABaseRow: Integer; const AIDRow: Integer);
      reintroduce; overload;

    property BaseRow: Integer read FBaseRow;
    procedure CurretBasePos(const AOfSet: Integer);
    property IDRow: Integer read FIDRow;
    property LisTRow: TList<Integer> read FLisTRow;
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;

{$REGION ' ���������� ����� ����� ������� '}
    /// <summary> ���������� �������� ������ ��� ������� ������� ������
    /// </summary>
    /// <param name="ANumColumn"> ����� �������
    /// (�������� ��������� ������� ������� � 1)
    /// </param>
    /// <param name="AValue"> ��������������� ��������
    /// </param>
    procedure SetValue(const ANumColumn: Integer; const AValue: Variant);
    /// <summary> �������� ������� ����� ��� ���� ��������, ����� ��������� ������ ����� ������� �������� �����
    /// ��� ������������ ������ ������� ������.
    /// </summary>
    /// <param name="AInRow"> TRowTable ������ ��������� ������ �������, �� ������� ������� �������
    /// </param>
    /// <param name="AInColumn"> ����� ������� ������, �� ������� ������� �������
    /// </param>
    /// <param name="ADefault"> ��������, �������� � ������ ���� ������ �� ����� �� ����� ��������
    /// </param>
    /// <param name="AResColumn"> ����� ������� ��� ������� ������ ���� ��������� �������
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer; const ADefault: Variant;
      const AResColumn: Integer = -1); overload;
    /// <summary> �������� ������� ����� ��� ����� ��������, ����� ��������� ������ ����� ������� �������� �����
    /// ��� ����������� ������ ������� ������.
    /// </summary>
    /// <param name="AInRow"> TRowTable ������ ��������� ������ �������, �� ������� ������� �������
    /// </param>
    /// <param name="AInColumn"> ����� ������� ������, �� ������� ������� �������
    /// </param>
    /// <param name="AListIterations">  TIntegerList ���� ��������� ��������, ��� ������ �� ������� ������� �������
    /// </param>
    /// <param name="ADefault"> �������� �������� � ������ ���� ������ �� ����� �� ����� ��������
    /// </param>
    /// <param name="AResColumn"> ����� ������� ��� ������� ������ ���� ��������� �������
    /// </param>
    procedure SetFormula(AInRow: TRowTable; const AInColumn: Integer;
      AListIterations: TIntegerList; const ADefault: Variant;
      const AResColumn: Integer = -1); overload;
    /// <summary> ���������� ������������ ����� ��� ������ (��������� ��� ��������� �������)
    /// </summary>
    /// <param name="AFonts"> ��������������� �����
    /// </param>
    /// <param name="AColumn"> ����� �������  (�������� ��������� ������� ������� � 1)
    /// </param>
    /// ����� ������� � ������
    procedure SetFont(const AFonts: TXLSFontData; const AColumn: Integer = 0);
    /// <summary> ���������� ������������ ���� ������ ��� ������ (��������� ��� ��������� �������)
    /// </summary>
    /// <param name="AColor"> ��������������� ���� ������
    /// </param>
    /// <param name="AColumn"> ����� ������� (�������� ��������� ������� ������� � 1)
    /// </param>
    procedure SetColorFont(const AColor: TColor; const AColumn: Integer); overload;
    /// <summary> ���������� ������������ ���� ������ ��� ���� ����� ������ ������� (��������� ��� ��������� �������)
    /// </summary>
    /// <param name="AColor"> ��������������� ���� ������
    /// </param>
    /// <remarks>  ���� ������ ����� ������ �� ������ �� ��������� ������� �������,
    /// �� ������� ������� ����� ����� �������
    /// </remarks>
    procedure SetColorFont(const AColor: TColor); overload;
    /// <summary> ���������� ������������ ���� ������� ������ (��������� ��� ��������� �������)
    /// </summary>
    /// <param name="ANumColumn"> ����� ������� (�������� ��������� ������� ������� � 1)
    /// </param>
    /// <param name="AColor"> ��������������� ����
    /// </param>
    procedure SetColor(const ANumColumn: Integer; const AColor: TColor); overload;
    /// <summary> ���������� ������������ ���� ������� ��� ���� ����� ������ ������� (��������� ��� ��������� �������)
    /// </summary>
    /// <param name="AColor"> ��������������� ���� ������
    /// </param>
    /// <remarks>  ���� ������� ����� ������ �� ������ �� ��������� ������� �������,
    /// �� ������� ������� ����� ����� �������
    /// </remarks>
    procedure SetColor(const AColor: TColor); overload;
{$ENDREGION}
{$REGION ' ������ �� �������� ������ '}
    /// <summary> �������/����������� ������ �� �������� ��������
    /// </summary>
    /// <remarks> ������ �������� ������ ���� ������ ���� �� ��� ���� �����������,
    /// ������ � �������� ������� exception
    /// </remarks>
    /// <returns> integer ������������ ������ ����� ����� ��������, � �� ����� ����� ������,
    /// ��� ������������� ��������� ������� ��� ������!
    /// </returns>
    function InsertRow: Integer;
    /// <summary> �������� ����� ������ ��� ������ �� ����������� ��������
    /// </summary>
    /// <param name="AIteration"> ����� ��������
    /// </param>
    /// <returns> integer ����� ������, ������� � 1
    /// </returns>
    function GetNumRow(const AIteration: Integer = -1): Integer;
    /// <summary> �������� ��-�� �������� ��� �������� �����
    /// </summary>
    /// <returns> integer ���-�� ��������
    /// </returns>
    function IterationCount: Integer;
    /// <summary> �������� ��-�� �������� ��� �������� �����
    /// </summary>
    /// <returns> integer ���-�� ��������
    /// </returns>
    function RowCount: Integer;
    /// <summary> �������� ������� ��������
    /// </summary>
    /// <returns> integer ����� ��������
    /// </returns>
    function GetIteration: Integer;
    /// <summary> ����������� ����� � ����� ������
    /// </summary>
    /// <param name="ABegColumn"> ����� ��������� �������
    /// </param>
    /// <param name="AEndColumn"> ����� �������� �������
    /// </param>
    /// <remarks> ������ ���������� ������� (�������� ��� ��������� ������� ������ ���������� � 1)
    /// </remarks>
    procedure SetColumnMerge(const ABegColumn: Integer = -1; const AEndColumn: Integer = -1);
    /// <summary> ����������� ����� � ������ � ���� � ��� ��������
    /// </summary>
    /// <param name="ABegColumn"> ����� ��������� �������
    /// </param>
    /// <param name="AEndColumn"> ����� �������� �������
    /// </param>
    /// <param name="AValue"> ��������������� ��������
    /// </param>
    /// <remarks> �������� ��� ��������� ������� ������ ����������� � 1
    /// </remarks>
    procedure SetValueAndMerge(const ABegColumn, AEndColumn: Integer; const AValue: Variant);
{$ENDREGION}
  end;

  (* ������� ��������� ������ ��� ������������� *)
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
    // �������� �������� �������
    FDropDef: Boolean;
    FDropTable: Boolean;

    // �������� ������� ������� �� ���������� �����������
    constructor Create(AOwner: TSheetTemplate; const ALeftPos, ARigthPos: TBasePointCells;
      const ANumTable: Integer = -1); reintroduce; overload;

{$REGION ' ������� �������������� ���������� �� ������� �������� � ������� �������� '}
    function ExtColumnCount: Integer;
    function GetFormula(const AColumn: Integer; AListIterations: TList<Integer>): string;
    // ���������� ����� � �������
    function IncRow(ABaseSheet: TSheetTemplate; ARow: TRowTable): Integer;
    Procedure IncColumn(ABaseSheet: TSheetTemplate; AColumn: TExtColumns);
    // ��������� ������ ������� � ������� �� �������� object (��������������� �� Row)
    procedure SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable; const ANumColumn: Integer;
      const AColor: TColor); overload;
    procedure SetColor(ABaseSheet: TSheetTemplate; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      const AColor: TColor); overload;
    // ��������� ������� �� �������� object (��������������� �� Row)
    procedure SetFont(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      AFonts: TXLSFontData);
    // ��������� �������� ����� �� �������� object (��������������� �� Row)
    procedure SetValue(ABaseSheet: TSheetTemplate; ARow: TRowTable; const ANumColumn: Integer;
      const AValue: Variant);
    // ��������� �������� object (��������������� �� Row)
    procedure SetFormula(ABaseSheet: TSheetTemplate; ARow: TRowTable; const AColumn: Integer;
      const AFormula: string);
    // ������������� �������� ������� � ������ ���������� �����
    procedure incOfSeTRow;
    // ���������� ������� ����� � �������
    procedure AddBaseRow(ABaseRow: TRowTable); overload;
    procedure AddBaseExtColumn(ABaseColumn: TExtColumns); overload;
    // ���������� ����� � �������
    procedure OverloadRowCol(ARecipient: TTableSheet);
    // ����������� �������������� ������
    function GetDesineRow(const ATypeRow: tBaseObj; const ABaseRow: Integer): Integer;
    // ����������� ��������� �����
    function GetDesineColumn(const ACoumn: Integer): Integer;
    // �������� ������� ���-�� ������� � �������
    function GetColumnCount: Integer;
    // ����������� ������� ��������� �������
    function GetDesineTable(const ATypePos: TCorrectionPos): Integer;
    // �������� �������� �������� �������
    function GeTRow(const AIDRow: Integer): TRowTable;
    // �������� ������� ������ ������������ �������
    function GetColumns(ABaseColun: TExtColumns): TExtColumns;
    // �������� ������� �����-��������� �������
    function GetBasePosition(const ATypePoint: TTypePoint): TBasePointCells;

    // ������� Column �� ������
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

    // ����� ������� �� ������� ��������
    property NumTable: Integer read FNumTable;
{$ENDREGION}
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;

{$REGION ' ���������������� ������� ������ � ��������� '}
    /// <summary> ������� ���������� ������ �� ������
    /// </summary>
    /// <param name="ARow"> ����� ������ ������ ������� (�� ����� ��������� �������)
    /// </param>
    /// <param name="AEndNumRow"> ����� ��������� ������ �������, �� ����� ����� �������������
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br> 1. ����������� ���� ������ ���� ������� �����������.  </br>
    /// <br> 2. ������� �� ��������, � ������ ������� ������ �������, ��� �� ������� ������� </br>
    /// ������ � TRowTable, �������������� �������� GetNumRow � TRowTable
    /// </remarks>
    procedure SetGroupRows(ARow: TRowTable; const AEndNumRow: Integer);
    /// <summary> ������� ������� ����� �� ����������� ������ ����������� ������
    /// (��������� �� �� �������� � �� ������������������)
    /// </summary>
    /// <param name="ARow"> ����� ������ �������
    /// </param>
    /// <param name="ANumColumn"> ����� ������� ������� (��������� ������� ������ � 1)
    /// </param>
    /// <returns> TPoint ��������� �� ���� x,y
    /// </returns>
    /// <remarks>
    /// <red> ������ ���������: </red> ������� �� ��������, � ������ ������� ������ �������, ��� �� ������� �������
    /// ������ � TRowTable �������������� �������� GetNumRow � TRowTable.
    /// ������� ��� �� Sheet ��� ������� ������� ������ �������.
    /// </remarks>
    function GetCell(const ANumRow: Integer; const ANumColumn: Integer = -1): TPoint; overload;
    /// <summary> �������� ���� ������� (������ ������ �� ����� �������)
    /// <param name="ATypePos"> ��� ����
    /// <br> ttpLeft= ����� ������� ���� ttpRight= ������ ������ ����</br>
    /// </param>
    /// <returns><b>Returns: TPoint</b> ��������� �� ���� x,y
    /// </returns>
    /// </summary>
    /// <remarks>
    /// ������ �������������� � ��� TExtColumns, � ���� ������� TRowTable,
    /// ����� ������� ��������� ����� ������ ����� �������.
    /// </remarks>
    function GetCell(const ATypePos: TTypePoint): TPoint; overload;
    /// <summary> ���������������� ������ ������� � �������
    /// </summary>
    /// <param name="ANumRow"> ����� ������ �� �������� �������� ���� �������
    /// </param>
    /// <returns> TRowTable C����� �� ������� ������ ��� �������
    /// </returns>
    /// <remarks><red> ������ ���������: </red>
    /// <br> 1. C������� ������ ������ ������� ���������� � ��� ������������� �������� �������.</br>
    /// <br> 2. � ������ ���� ������� ����� ����� ����� ��������� ������ �� ��� ���� ������ �����������,
    /// �������� Sheet ������� ������ ���� � ����������� ������. </br>
    /// <br> 3. ������� ������������ object ������. </br>
    /// <br> 4. �������� ������� ����� �� ������� ������������ ���� �������.</br>
    /// </remarks>
    function AddBaseRow(const ANumRow: Integer): TRowTable; overload;
    /// <summary> ���������������� ������� ������� � �������
    /// </summary>
    /// <param name="ABegColumn"> ����� ������ ������������ �������
    /// </param>
    /// <param name="AACountColumn"> ����� ������� ������������ �������(��� ����������� �������)
    /// </param>
    /// <returns> TExtColumns c����� �� ������� ������������ ������� ��� �������
    /// </returns>
    /// <remarks> <red> ������ ���������: </red>
    /// <br> 1. ��������� ������� �������� ���������� ��� ������ ���-�� �����������(��� ���� �� ��� ���� ������� ��������)
    /// � ��������� �� ������������ ��������� �������������� ����� ��������� ������ �������� ��� TExtColumns (�.SetCurrentIter)</br>
    /// <br> 2. � ������ ���� ������� ����� ����� ����� ��������� ������, �� ��� ���� ������ �����������,
    /// �������� Sheet ������ ��� TExtColumns. </br>
    /// <br> 3. ������� ������������ object ������ </br>
    /// <br> 4. �������� ������� TExtColumns �� ������� ������������ ���� ������� </br>
    /// </remarks>
    function AddBaseExtColumn(const ABegColumn: Integer; const ACountColumn: Integer = 1)
      : TExtColumns; overload;
    /// <summary> ��������������� �������������� �������
    /// <param name="AView"> ��� ���� ��������
    /// tcpLeftY= ����� ������� ���� �� ��� Y, tcpLeftX= ����� ������� ���� �� ��� X,
    /// tcpRightY= ������ ������ ���� �� ��� Y,  tcpRightX= ������ ������ ���� �� ��� X
    /// </param>
    /// <param name="AValue"> �������� �������� ��������
    /// </param>
    /// </summary>
    procedure CorretionBasePos(const AView: TCorrectionPos; const AValue: Integer);
    /// <summary> ��������� ��������� ����� �������
    /// <param name="AColumn"> ����� ������� � ������� (������� � 1), ���������� �� ������� ������� �� ����
    /// </param>
    /// <param name="AValue"> �������� ������ � �������
    /// </param>
    /// <param name="ARowTitle"> ���� ����� � ��������� ����� ����� 1, �� ����� ������� ������ �� �������
    /// </param>
    /// </summary>
    procedure SetTitleColumns(const AColumn: Integer; const AValue: Variant;
      const ARowTitle: Integer = 0);
    /// <summary> ��������� ���� �� ������� � ������ ����������
    /// <param name="ARow"> TRowTable ��� �������� �������� ����� ������� �������� (��� ������� �������)
    /// </param>
    /// <param name="AColumn">����� ������� � ������� (������� � 1)
    /// </param>
    /// <param name="ADefault">�������� �� ���������, ���� �� ������� ����� TRowTable ������ �� ����� ��������
    /// </param>
    /// <param name="RowResult">���� �������� ����� ����� ����� �� ����� ������� �����(������ �� ��������� ���������� TRowTable)
    /// </param>
    /// </summary>
    procedure SetSumTable(ARow: TRowTable; const AColumn: Integer; const ADefault: Variant;
      const RowResult: Integer = 1);
    /// <summary> ������� �������, ��� ������� ��� ��� �� ����� ������, �������� ������ �������� �������
    /// <param name="OffSetTName"> ������� ����� � ���� ����� ������� �� �������� ���� (������ ��� ������������ �������)
    /// </param>
    /// </summary>
    /// <remarks><red> ������ ���������: </red>
    /// <br>1. ������ �������� ��������������� �� ��� ���������� ������� �������, ������� �������� ������������� ��� ������ ����� ��������.</br>
    /// <br>2. �������� ������� ������� ��������� �� ����������� ���-�� ����� (��� ������ �������� ����� � �� ��������)</br>
    /// </remarks>
    procedure DropDef(const OffSetTName: Integer = 0);
    /// <summary> ������� ������� � ����� ������ �� ������� Sheet
    /// <param name="OffSetTName"> ������� ����� � ���� ����� ������� �� �������� ���� (������ ��� ������������ �������)
    /// </param>
    /// </summary>
    /// <remarks><red> ������ ���������: </red>
    /// <br>1. �������� ������� ������� ��������� �� ����������� ���-�� ����� (��� ������ �������� ����� � �� ��������).</br>
    /// <br>2. ������� �� ������� Sheet ����� ������� � ����� ������ (���� ���� � ��� ���� ��������� ������)</br>
    /// </remarks>
    procedure DropTable(const OffSetTName: Integer = 0);
    /// <summary> ������� ������� ������ �������
    /// <param name="ABegNumRow"> ��������� ������ ����������
    /// </param>
    /// <param name="AEndNumRow"> ������� ������ ����������
    /// </param>
    /// <param name="ABeginCol"> ��������� ������� ����������
    /// </param>
    /// <param name="AEndCol"> ������� ������� ����������
    /// </param>
    /// </summary>
    /// <remarks>
    /// <red> ������ ���������: </red>
    /// <br> 1. ����� ��������� ������ ������ �������, � �� ������ �������� �����.</br>
    /// <br> 2. ��������� ������� ���������� � 1 </br>
    /// <br> 3. �������� ��������� ��������� � �������� ������� ����������, ��� ������������ �������� �������� �����(��������� ������� ����� �������)</br>
    /// </remarks>
    procedure SetMergeGroupRow(const ABegNumRow, AEndNumRow: Integer;
      const ABeginCol, AEndCol: Integer);
{$ENDREGION}
  end;

  (* Sheet ��������� ������ ��� ������������� � ��� ������� �����-������� ������ ���� ���� *)
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
    // �������� ���������� ������� ��� ������ � ��������� �� ������� �����������
    function AddBaseTable(const ALeftPos, ARigthPos: TBasePointCells; const ANumTable: Integer = -1)
      : TTableSheet; overload;
    // �������� ���������� ������ �� �����������
    function AddBaseCells(ABaseSheet: TSheetTemplate; const AFBasePos: TBasePointCells;
      const ADefault: Variant; const ASheetType: TTypeSheet; const AIDCell: Integer)
      : TCellSheet; overload;
  protected
    // ������ ������� ������� �� ��������� ��������������� ������� �������
    procedure ReassignSheet(const ANum: Integer); virtual;
  private
    // ������������ � ����������� ������ � �������
    constructor Create(AOwner: TFlexCelTemplate; const AIndex: Integer; const ABaseName: string;
      const ATypeSheet: TTypeSheet = ttsNoBase); reintroduce;

    // �������� Sheet �� ������� ���������
    procedure DeleteBaseRow(ATable: TTableSheet; ARow: TRowTable);
    procedure DeleteBaseColumns(ATable: TTableSheet; AColumns: TExtColumns);
    // ��������� ������� ��������� �� ������� �� ���������
    function GetPosConst(const AConst: string): TBasePointCells;
    function IncRow(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable): Integer;
    procedure IncColumn(ABaseSheet: TSheetTemplate; ATable: TTableSheet; AColumn: TExtColumns);
    // ��������� ������ � ������
    procedure SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ANumColumn: Integer; const AColor: TColor); overload;
    // ��������� ������
    procedure SetColor(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColor: TColor); overload;
    // ������� ���� ������
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColor: TColor); overload;
    procedure SetColorFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; const AColor: TColor); overload;
    procedure SetFont(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; AFonts: TXLSFontData);
    // ������� ������ � ������� �� ��������
    procedure SetColumnMerge(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ABegColumn, AEndColumn: Integer);
    // ��������� �������� ������
    procedure SetValue(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const ANumColumn: Integer; const AValue: Variant); overload;
    // ��������� �������� ������ � ���������� TCellSheet
    procedure SetValue(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AValue: Variant); overload;
    // ��������� �������� ������ �� ������
    procedure SetValue(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
      const AValue: Variant); overload;
    // ��������� �������� ����� ������ � ���������� TCellSheet
    procedure SetColor(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AColor: TColor); overload;
    // ���������� ��
    procedure SetBarCodeEAN13Image(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const ABarCode: Int64; const AForeColor: TColor; const aDy, aDx: Integer);
    // ������ ������
    procedure SetFormula(ABaseSheet: TSheetTemplate; ATable: TTableSheet; ARow: TRowTable;
      const AColumn: Integer; const AFormula: string); overload;
    procedure SetFormula(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
      const AFormula: string); overload;
    procedure SetFormula(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
      const AFormula: String); overload;
    // ��������� �������� �� ���������
    procedure SetValue(const AFBasePos: TBasePointCells; const AValue: Variant;
      const ATypeSheet: TTypeSheet); overload;

    procedure DropTable(ATable: TTableSheet);
    function GetTable(const ANumTable: Integer): TTableSheet;
    function GetCell(const AIDCell: Integer): TCellSheet;
    // �������� ������� ������������� �������������
    procedure CleaningSheet;
    // ��������� �������� Sheet (����������� �����)
    procedure SetSheet(ASheet: TSheetTemplate);
    // ���������� �������� � �������� Sheet
    procedure AddBaseCells(ABaseSheet: TSheetTemplate; ACells: TCellSheet); overload;

    property FlexCelTemplate: TFlexCelTemplate read FOwner;
    // ��� �������
    property SheetType: TTypeSheet read FSheetType;
    // ������� ��������
    property Sheet: TSheetTemplate read FSheetCurrent;
    // ������� ����� �������
    property SheetNum: Integer read FSheetNum;
    // ������ �� ������� �������
    property BaseSheet: Integer read FBaseSheet;
    property BaseName: string read FBaseName;

  protected
    // ��������� �������� �� ���������
    procedure SetPointValue(const APoint: TPoint; const AValue: Variant);
  public
    destructor Destroy; override;
    procedure BeforeDestruction; override;
{$REGION ' ������ ������ ��������� ��� ���������� '}
    /// <summary> ���������� ������� TCellSheet � ������� �� ���������/T�g (�������������)
    /// </summary>
    /// <param name="ATag"> �������� Tag(���������)
    /// </param>
    /// <param name="ADefault"> �������� ������ �� ���������
    /// </param>
    /// <returns>  TCellSheet ������������ ������� TCellSheet
    /// </returns>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ��������� ������ ��� ������� Sheet.  </br>
    /// <br>2. ��� �� ��������� ����� ����� ADefault </br>
    /// <br>3. object ������� � ������� TSheetTemplate � �� ��������� ������ � ��������</br>
    /// </remarks>
    function AddBaseCells(const ATag: string; const ADefault: Variant): TCellSheet; overload;
    /// <summary> ���������� ������� TCellSheet �� ���������� ������ Excel (�������������)
    /// </summary>
    /// <param name="AXlsAddr"> ����� ������ ������� Excel
    /// </param>
    /// <param name="ADefault"> �������� ������ �� ���������
    /// </param>
    /// <returns> TCellSheet ������������ ������� TCellSheet
    /// </returns>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ��������� ������ ��� ������� Sheet.  </br>
    /// <br>2. ��� �� ��������� ����� ����� ADefault </br>
    /// <br>3. object ������� � ������� TSheetTemplate � �� ��������� ������ � ��������</br>
    /// </remarks>
    function AddBaseCellsByAddr(const AXlsAddr: string; const ADefault: Variant)
      : TCellSheet; overload;
    /// <summary> ���������� ����� ������� �� ��������� �������
    /// </summary>
    /// <param name="ASheetName"> ��������� ��� ����� �������
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ���� ��� Sheet ��������� �������? �� ������� ����� ������, ���� �������� �� ����� ������ � ��������� #x ��� x ����� ��������� �����.  </br>
    /// <br>2. ��� ���������� ��� ���������� Sheet (TCellSheet,TTableSheet...) ��������� � �������� ��������� </br>
    /// <br>3. ��� �������� �������� ����������� ����� �������� � 0</br>
    /// </remarks>
    procedure CopyAddSheet(const ASheetName: string);
    /// <summary> ���������� ������� �� ������� ������� �� �����������(�������������)
    /// </summary>
    /// <param name="ALeftRow"> ����� ������ �� ��� Y ��� �������� ������ ����
    /// </param>
    /// <param name="ALeftCol"> ����� ������ �� o�� X ��� �������� ������ ����
    /// </param>
    /// <param name="ARigthRow"> ����� ������ �� ��� Y ��� ������� ������� ����
    /// </param>
    /// <param name="ARigthCol"> ����� ������ �� o�� X ��� ������� ������� ����
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ���������, ������ ��� ������� Sheet.</br>
    /// <br>2. ������������� ����� �������, �������� � ���� ����������.</br>
    /// <br>3. object ������� � ������� TSheetTemplate, � �� ��������� ������ � ��������.</br>
    /// </remarks>
    function AddBaseTable(const ALeftRow, ALeftCol, ARigthRow, ARigthCol: Integer)
      : TTableSheet; overload;
    /// <summary> ���������� ������� �� ������� ������� �� ���� ����������/tag (�������������)
    /// </summary>
    /// <param name="ALeftConst">���������/�������� Tag, �������� ������ ����
    /// </param>
    /// <param name="ARigthConst">���������/�������� Tag, ������� ������� ����
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ��������� ������ ��� ������� Sheet.</br>
    /// <br>2. ������������� ����� ������� ��������� � ���� ����������.</br>
    /// <br>3. object ������� � ������� TSheetTemplate, � �� ��������� ������ � ��������.</br>
    /// </remarks>
    function AddBaseTable(const ALeftConst, ARigthConst: string): TTableSheet; overload;
    /// <summary> ���������� ������� �� ������� ������� �� ���������/tag � ��������(�������������)
    /// </summary>
    /// <param name="ALeftConst">���������/�������� Tag �������� ������ ����
    /// </param>
    /// <param name="ARigthOffset">������ �������� ��� ������� ������� ���� �� ���� X,Y
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ��������� ������ ��� ������� Sheet.</br>
    /// <br>2. ������������� ����� �������, �������� � ���� ����������.</br>
    /// <br>3. object ������� � ������� TSheetTemplate, � �� ��������� ������ � ��������.</br>
    /// </remarks>
    function AddBaseTable(const ALeftConst: string; ARigthOffset: TPoint): TTableSheet; overload;
    /// <summary> ���������� ������� �� ������� ������� �� ����������� ������� Excel(�������������)
    /// </summary>
    /// <param name="ALeftConst">���������� �������� ������ ����
    /// </param>
    /// <param name="ARigthConst">���������� ������� ������� ����
    /// </param>
    /// <remarks> <red> ������ ���������: </red>
    /// <br>1. ��������� ��������� ������ ��� ������� Sheet.</br>
    /// <br>2. ������������� ����� ������� �������� � ���� ����������.</br>
    /// <br>3. object ������� � ������� TSheetTemplate, � �� ��������� ������ � ��������.</br>
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
    raise EASDException.Create('�� ������ ���� ������� ������ Excel ' + sLineBreak + '�' +
      AFileTemplate + '�');

  try
    FFLEX := TXLSFileReport.Create(AOwner, AFileTemplate);
  except
    on E: Exception do
      raise EASDException.Create('������ ��� �������� ������� ������: ' + sLineBreak + '�' +
        E.Message + '�' + sLineBreak + '�������� �� ������ ������ �������������');
  end;

  // ������� �������� ������� ������ ���������� ��������� ��� ����� �������
  FListBaseSheet := TObjectList<TSheetTemplate>.Create(TComparer<TSheetTemplate>.Construct(
    function(const Left, Right: TSheetTemplate): Integer
    begin
      Result := CompareValue(Right.BaseSheet, Left.BaseSheet);
    end), True);

  // ������� �������� ������� ������ ���������� ��������� ��� ����� �������
  FListSheet := TObjectList<TSheetTemplate>.Create(True);

  // �������� ����� ������� ��������
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
      ('TFlexCelTemplate.GetSheetIndex ��������� � ��������������� ����� �������');

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
      ('TFlexCelTemplate.GetSheetName ��������� � ��������������� ����� �������');
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
      ('TFlexCelTemplate.OpenResult ��������� �������� ����������, �����������');

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

  // ��� ����� �������� ������ � 0 � ����������� � 1 �� ��� X
  Result.X := Result.X - 1;
  Result.Y := StrToIntDef(Addr, -1);

  Assert((Result.Y > -1) or (Result.Y > -1),
    'TFlexCelTemplate.AddrToPoint ������ ������� ��������� ������ Excel');
end;
{$ENDREGION}
{$REGION ' TSheetTemplate ....}

// ��������� ������� �� ������ ��������
function TSheetTemplate.AddBaseTable(const ALeftConst, ARigthConst: string): TTableSheet;
begin
  Result := AddBaseTable(GetPosConst(ALeftConst), GetPosConst(ARigthConst));
end;

// ��������� ������� �� �����������
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
      ('TSheetTemplate.BeforeDestruction ������� ������� ������ ��������� ������ ����������');
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

  // ������� ��������� �� ����������� ������� ������
  for tmpTb in Self.FTable do
  begin
    RlsTb := tsh.AddBaseTable(tmpTb.GetBasePosition(ttpLeft), tmpTb.GetBasePosition(ttpRight),
      tmpTb.NumTable);
    RlsTb.Name := 'Sheet' + IntToStr(tsh.FSheetNum) + '_' + 'CurrentTable' +
      IntToStr(tmpTb.NumTable);
    RlsTb.FDropDef := tmpTb.FDropDef;
    tmpTb.OverloadRowCol(RlsTb);
  end;

  // ������� ��������� �� ����������� ������
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

  // ������ ������ �������� � ������� ����� ����������� ��� ����������� ���������
  FTable := TObjectList<TTableSheet>.Create(TComparer<TTableSheet>.Construct(
    function(const Left, Right: TTableSheet): Integer
    begin
      Result := CompareValue(Right.Position.Y, Left.Position.Y);
    end), True);

  // ������ ���������� ����� �������
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

  // ���������� ���������� �������
  Left := ATable.GetBasePosition(ttpLeft);
  Right := ATable.GetBasePosition(ttpRight);

  // ����������� ������� � ������ ����
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
    raise EASDException.Create('TSheetTemplate.GetCell �� ������ ������� ���������');
end;

function TSheetTemplate.GetPosConst(const AConst: string): TBasePointCells;
var
  sh: Integer;
begin
  if Assigned(FSheetCurrent) then
    raise EASDException.Create
      ('TSheetTemplate.GetPosConst ������������� ���� �������� �������� ������ � ������� � ������ �� ������ �������');

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
        ('������ TSheetTemplate.GetPosConst �� ������� ��������� �� ����� Excel ' + AConst);

    // �������� �������� ���������
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
    raise EASDException.Create('TSheetTemplate.GetTable � ������ ������ �� ������ ������� �������');
end;

procedure TSheetTemplate.IncColumn(ABaseSheet: TSheetTemplate; ATable: TTableSheet;
AColumn: TExtColumns);
var
  SrcRange: TXlsCellRange;
  Left, Right: TBasePointCells;
  rb, re, cb: Integer;
begin
  SetSheet(Self);

  // ���������� ���������� �������
  Left := ATable.GetBasePosition(ttpLeft);
  Right := ATable.GetBasePosition(ttpRight);

  // ����������� ������� � ������ ����
  rb := Left.FBasePoint.Y + ATable.FOffSet[0];
  re := Left.FBasePoint.Y + ATable.FOffSet[0] + Right.FBasePoint.Y + ATable.FRowHistory.Count + 1;
  cb := ATable.GetDesineColumn(AColumn.BaseColumn - 1) + 1;

  SrcRange := TXlsCellRange.Create(rb, cb, re, (cb + (AColumn.CountColumn)) - 1);

  with ABaseSheet.FlexCelTemplate.FlexCel do
  begin
    XF.InsertAndCopyRange(SrcRange,
    // ��������� ��������
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

      // ����� ����� �������� ����������� ������
      ABaseSheet.FlexCelTemplate.FlexCel.InsertAndCopyRow(tmpTable.GetDesineRow(tboBase,
        ARow.BaseRow), tmpTable.LastRow);

      // ����������� ������� ������� ������
      inc(tmpTable.FLasTRow);
      Result := tmpTable.FRowHistory.Count;
      ARow.FCorrenTRow := Result;

      Break;
    end
    else
      tmpTable.incOfSeTRow;
  end;

  if not Assigned(tmpTable) then
    raise EASDException.Create('TSheetTemplate.IncRow �� ������� ����������� �������')
  else
  begin
    // K������� �������� ��� ��������� �� �������� ����� ���������
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
      // ����� ��������� ��������
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
      // ����� ������ ���� ������
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
      // ����� ������ ���� ������
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
  // ��������� ��������
  ABaseSheet.FlexCelTemplate.FlexCel.InsertEAN13Image(ABarCode, AForeColor,
    ACell.Position.FBasePoint.Y, ACell.Position.FBasePoint.X, aDy, aDx);
end;

procedure TSheetTemplate.SetColor(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const AColor: TColor);
begin
  SetSheet(Self);
  // ����� ��������� ��������
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
      // ����� ������ ���� ������
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
      // ����� ������ ���� ������
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
  // ����� ��������� ��������
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
      // ����� ������ ���� ������
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
      // ����� ��������� ��������
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
  // ����� ��������� ��������
  ABaseSheet.FlexCelTemplate.FlexCel.SetFormula(ARow, AColumn, AFormula);
end;

procedure TSheetTemplate.SetPointValue(const APoint: TPoint; const AValue: Variant);
begin
  with Self.FlexCelTemplate.FlexCel do
  begin
    SetSheet(Self);
    // ����� ��������� ��������
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
  // ����� ��������� ��������
  ABaseSheet.FlexCelTemplate.FlexCel.SetValue(ACell.Position.FBasePoint.Y,
    ACell.Position.FBasePoint.X, AValue);
end;

procedure TSheetTemplate.SetFormula(ABaseSheet: TSheetTemplate; ACell: TCellSheet;
const AFormula: String);
begin
  SetSheet(Self);
  // ����� ��������� ��������
  ABaseSheet.FlexCelTemplate.FlexCel.SetFormula(ACell.Position.FBasePoint.Y,
    ACell.Position.FBasePoint.X, AFormula);
end;

procedure TSheetTemplate.SetValue(ABaseSheet: TSheetTemplate; const ARow, AColumn: Integer;
const AValue: Variant);
begin
  SetSheet(Self);
  // ����� ��������� ��������
  ABaseSheet.FlexCelTemplate.FlexCel.SetValue(ARow, AColumn, AValue);
end;

procedure TSheetTemplate.ReassignSheet(const ANum: Integer);
begin
  if ANum = 0 then
    raise EASDException.Create('TSheetTemplate.ReassignSheet ������� ������ �� �������� ���������');

  if not Assigned(Self.FSheetCurrent) then
    raise EASDException.Create
      ('TSheetTemplate.ReassignSheet ���������� �������� ������ ��� ����������� ������');

  if (FListOwner.Count = 0) or (FListOwner.Count < ANum) then
    raise EASDException.Create('TSheetTemplate.ReassignSheet ��������� � ��������������� �����');

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
    // ������������� �������� �� ���������
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
      ('TTableSheet.AddBaseExtColumn ������� �������� ������ ������� � ������� ��������');

  Result := TmpCol;
end;

function TTableSheet.AddBaseRow(const ANumRow: Integer): TRowTable;
begin
  if not(Self.FOwner.SheetNum = 0) then
    raise EASDException.Create
      ('TTableSheet.AddBaseRow ���������� ����� ������ �������� �� ������� ���������� ������');

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
      ('TTableSheet.BeforeDestruction ������� ������� ������ ��������� ������ ����������');
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
          // ��� ������������ �������� ���� ������� �� ����������� ������������ ������ �����
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
    raise EASDException.Create('TTableSheet.Create ���������� ������� ����� ������� �� �����');

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

  // ���� ������� ��� ������ ������ � ������� ���������� ��� �������� �������� � ������� �������� ��
  FRowHistory := TList<Integer>.Create;

  // ������� ������ ����� �������
  FRow := TObjectList<TRowTable>.Create(TComparer<TRowTable>.Construct(
    function(const Left, Right: TRowTable): Integer
    begin
      Result := CompareValue(Right.BaseRow, Left.BaseRow);
    end), True);

  // ���������  �������
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
    raise EASDException.Create('TTableSheet.GeTRow � ������� ������� �� ������� ������');

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
        ('TTableSheet.GetDesineColumn ������� ������� �� �������� ���������');

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
    raise EASDException.Create('TTableSheet.GetDisineRow �� ������ ������ ����� �������');

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
  // ��������� �������� �� ���� ����������������
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
    raise EASDException.Create('TTableSheet.GeTRow � ������� ������� ��������� ������');
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
      ('TTableSheet.GetPoinTRow ��������� � �������������� ������ �������');

  nc := ifthen(ANumColumn < 1, 1, ANumColumn);
  Result.Y := tmpTable.Position.Y + tmpTable.FRow[0].BaseRow + ANumRow + 1;
  Result.X := tmpTable.GetDesineColumn(nc) - 1;

  // ����������� FlexCel �� ������ �������
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
    raise EASDException.Create('TTableSheet.GetRealRow ������ �� �������������� ������');

  if (AIteration < 1) then
    Result := ARow.LisTRow[ARow.LisTRow.Count - 1]
  else
    Result := ARow.LisTRow[AIteration - 1];
end;

function TTableSheet.GetCurrentSheet: TSheetTemplate;
begin
  if not Assigned(SheetTemplate.Sheet) then
    raise EASDException.Create('TTableSheet.SetSumTable ��������� � �������������� �������');

  Result := SheetTemplate.Sheet;
end;

function TTableSheet.GetCurrentTable(ASheet: TSheetTemplate): TTableSheet;
var
  tmpTable: TTableSheet;
begin
  tmpTable := ASheet.GetTable(Self.FNumTable);
  if not Assigned(tmpTable) then
    raise EASDException.Create('TTableSheet.SetSumTable ��������� � �������������� �������');

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

  // ���� ��� �� ������ ��������
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
    raise EASDException.Create('TTableSheet.GetDisineRow �� ������ ������ ����� �������');

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
    raise EASDException.Create('TFlexCell.Create ������ �������� ��������� �� ������� ������');

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
      ('TCellSheet.BeforeDestruction ������� ������� ������ ��������� ������ ����������');
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
    raise EASDException.Create('TTableSheet.GetPoinTRow ��������� � �������������� ������e');

  Result.Y := FCell.Position.FBasePoint.Y + 1;
  Result.X := FCell.Position.FBasePoint.X;

  // ����������� FlexCel �� ������ �������
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
      ('TRowTable.BeforeDestruction ������� ������� ������ ��������� ������ ����������');
end;

constructor TRowTable.Create(AOwner: TTableSheet; const ABaseRow: Integer;
const AIDRow: Integer);
begin
  if not(AOwner.InheritsFrom(TTableSheet)) then
    raise EASDException.Create('TRowTable.Create ������ �������� ��������� �� ������� ������');

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
    raise EASDException.Create('TRowTable.GetCurrenTRow ����� �������������� ������');

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
    raise EASDException.Create('TRowTable.GetCurrentElement ��������� � ��������������� �����');

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
    raise EASDException.Create('TRowTable.InserTRow ������ ���������� ������');

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
      ('TExtColumns.BeforeDestruction ������� ������� ������ ��������� ������ ����������');
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
    raise EASDException.Create('TRowTable.GetCurrentElement ��������� � ��������������� �����');

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
    raise EASDException.Create('TExtColumns.SetCurrentIter ��������� � ������� �������');

  GetCurrentColumn;

  if FColumn.FIterariton < AIteration then
    raise EASDException.Create('TExtColumns.SetCurrentIter ��������� � �������������� �������');

  FColumn.FCurrentIter := AIteration;
end;
{$ENDREGION}

end.
